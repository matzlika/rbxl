module Rbxl
  # Read-modify-save workbook for surgical edits to an existing +.xlsx+.
  #
  # The design promise mirrors +rbpptx+: <em>what we don't understand, we
  # don't touch</em>. The package is opened as a ZIP, each part you mutate is
  # re-serialized, and every other entry — styles, drawings, charts, comments,
  # pivot caches, custom XML, untouched worksheets — round-trips byte-for-byte
  # via {Zip::Entry#copy_raw_entry}. Inside a worksheet you do edit, only the
  # specific +<c>+ element you target is rewritten; surrounding cells, the
  # row's other attributes, +<mergeCells>+, +<conditionalFormatting>+,
  # +<dataValidations>+, and any unknown OOXML extensions remain in place.
  # The cell's existing +s+ (style index) attribute is preserved, so template
  # number formats, fonts, and fills carry through to the new value.
  #
  # The editable mode is the right tool for template-style fill-ins: open a
  # template with named cells, write a handful of values, save back. It is
  # explicitly <em>not</em> the right tool for rewriting the data area of a
  # large worksheet — the touched sheet is parsed as a Nokogiri DOM, so peak
  # memory scales with that sheet's on-disk size. Use the write-only mode
  # (+Rbxl.new+) for that case instead.
  #
  # == Out of scope (1.4.0)
  #
  # * inserting / deleting / reordering / duplicating sheets
  # * editing styles, formulas, named ranges, drawings, or shared strings
  # * +Date+ / +Time+ / +DateTime+ values (raise {EditableCellTypeError};
  #   convert to a numeric serial yourself if you need a date cell)
  # * recomputing the worksheet +<dimension>+ when a write expands the bounds
  #
  # == Strings on write
  #
  # Cells written through this mode become inline strings
  # (+t="inlineStr"+), so +xl/sharedStrings.xml+ is never mutated. Existing
  # +t="s"+ cells you don't touch keep resolving through the SST as usual;
  # only cells you actually overwrite drop their SST reference.
  class EditableWorkbook
    # Namespace for the main SpreadsheetML schema.
    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".freeze

    # Namespace used for document-level relationships.
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships".freeze

    # Namespace used by the OPC package relationships layer.
    PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships".freeze

    # Relationship type identifying the workbook part inside +_rels/.rels+.
    OFFICE_DOC_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".freeze

    OLE_CFB_MAGIC = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b.freeze
    private_constant :OLE_CFB_MAGIC

    ZIP_LOCAL_MAGIC = "PK\x03\x04".b.freeze
    private_constant :ZIP_LOCAL_MAGIC

    # @return [String] filesystem path the workbook was opened from
    attr_reader :path

    # @return [Array<String>] visible sheet names in workbook order
    attr_reader :sheet_names

    # Convenience constructor equivalent to +new(path)+. When a block is
    # given, the workbook is yielded and {#close} is called automatically
    # when the block returns or raises.
    #
    # @param path [String, #to_path]
    # @yieldparam book [Rbxl::EditableWorkbook]
    # @return [Rbxl::EditableWorkbook, Object] the workbook when no block is
    #   given, otherwise the block's return value
    def self.open(path)
      book = new(path)
      return book unless block_given?

      begin
        yield book
      ensure
        book.close
      end
    end

    # Opens the package, validates the format, and indexes worksheet parts
    # by visible sheet name. Worksheet XML is not parsed until the caller
    # touches that sheet via {#sheet}.
    #
    # @param path [String, #to_path] path to the +.xlsx+ file
    # @raise [Rbxl::UnsupportedFormatError] if the file is not a valid
    #   +.xlsx+ container (e.g. a legacy +.xls+, or non-ZIP bytes)
    # @raise [Rbxl::WorkbookFormatError] if +xl/workbook.xml+ or its rels are
    #   missing, malformed, or internally inconsistent
    def initialize(path)
      @path = path.to_s
      ensure_xlsx_format!(@path)
      @zip = Zip::File.open(@path)
      @closed = false
      @workbook_part = locate_workbook_part
      @workbook_dir = File.dirname(@workbook_part)
      @sheet_entries = load_sheet_entries
      @sheet_names = @sheet_entries.keys.freeze
      @shared_strings = nil
      @sheets_by_name = {}
    end

    # Returns the editable worksheet for +name_or_index+. Repeated calls for
    # the same sheet return the same in-memory object so edits accumulate
    # across calls before {#save}.
    #
    # @param name_or_index [String, Integer] visible sheet name as listed in
    #   {#sheet_names}, or an integer index (negatives count from the end)
    # @return [Rbxl::EditableWorksheet]
    # @raise [Rbxl::SheetNotFoundError] if +name_or_index+ does not resolve
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def sheet(name_or_index)
      ensure_open!

      name = resolve_sheet_name(name_or_index)
      @sheets_by_name[name] ||= EditableWorksheet.new(
        zip: @zip,
        entry_path: @sheet_entries.fetch(name) {
          raise SheetNotFoundError, "sheet not found: #{name}"
        },
        workbook_path: @path,
        shared_strings: shared_strings,
        name: name
      )
    end

    # Iterates worksheets in workbook order. Worksheets are constructed on
    # demand and memoized, so iterating then editing is consistent with
    # {#sheet}.
    #
    # @yieldparam worksheet [Rbxl::EditableWorksheet]
    # @return [Enumerator<Rbxl::EditableWorksheet>] when no block is given
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def sheets
      ensure_open!
      return enum_for(:sheets) unless block_given?

      @sheet_names.each { |name| yield sheet(name) }
    end

    # Writes the workbook out, preserving every part that has not been
    # mutated byte-for-byte. Worksheets whose cells have been edited are
    # re-serialized from their in-memory Nokogiri document; all other
    # entries (styles, sharedStrings, drawings, charts, pivot caches,
    # custom XML, rels) are streamed straight from the source ZIP without
    # re-parsing.
    #
    # +path+ defaults to the original load path; passing +nil+ or omitting
    # it saves in place. The new file is written to a temp file in the same
    # directory and atomically renamed into place, so a crash mid-write
    # never leaves a half-written workbook. On success, dirty flags on each
    # touched worksheet are cleared, so the object is reusable for further
    # edits and another {#save}.
    #
    # @param path [String, #to_path, nil] destination path; defaults to the
    #   path the workbook was opened from
    # @return [String] the path that was written
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def save(path = nil)
      ensure_open!
      out_path = (path || @path).to_s
      overrides = collect_overrides

      tmp_path = "#{out_path}.rbxl-tmp.#{Process.pid}.#{rand(1 << 32).to_s(16)}"
      begin
        Zip::OutputStream.open(tmp_path) do |out|
          @zip.each do |entry|
            next if entry.directory?

            if (override_xml = overrides[entry.name])
              out.put_next_entry(entry.name)
              out.write(override_xml)
            else
              out.copy_raw_entry(entry)
            end
          end
        end
        File.rename(tmp_path, out_path)
      rescue StandardError
        File.unlink(tmp_path) if File.exist?(tmp_path)
        raise
      end

      @sheets_by_name.each_value(&:clear_dirty!)
      out_path
    end

    # Releases the underlying ZIP file. Idempotent.
    #
    # @return [Boolean] +true+ on the first call, +false+ on subsequent calls
    def close
      return false if @closed

      @zip&.close
      @zip = nil
      @closed = true
      true
    end

    # @return [Boolean]
    def closed?
      @closed
    end

    private

    def ensure_open!
      raise ClosedWorkbookError, "workbook has been closed" if @closed
    end

    def resolve_sheet_name(key)
      return key unless key.is_a?(Integer)

      name = @sheet_names[key]
      return name if name

      raise SheetNotFoundError, "sheet index out of range: #{key} (#{@sheet_names.length} sheet(s))"
    end

    def ensure_xlsx_format!(path)
      header = begin
        File.binread(path, 8)
      rescue Errno::ENOENT, Errno::EISDIR, Errno::EACCES => e
        raise UnsupportedFormatError, "#{path}: #{e.message}"
      end

      raise UnsupportedFormatError, "#{path}: file is empty or unreadable" if header.nil? || header.empty?
      return if header.start_with?(ZIP_LOCAL_MAGIC)

      if header.start_with?(OLE_CFB_MAGIC)
        raise UnsupportedFormatError,
              "#{path} looks like a legacy .xls (BIFF/CFB). " \
              "rbxl supports .xlsx (OOXML) only; convert first, e.g. " \
              "`libreoffice --headless --convert-to xlsx #{File.basename(path.to_s)}`."
      end

      raise UnsupportedFormatError,
            "#{path} is not a valid .xlsx (no ZIP signature at offset 0)."
    end

    def locate_workbook_part
      doc = parse_xml("_rels/.rels")
      rel = doc.at_xpath(
        "/pkg:Relationships/pkg:Relationship[@Type=$type]",
        { "pkg" => PACKAGE_REL_NS },
        { "type" => OFFICE_DOC_REL_TYPE }
      )
      raise WorkbookFormatError, "#{@path}: officeDocument relationship missing from _rels/.rels" unless rel

      target = rel["Target"] or raise WorkbookFormatError, "#{@path}: officeDocument relationship has no Target"
      target.sub(%r{\A/}, "")
    end

    def load_sheet_entries
      rels = parse_rels(rels_path_for(@workbook_part))
      doc = parse_xml(@workbook_part)
      sheets = {}

      doc.xpath("/main:workbook/main:sheets/main:sheet", "main" => MAIN_NS).each do |sheet_node|
        name = sheet_node["name"]
        rid = sheet_node.attribute_with_ns("id", REL_NS)&.value
        next unless name && rid

        target = rels.fetch(rid) do
          raise WorkbookFormatError,
                "workbook #{@path} references missing relationship #{rid.inspect} for sheet #{name.inspect}"
        end
        sheets[name] = resolve_relative(@workbook_dir, target)
      end

      sheets
    end

    def shared_strings
      @shared_strings ||= SharedStringsLoader.load(@zip)
    end

    def collect_overrides
      @sheets_by_name.each_with_object({}) do |(_, ws), h|
        h[ws.entry_path] = ws.to_xml if ws.dirty?
      end
    end

    def parse_xml(part_name)
      entry = @zip.find_entry(part_name)
      raise WorkbookFormatError, "#{@path}: missing part #{part_name}" unless entry

      doc = Nokogiri::XML(entry.get_input_stream.read)
      raise WorkbookFormatError, "#{@path}: #{part_name}: #{doc.errors.first}" unless doc.errors.empty?

      doc
    end

    def parse_rels(rels_part)
      entry = @zip.find_entry(rels_part)
      return {} unless entry

      doc = Nokogiri::XML(entry.get_input_stream.read)
      doc.xpath("/pkg:Relationships/pkg:Relationship", "pkg" => PACKAGE_REL_NS).each_with_object({}) do |r, h|
        h[r["Id"]] = r["Target"]
      end
    end

    def rels_path_for(part_name)
      dir = File.dirname(part_name)
      base = File.basename(part_name)
      dir == "." ? "_rels/#{base}.rels" : "#{dir}/_rels/#{base}.rels"
    end

    def resolve_relative(base_dir, target)
      return target.sub(%r{\A/}, "") if target.start_with?("/")

      File.expand_path(target, "/#{base_dir}").sub(%r{\A/}, "")
    end
  end
end
