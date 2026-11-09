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
  # == Styles
  #
  # {EditableCell#style} tweaks a cell's existing (template) style and
  # {#add_style} registers a new one. Both only <em>append</em> entries to
  # +xl/styles.xml+, so the style indices existing cells use never change;
  # the part is rewritten only when a style was actually added.
  #
  # == Recalculation
  #
  # Cached formula results in a template (e.g. a +SUM+ over line items) go
  # stale once you write the inputs. When any cell was edited, {#save} sets
  # <tt><calcPr fullCalcOnLoad="1"/></tt> in +xl/workbook.xml+ so Excel and
  # LibreOffice recompute on open; pass <tt>recalc: false</tt> to skip it.
  # If an edit replaced a formula with a value, +xl/calcChain.xml+ is
  # dropped (Excel rebuilds it) so it never points at a non-formula cell.
  #
  # == Out of scope
  #
  # * inserting / deleting rows or columns
  # * deleting / renaming / reordering / duplicating sheets ({#add_sheet}
  #   appends a new empty sheet)
  # * editing formulas, defined names, drawings, or shared strings
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
      @workbook_doc = nil
      @workbook_rels_doc = nil
      @content_types_doc = nil
      @modified_parts = Set.new
      @removed_parts = Set.new
      @styles = nil
      @date_1904 = nil
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
        workbook: self,
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

    # Appends a new, empty worksheet after the existing ones and returns it
    # for editing. The sheet part, its workbook relationship, and its
    # content-type override are created; nothing else in the package
    # changes. (+docProps/app.xml+ keeps listing the old sheet titles;
    # Excel and LibreOffice ignore that metadata on open.)
    #
    # @param name [String] sheet name (1–31 characters, none of
    #   <tt>[ ] : * ? / \\</tt>, not starting or ending with +'+, unique
    #   case-insensitively)
    # @return [Rbxl::EditableWorksheet]
    # @raise [ArgumentError] if +name+ is invalid or already taken
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def add_sheet(name)
      ensure_open!
      validate_sheet_name!(name)

      part = next_sheet_part_name
      rid = add_workbook_relationship(SHEET_REL_TYPE, relative_target(part))
      add_content_type_override(part, SHEET_CONTENT_TYPE)

      sheets_node = workbook_doc.at_xpath("/main:workbook/main:sheets", "main" => MAIN_NS) or
        raise WorkbookFormatError, "#{@path}: #{@workbook_part} has no <sheets>"
      ids = XmlEdit.children(sheets_node, "sheet").map { |n| n["sheetId"].to_i }
      prefix = relationship_prefix(workbook_doc)
      XmlEdit.append(
        sheets_node,
        %(<sheet name="#{StyleXml.escape_attr(name)}" sheetId="#{(ids.max || 0) + 1}" #{prefix}:id="#{rid}"/>)
      )
      @modified_parts << @workbook_part

      @sheet_entries[name] = part
      @sheet_names = @sheet_entries.keys.freeze
      @sheets_by_name[name] = EditableWorksheet.new(
        workbook: self,
        zip: @zip,
        entry_path: part,
        workbook_path: @path,
        shared_strings: shared_strings,
        name: name,
        source_xml: EMPTY_SHEET_XML
      )
    end

    # Registers a new cell style and returns its +style_id+ for
    # {EditableCell#style_id=}. Accepts the same keywords as
    # {WriteOnlyWorkbook#add_style}; unspecified components fall back to the
    # workbook's defaults (entry 0). To change only part of a template
    # cell's look, prefer {EditableCell#style}.
    #
    #   red = book.add_style(font: { color: "C00000", bold: true })
    #   sheet["F20"].style_id = red
    #
    # @return [Integer] the new style's +cellXfs+ index
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def add_style(font: nil, fill: nil, border: nil, alignment: nil, number_format: nil)
      ensure_open!
      styles.register(
        Style.new(font: font, fill: fill, border: border, alignment: alignment, number_format: number_format)
      )
    end

    # @return [Array<String>] defined names in workbook order (a name that
    #   exists both globally and sheet-locally is listed once)
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def defined_names
      ensure_open!
      defined_name_nodes.map { |n| n["name"] }.uniq
    end

    # Resolves a defined name ("名前付きセル") that refers to a single cell
    # and returns that cell, so templates can be filled by name instead of
    # by coordinate:
    #
    #   book.defined_name("顧客名").value = "株式会社Acme"
    #
    # Names match case-insensitively, as in Excel. A workbook-scoped name
    # wins over sheet-scoped ones; pass +sheet:+ to pick a sheet-scoped name.
    #
    # @param name [String]
    # @param sheet [String, Integer, nil] restrict to the name scoped to this
    #   sheet
    # @return [Rbxl::EditableCell]
    # @raise [Rbxl::DefinedNameError] if the name is missing, ambiguous, or
    #   refers to anything other than a single cell
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def defined_name(name, sheet: nil)
      ensure_open!
      node = find_defined_name(name.to_s, sheet)
      sheet_name, coordinate = parse_single_cell_reference(node)
      unless @sheet_entries.key?(sheet_name)
        raise DefinedNameError, "defined name #{name.inspect} refers to unknown sheet #{sheet_name.inspect}"
      end

      self.sheet(sheet_name).cell(coordinate)
    end

    # Writes the workbook out, preserving every part that has not been
    # mutated byte-for-byte. Edited worksheets, and package parts changed by
    # {#add_sheet} / {#add_style} / recalculation, are re-serialized from
    # memory; all other entries (styles, sharedStrings, drawings, charts,
    # pivot caches, custom XML, rels) are streamed straight from the source
    # ZIP without re-parsing.
    #
    # +path+ defaults to the original load path; passing +nil+ or omitting
    # it saves in place. The new file is written to a temp file in the same
    # directory and atomically renamed into place, so a crash mid-write
    # never leaves a half-written workbook. The object stays usable: further
    # edits and saves include everything edited so far.
    #
    # @param path [String, #to_path, nil] destination path; defaults to the
    #   path the workbook was opened from
    # @param recalc [Boolean] when any cell was edited, ask Excel to
    #   recompute every formula on open (see "Recalculation" above)
    # @return [String] the path that was written
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def save(path = nil, recalc: true)
      ensure_open!
      out_path = (path || @path).to_s
      prepare_workbook_parts(recalc)
      parts = collect_parts

      tmp_path = "#{out_path}.rbxl-tmp.#{Process.pid}.#{rand(1 << 32).to_s(16)}"
      begin
        Zip::OutputStream.open(tmp_path) do |out|
          written = Set.new
          @zip.each do |entry|
            next if entry.directory? || @removed_parts.include?(entry.name)

            if (xml = parts[entry.name])
              out.put_next_entry(entry.name)
              out.write(xml)
            else
              out.copy_raw_entry(entry)
            end
            written << entry.name
          end

          parts.each do |name, xml|
            next if written.include?(name)

            out.put_next_entry(name)
            out.write(xml)
          end
        end
        File.rename(tmp_path, out_path)
      rescue StandardError
        File.unlink(tmp_path) if File.exist?(tmp_path)
        raise
      end

      reopen_source! if File.expand_path(out_path) == File.expand_path(@path)
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

    # @api private
    # Lazily loaded, append-only view of the workbook's styles part. A
    # minimal styles part is created when the workbook has none.
    def styles
      @styles ||= load_styles
    end

    # @api private
    # @return [Boolean] whether the workbook uses the 1904 date system
    def date_1904?
      if @date_1904.nil?
        pr = workbook_doc.at_xpath("/main:workbook/main:workbookPr", "main" => MAIN_NS)
        @date_1904 = %w[1 true].include?(pr&.[]("date1904"))
      end
      @date_1904
    end
    public :styles, :date_1904?

    SHEET_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet".freeze
    STYLES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles".freeze
    CALC_CHAIN_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain".freeze
    SHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml".freeze
    STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml".freeze
    CONTENT_TYPES_PART = "[Content_Types].xml".freeze
    CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types".freeze
    private_constant :SHEET_REL_TYPE, :STYLES_REL_TYPE, :CALC_CHAIN_REL_TYPE,
                     :SHEET_CONTENT_TYPE, :STYLES_CONTENT_TYPE, :CONTENT_TYPES_PART, :CONTENT_TYPES_NS

    EMPTY_SHEET_XML = <<~XML.freeze
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="#{MAIN_NS}" xmlns:r="#{REL_NS}"><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>
    XML
    private_constant :EMPTY_SHEET_XML

    # <workbook> child sequence (ECMA-376 part 1 §18.2.27), for calcPr.
    WORKBOOK_ORDER = %w[fileVersion fileSharing workbookPr workbookProtection bookViews sheets
                        functionGroups externalReferences definedNames calcPr oleSize
                        customWorkbookViews pivotCaches smartTagPr smartTagTypes webPublishing
                        fileRecoveryPr webPublishObjects extLst].freeze
    private_constant :WORKBOOK_ORDER

    INVALID_SHEET_NAME_CHARS = %r{[\[\]:*?/\\]}.freeze
    private_constant :INVALID_SHEET_NAME_CHARS

    def load_styles
      rel = workbook_rels_doc.xpath("/pkg:Relationships/pkg:Relationship", "pkg" => PACKAGE_REL_NS)
                             .find { |r| r["Type"] == STYLES_REL_TYPE }
      if rel
        part = resolve_relative(@workbook_dir, rel["Target"])
        entry = @zip.find_entry(part) or raise WorkbookFormatError, "#{@path}: missing styles part #{part}"
        EditableStyles.new(entry.get_input_stream.read, part_name: part, workbook_path: @path)
      else
        part = @workbook_dir == "." ? "styles.xml" : "#{@workbook_dir}/styles.xml"
        add_workbook_relationship(STYLES_REL_TYPE, relative_target(part))
        add_content_type_override(part, STYLES_CONTENT_TYPE)
        @created_styles = true
        EditableStyles.new(EditableStyles::MINIMAL_XML, part_name: part, workbook_path: @path)
      end
    end

    def prepare_workbook_parts(recalc)
      edited = @sheets_by_name.values.select(&:modified?)
      drop_calc_chain! if edited.any?(&:formula_removed?)
      request_full_calc_on_load! if recalc && edited.any?
    end

    def request_full_calc_on_load!
      root = workbook_doc.root
      calc_pr = XmlEdit.child(root, "calcPr") ||
                XmlEdit.insert_in_order(root, "calcPr", "<calcPr/>", WORKBOOK_ORDER)
      return if calc_pr["fullCalcOnLoad"] == "1"

      calc_pr["fullCalcOnLoad"] = "1"
      @modified_parts << @workbook_part
    end

    def drop_calc_chain!
      rels = workbook_rels_doc.xpath("/pkg:Relationships/pkg:Relationship", "pkg" => PACKAGE_REL_NS)
                              .select { |r| r["Type"] == CALC_CHAIN_REL_TYPE }
      return if rels.empty?

      rels.each do |rel|
        part = resolve_relative(@workbook_dir, rel["Target"])
        @removed_parts << part
        override = content_types_doc.root.element_children.find do |n|
          n.name == "Override" && n["PartName"] == "/#{part}"
        end
        if override
          override.unlink
          @modified_parts << CONTENT_TYPES_PART
        end
        rel.unlink
      end
      @modified_parts << rels_path_for(@workbook_part)
    end

    def collect_parts
      parts = {}
      @sheets_by_name.each_value { |ws| parts[ws.entry_path] = ws.to_xml if ws.modified? }
      parts[@styles.part_name] = @styles.to_xml if @styles && (@styles.modified? || @created_styles)
      parts[@workbook_part] = XmlEdit.serialize(workbook_doc) if @modified_parts.include?(@workbook_part)
      rels = rels_path_for(@workbook_part)
      parts[rels] = XmlEdit.serialize(workbook_rels_doc) if @modified_parts.include?(rels)
      parts[CONTENT_TYPES_PART] = XmlEdit.serialize(content_types_doc) if @modified_parts.include?(CONTENT_TYPES_PART)
      parts
    end

    # rubyzip re-opens the archive by file name for every lazy read, so once
    # an in-place save has replaced that file the old central directory no
    # longer matches. Re-open and hand the fresh handle to every sheet.
    def reopen_source!
      @zip.close
      @zip = Zip::File.open(@path)
      @sheets_by_name.each_value { |ws| ws.rebind_zip!(@zip) }
      @removed_parts.clear
    end

    def workbook_doc
      @workbook_doc ||= parse_xml(@workbook_part)
    end

    def workbook_rels_doc
      @workbook_rels_doc ||= begin
        path = rels_path_for(@workbook_part)
        if @zip.find_entry(path)
          parse_xml(path)
        else
          Nokogiri::XML(%(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="#{PACKAGE_REL_NS}"/>))
        end
      end
    end

    def content_types_doc
      @content_types_doc ||= parse_xml(CONTENT_TYPES_PART)
    end

    def add_workbook_relationship(type, target)
      root = workbook_rels_doc.root
      used = root.element_children.map { |r| r["Id"] }
      n = used.filter_map { |id| id[/\ArId(\d+)\z/, 1]&.to_i }.max.to_i + 1
      n += 1 while used.include?("rId#{n}")
      rid = "rId#{n}"
      XmlEdit.append(root, %(<Relationship Id="#{rid}" Type="#{type}" Target="#{StyleXml.escape_attr(target)}"/>))
      @modified_parts << rels_path_for(@workbook_part)
      rid
    end

    def add_content_type_override(part, content_type)
      XmlEdit.append(content_types_doc.root,
                     %(<Override PartName="/#{StyleXml.escape_attr(part)}" ContentType="#{content_type}"/>))
      @modified_parts << CONTENT_TYPES_PART
    end

    def relationship_prefix(doc)
      prefix = doc.root.namespaces.find { |_, href| href == REL_NS }&.first&.delete_prefix("xmlns:")
      return prefix if prefix && !prefix.empty? && prefix != "xmlns"

      doc.root.add_namespace_definition("r", REL_NS)
      "r"
    end

    def relative_target(part)
      prefix = @workbook_dir == "." ? "" : "#{@workbook_dir}/"
      part.start_with?(prefix) ? part.delete_prefix(prefix) : "/#{part}"
    end

    def next_sheet_part_name
      dir = @workbook_dir == "." ? "worksheets" : "#{@workbook_dir}/worksheets"
      taken = Set.new(@zip.entries.map(&:name)) | @sheet_entries.values
      n = 1
      n += 1 while taken.include?("#{dir}/sheet#{n}.xml")
      "#{dir}/sheet#{n}.xml"
    end

    def validate_sheet_name!(name)
      raise ArgumentError, "sheet name must be a String" unless name.is_a?(String)
      raise ArgumentError, "sheet name must be 1-31 characters: #{name.inspect}" if name.empty? || name.length > 31
      raise ArgumentError, "sheet name contains an invalid character ([]:*?/\\): #{name.inspect}" if name.match?(INVALID_SHEET_NAME_CHARS)
      raise ArgumentError, "sheet name cannot start or end with an apostrophe: #{name.inspect}" if name.start_with?("'") || name.end_with?("'")
      return unless @sheet_entries.keys.any? { |existing| existing.casecmp?(name) }

      raise ArgumentError, "sheet already exists: #{name.inspect}"
    end

    def defined_name_nodes
      workbook_doc.xpath("/main:workbook/main:definedNames/main:definedName", "main" => MAIN_NS).to_a
    end

    def find_defined_name(name, sheet)
      matches = defined_name_nodes.select { |n| n["name"]&.casecmp?(name) }
      raise DefinedNameError, "defined name not found: #{name.inspect}" if matches.empty?

      if sheet
        index = @sheet_entries.keys.index(resolve_sheet_name(sheet)) or
          raise SheetNotFoundError, "sheet not found: #{sheet}"
        found = matches.find { |n| n["localSheetId"]&.to_i == index }
        return found if found

        raise DefinedNameError, "defined name #{name.inspect} is not scoped to sheet #{sheet.inspect}"
      end

      global = matches.find { |n| n["localSheetId"].nil? }
      return global if global
      return matches.first if matches.length == 1

      raise DefinedNameError,
            "defined name #{name.inspect} exists on several sheets; pass sheet: to choose one"
    end

    SINGLE_CELL_REF_RE = /\A(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?\z/.freeze
    private_constant :SINGLE_CELL_REF_RE

    def parse_single_cell_reference(node)
      ref = node.text.strip
      m = ref.match(SINGLE_CELL_REF_RE)
      single = m && (m[5].nil? || (m[5].casecmp?(m[3]) && m[6] == m[4]))
      unless single
        raise DefinedNameError, "defined name #{node['name'].inspect} refers to #{ref.inspect}, not a single cell"
      end

      sheet_name = m[1] ? m[1].gsub("''", "'") : m[2]
      [sheet_name, "#{m[3].upcase}#{m[4]}"]
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
