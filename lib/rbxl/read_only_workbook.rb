module Rbxl
  # Read-only workbook backed by a ZIP archive.
  #
  # The workbook opens the underlying <tt>.xlsx</tt> once and keeps a single
  # +Zip::File+ handle open for the lifetime of the object. Worksheets are
  # opened lazily via {#sheet}, so callers can process very large sheets
  # without materializing the full workbook in memory.
  #
  # Typical use:
  #
  #   book = Rbxl.open("big.xlsx", read_only: true)
  #   begin
  #     book.sheet_names                    # => ["Data"]
  #     book.sheet("Data").each_row do |row|
  #       process(row.values)
  #     end
  #   ensure
  #     book.close
  #   end
  #
  # After {#close} every subsequent {#sheet} call raises
  # {Rbxl::ClosedWorkbookError}.
  class ReadOnlyWorkbook
    # Namespace for the main SpreadsheetML schema.
    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    # Namespace used for document-level relationships.
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    # Namespace used by the OPC package relationships layer.
    PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    # First 8 bytes of the OLE Compound File Binary format (legacy .xls,
    # .doc, .ppt). Sniffed to short-circuit into a typed error before
    # rubyzip bubbles up an opaque "end of central directory" failure.
    OLE_CFB_MAGIC = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b.freeze
    private_constant :OLE_CFB_MAGIC

    # ZIP local file header signature — the first bytes of every .xlsx.
    ZIP_LOCAL_MAGIC = "PK\x03\x04".b.freeze
    private_constant :ZIP_LOCAL_MAGIC

    # @return [String] filesystem path the workbook was opened from
    attr_reader :path

    # @return [Array<String>] visible sheet names in workbook order
    attr_reader :sheet_names

    # Convenience constructor equivalent to
    # <tt>new(path, streaming:, date_conversion:)</tt>.
    #
    # When a block is given, the workbook is yielded to the block and
    # {#close} is called automatically when the block returns (or raises).
    # The block's return value is returned to the caller, matching the
    # +File.open+ / +Zip::File.open+ idiom.
    #
    # @param path [String, #to_path] path to the <tt>.xlsx</tt> file
    # @param streaming [Boolean] feed worksheet XML to the native parser in
    #   chunks (see {Rbxl.open})
    # @param date_conversion [Boolean] convert numeric cells backed by a
    #   date/time +numFmt+ to Ruby date/time objects (see {Rbxl.open})
    # @yieldparam book [Rbxl::ReadOnlyWorkbook] the opened workbook
    # @return [Rbxl::ReadOnlyWorkbook, Object] the workbook when no block is
    #   given, otherwise the block's return value
    def self.open(path, streaming: false, date_conversion: false)
      book = new(path, streaming: streaming, date_conversion: date_conversion)
      return book unless block_given?

      begin
        yield book
      ensure
        book.close
      end
    end

    # Opens the ZIP archive, pre-loads shared strings, and indexes the
    # worksheet entries keyed by visible sheet name.
    #
    # @param path [String, #to_path] path to the <tt>.xlsx</tt> file
    # @param streaming [Boolean] forwarded to produced worksheets
    # @param date_conversion [Boolean] lazily load styles.xml and forward the
    #   date-style lookup table to produced worksheets
    def initialize(path, streaming: false, date_conversion: false)
      @path = path
      ensure_xlsx_format!(path)
      @zip = Zip::File.open(path)
      @streaming = streaming
      @date_conversion = date_conversion
      @shared_strings = SharedStringsLoader.load(@zip)
      @sheet_entries = load_sheet_entries
      @sheet_names = @sheet_entries.keys.freeze
      @date_styles = nil
      @date_1904 = nil
      @closed = false
    end

    # Returns a row-by-row worksheet by visible sheet name or by 0-based
    # index into {#sheet_names}. Negative indexes count from the end, so
    # <tt>sheet(-1)</tt> returns the last sheet.
    #
    # The returned object shares the workbook's ZIP handle. Closing the
    # workbook invalidates any worksheets produced by prior calls.
    #
    # @param name_or_index [String, Integer] visible sheet name as listed in
    #   {#sheet_names}, or an integer index into that list
    # @return [Rbxl::ReadOnlyWorksheet]
    # @raise [Rbxl::SheetNotFoundError] if +name_or_index+ does not resolve
    #   to a sheet
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def sheet(name_or_index)
      ensure_open!

      name = resolve_sheet_name(name_or_index)
      entry_path = @sheet_entries.fetch(name) do
        raise SheetNotFoundError, "sheet not found: #{name}"
      end

      ReadOnlyWorksheet.new(
        zip: @zip,
        entry_path: entry_path,
        workbook_path: @path,
        shared_strings: @shared_strings,
        name: name,
        streaming: @streaming,
        date_styles: date_styles,
        date_1904: date_1904?
      )
    end

    # Iterates the workbook's sheets in workbook order. Each worksheet is
    # constructed on demand, so <tt>sheets.first</tt> allocates only the
    # first sheet and <tt>sheets.lazy.find { ... }</tt> stops as soon as a
    # match is found. Returned objects share the same ZIP handle and
    # cached shared-strings / date-style tables as {#sheet}.
    #
    # @yieldparam worksheet [Rbxl::ReadOnlyWorksheet]
    # @return [Enumerator<Rbxl::ReadOnlyWorksheet>] when no block is given
    # @return [void] when a block is given
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def sheets
      ensure_open!
      return enum_for(:sheets) unless block_given?

      @sheet_names.each { |name| yield sheet(name) }
    end

    # Releases the underlying ZIP file handle. Idempotent; subsequent calls
    # are no-ops.
    #
    # @return [void]
    def close
      return if closed?

      @zip.close
      @closed = true
    end

    # @return [Boolean] whether {#close} has been called
    def closed?
      @closed
    end

    private

    def ensure_open!
      raise ClosedWorkbookError, "workbook has been closed" if closed?
    end

    def resolve_sheet_name(key)
      return key unless key.is_a?(Integer)

      name = @sheet_names[key]
      return name if name

      raise SheetNotFoundError, "sheet index out of range: #{key} (#{@sheet_names.length} sheet(s))"
    end

    def ensure_xlsx_format!(path)
      header = File.binread(path, 8)
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

    # Built-in numFmtId values that Excel resolves to date/time formats.
    # Ids outside this set are dates only when the workbook provides a
    # matching custom +<numFmt>+ entry whose format code contains date
    # tokens. See ECMA-376 part 1 §18.8.30.
    BUILTIN_DATE_FMT_IDS = Set.new([14, 15, 16, 17, 18, 19, 20, 21, 22,
                                    27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
                                    45, 46, 47, 50, 51, 52, 53, 54, 55, 56,
                                    57, 58]).freeze

    def date_styles
      return nil unless @date_conversion

      @date_styles ||= load_date_styles
    end

    def date_1904?
      return false unless @date_conversion

      @date_1904 = load_date_1904 if @date_1904.nil?
      @date_1904
    end

    def load_date_styles
      entry = @zip.find_entry("xl/styles.xml")
      return [].freeze unless entry

      custom_date_ids = Set.new
      date_styles = []
      in_cell_xfs = false

      each_xml_node("xl/styles.xml") do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.local_name
          when "cellXfs"
            in_cell_xfs = true
          when "numFmt"
            id = node.attribute("numFmtId")
            code = node.attribute("formatCode")
            custom_date_ids << id.to_i if id && code && date_format_code?(code)
          when "xf"
            next unless in_cell_xfs

            fmt_id_int = node.attribute("numFmtId")&.to_i
            date_styles << (!fmt_id_int.nil? &&
                            (BUILTIN_DATE_FMT_IDS.include?(fmt_id_int) || custom_date_ids.include?(fmt_id_int)))
          end
        when Nokogiri::XML::Reader::TYPE_END_ELEMENT
          in_cell_xfs = false if node.local_name == "cellXfs"
        end
      end

      date_styles.freeze
    end

    # Quoted literals, bracketed directives (e.g. [Red], [$-409]), and
    # backslash-escaped characters never introduce date tokens, so strip
    # them before looking for +y/m/d/h/s+.
    def date_format_code?(code)
      stripped = code.dup
      stripped.gsub!(/\[[^\]]*\]/, "")
      stripped.gsub!(/"[^"]*"/, "")
      stripped.gsub!(/\\./, "")
      stripped.match?(/[ymdhs]/i)
    end

    def load_sheet_entries
      relationships = load_relationship_targets("xl/_rels/workbook.xml.rels")
      sheets = {}

      each_xml_node("xl/workbook.xml") do |node|
        next unless node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
        next unless node.local_name == "sheet"

        name = node.attribute("name")
        rid = node.attribute("r:id")
        next unless name && rid

        target = relationships.fetch(rid) do
          raise WorkbookFormatError,
                "workbook #{@path} references missing relationship #{rid.inspect} for sheet #{name.inspect}"
        end
        sheets[name] = "xl/#{target}".gsub(%r{/+}, "/")
      end

      sheets
    end

    def load_date_1904
      each_xml_node("xl/workbook.xml") do |node|
        next unless node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
        next unless node.local_name == "workbookPr"

        return xml_truthy?(node.attribute("date1904"))
      end

      false
    end

    def load_relationship_targets(entry_path)
      relationships = {}

      each_xml_node(entry_path) do |node|
        next unless node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
        next unless node.local_name == "Relationship"

        id = node.attribute("Id")
        target = node.attribute("Target")
        next unless id && target

        relationships[id] = target
      end

      relationships
    end

    def each_xml_node(entry_path)
      entry = @zip.get_entry(entry_path)
      raise WorkbookFormatError, "workbook #{@path} is missing required entry #{entry_path.inspect}" unless entry

      io = entry.get_input_stream
      reader = Nokogiri::XML::Reader(io)
      reader.each { |node| yield node }
    rescue Nokogiri::XML::SyntaxError => e
      raise WorkbookFormatError, "invalid workbook XML in #{@path} at #{entry_path}: #{e.message}"
    ensure
      io&.close
    end

    def xml_truthy?(value)
      value == "1" || value == "true"
    end
  end
end
