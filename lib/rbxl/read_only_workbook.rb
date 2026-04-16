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

    # @return [String] filesystem path the workbook was opened from
    attr_reader :path

    # @return [Array<String>] visible sheet names in workbook order
    attr_reader :sheet_names

    # Convenience constructor equivalent to <tt>new(path)</tt>.
    #
    # @param path [String, #to_path] path to the <tt>.xlsx</tt> file
    # @return [Rbxl::ReadOnlyWorkbook]
    def self.open(path)
      new(path)
    end

    # Opens the ZIP archive, pre-loads shared strings, and indexes the
    # worksheet entries keyed by visible sheet name.
    #
    # @param path [String, #to_path] path to the <tt>.xlsx</tt> file
    def initialize(path)
      @path = path
      @zip = Zip::File.open(path)
      @shared_strings = load_shared_strings
      @sheet_entries = load_sheet_entries
      @sheet_names = @sheet_entries.keys.freeze
      @closed = false
    end

    # Returns a streaming worksheet by visible sheet name.
    #
    # The returned object shares the workbook's ZIP handle. Closing the
    # workbook invalidates any worksheets produced by prior calls.
    #
    # @param name [String] visible sheet name as listed in {#sheet_names}
    # @return [Rbxl::ReadOnlyWorksheet]
    # @raise [Rbxl::SheetNotFoundError] if +name+ is not present
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    def sheet(name)
      ensure_open!

      entry_path = @sheet_entries.fetch(name) do
        raise SheetNotFoundError, "sheet not found: #{name}"
      end

      ReadOnlyWorksheet.new(zip: @zip, entry_path: entry_path, shared_strings: @shared_strings, name: name)
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

    def load_shared_strings
      entry = @zip.find_entry("xl/sharedStrings.xml")
      return [] unless entry

      max_count = Rbxl.max_shared_strings
      max_bytes = Rbxl.max_shared_string_bytes

      # Reject zip-bomb style entries up front using the ZIP directory's
      # declared uncompressed size, before allocating any decompression buffer.
      if max_bytes && entry.size && entry.size > max_bytes
        raise SharedStringsTooLargeError,
              "shared strings uncompressed size #{entry.size} exceeds limit #{max_bytes}"
      end

      strings = []
      total_bytes = 0
      io = entry.get_input_stream
      reader = Nokogiri::XML::Reader(io)

      in_si = false
      in_run = false
      in_phonetic = false
      collecting_text = false
      buffer = +""
      current_fragments = []

      reader.each do |node|
        case node.node_type
        when Nokogiri::XML::Reader::TYPE_ELEMENT
          case node.local_name
          when "si"
            in_si = true
            current_fragments = []
          when "r"
            in_run = true if in_si
          when "rPh"
            in_phonetic = true if in_si
          when "t"
            next unless in_si && !in_phonetic

            collecting_text = !in_run || node.depth.positive?
            buffer.clear if collecting_text
          end
        when Nokogiri::XML::Reader::TYPE_TEXT, Nokogiri::XML::Reader::TYPE_CDATA
          buffer << node.value if collecting_text
        when Nokogiri::XML::Reader::TYPE_END_ELEMENT
          case node.local_name
          when "t"
            if collecting_text
              current_fragments << buffer.dup
              collecting_text = false
            end
          when "r"
            in_run = false
          when "rPh"
            in_phonetic = false
          when "si"
            value = current_fragments.join.freeze
            total_bytes += value.bytesize
            if max_bytes && total_bytes > max_bytes
              raise SharedStringsTooLargeError,
                    "shared strings total size exceeds limit #{max_bytes}"
            end
            strings << value
            if max_count && strings.size > max_count
              raise SharedStringsTooLargeError,
                    "shared strings count exceeds limit #{max_count}"
            end
            in_si = false
            in_run = false
            in_phonetic = false
            collecting_text = false
          end
        end
      end

      strings
    ensure
      io&.close
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

        target = relationships.fetch(rid)
        sheets[name] = "xl/#{target}".gsub(%r{/+}, "/")
      end

      sheets
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
      io = @zip.get_entry(entry_path).get_input_stream
      reader = Nokogiri::XML::Reader(io)
      reader.each { |node| yield node }
    ensure
      io&.close
    end
  end
end
