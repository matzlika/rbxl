module Rbxl
  class ReadOnlyWorkbook
    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    attr_reader :path, :sheet_names

    def self.open(path)
      new(path)
    end

    def initialize(path)
      @path = path
      @zip = Zip::File.open(path)
      @shared_strings = load_shared_strings
      @sheet_entries = load_sheet_entries
      @sheet_names = @sheet_entries.keys.freeze
      @closed = false
    end

    def sheet(name)
      ensure_open!

      entry_path = @sheet_entries.fetch(name) do
        raise SheetNotFoundError, "sheet not found: #{name}"
      end

      ReadOnlyWorksheet.new(zip: @zip, entry_path: entry_path, shared_strings: @shared_strings, name: name)
    end

    def close
      return if closed?

      @zip.close
      @closed = true
    end

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

      strings = []
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
            strings << current_fragments.join.freeze
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
