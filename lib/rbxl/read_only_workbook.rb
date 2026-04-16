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
            strings << current_fragments.join
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
      workbook_xml = REXML::Document.new(read_entry("xl/workbook.xml"))
      rels_xml = REXML::Document.new(read_entry("xl/_rels/workbook.xml.rels"))

      relationships = {}
      REXML::XPath.each(rels_xml, "//rel:Relationship", { "rel" => PACKAGE_REL_NS }) do |rel|
        relationships[rel.attributes["Id"]] = rel.attributes["Target"]
      end

      sheets = {}
      REXML::XPath.each(workbook_xml, "//main:sheets/main:sheet", { "main" => MAIN_NS }) do |sheet|
        name = sheet.attributes["name"]
        rid = sheet.attributes["r:id"]
        target = relationships.fetch(rid)
        sheets[name] = "xl/#{target}".gsub(%r{/+}, "/")
      end
      sheets
    end

    def read_entry(name)
      @zip.get_entry(name).get_input_stream.read
    end
  end
end
