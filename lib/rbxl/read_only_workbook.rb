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

      xml = REXML::Document.new(entry.get_input_stream.read)
      strings = []
      REXML::XPath.each(xml, "//main:si", { "main" => MAIN_NS }) do |node|
        strings << shared_string_text(node)
      end
      strings
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

    def shared_string_text(node)
      fragments = []

      node.children.each do |child|
        next unless child.is_a?(REXML::Element)

        case child.name
        when "t"
          fragments << child.text.to_s
        when "r"
          text = REXML::XPath.first(child, "./main:t", { "main" => MAIN_NS })
          fragments << text.text.to_s if text
        end
      end

      fragments.join
    end
  end
end
