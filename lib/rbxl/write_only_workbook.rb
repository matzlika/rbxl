module Rbxl
  class WriteOnlyWorkbook
    attr_reader :worksheets

    def initialize
      @worksheets = []
      @closed = false
      @saved = false
    end

    def add_sheet(name)
      ensure_writable!

      sheet = WriteOnlyWorksheet.new(name: name)
      @worksheets << sheet
      sheet
    end

    def save(path)
      ensure_writable!
      raise Error, "at least one worksheet is required" if worksheets.empty?

      Zip::OutputStream.open(path) do |zip|
        write_entry(zip, "[Content_Types].xml", content_types_xml)
        write_entry(zip, "_rels/.rels", root_rels_xml)
        write_entry(zip, "xl/workbook.xml", workbook_xml)
        write_entry(zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml)
        write_entry(zip, "xl/styles.xml", styles_xml)

        worksheets.each_with_index do |sheet, index|
          write_entry(zip, "xl/worksheets/sheet#{index + 1}.xml", sheet.to_xml)
        end
      end

      @saved = true
      close
      path
    end

    def close
      @closed = true
    end

    def closed?
      @closed
    end

    private

    def ensure_writable!
      raise ClosedWorkbookError, "workbook has been closed" if closed?
      raise WorkbookAlreadySavedError, "write-only workbook can only be saved once" if @saved
    end

    def write_entry(zip, name, content)
      zip.put_next_entry(name)
      zip.write(content)
    end

    def content_types_xml
      worksheet_overrides = worksheets.each_index.map do |index|
        %(<Override PartName="/xl/worksheets/sheet#{index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)
      end.join

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          #{worksheet_overrides}
        </Types>
      XML
    end

    def root_rels_xml
      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      XML
    end

    def workbook_xml
      sheet_nodes = worksheets.each_with_index.map do |sheet, index|
        %(<sheet name="#{escape(sheet.name)}" sheetId="#{index + 1}" r:id="rId#{index + 1}"/>)
      end.join

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>#{sheet_nodes}</sheets>
        </workbook>
      XML
    end

    def workbook_rels_xml
      relationship_nodes = worksheets.each_with_index.map do |_, index|
        %(<Relationship Id="rId#{index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{index + 1}.xml"/>)
      end
      relationship_nodes << %(<Relationship Id="rId#{worksheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>)

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          #{relationship_nodes.join}
        </Relationships>
      XML
    end

    def styles_xml
      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1">
            <font><sz val="11"/><name val="Calibri"/></font>
          </fonts>
          <fills count="1">
            <fill><patternFill patternType="none"/></fill>
          </fills>
          <borders count="1">
            <border><left/><right/><top/><bottom/><diagonal/></border>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
          </cellXfs>
          <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
          </cellStyles>
        </styleSheet>
      XML
    end

    def escape(value)
      CGI.escapeHTML(value.to_s)
    end
  end
end
