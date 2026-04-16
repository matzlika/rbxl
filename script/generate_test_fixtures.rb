#!/usr/bin/env ruby

require "fileutils"
require "zip"

root = File.expand_path("..", __dir__)
fixtures_dir = File.join(root, "test", "fixtures")

FileUtils.mkdir_p(fixtures_dir)

def write_entry(zip, name, content)
  zip.put_next_entry(name)
  zip.write(content)
end

def default_styles_xml
  <<~XML.chomp
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    </styleSheet>
  XML
end

def workbook_xml(sheet_name)
  <<~XML.chomp
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="#{sheet_name}" sheetId="1" r:id="rId1"/></sheets>
    </workbook>
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

def workbook_rels_xml(extra = "")
  <<~XML.chomp
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      #{extra}
    </Relationships>
  XML
end

def content_types_xml(extra = "")
  <<~XML.chomp
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      #{extra}
    </Types>
  XML
end

def create_fixture(path)
  FileUtils.rm_f(path)
  Zip::OutputStream.open(path) { |zip| yield zip }
end

create_fixture(File.join(fixtures_dir, "sparse.xlsx")) do |zip|
  write_entry(zip, "[Content_Types].xml", content_types_xml)
  write_entry(zip, "_rels/.rels", root_rels_xml)
  write_entry(zip, "xl/workbook.xml", workbook_xml("Sparse"))
  write_entry(zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml)
  write_entry(zip, "xl/styles.xml", default_styles_xml)
  write_entry(zip, "xl/worksheets/sheet1.xml", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:C1"/>
      <sheetData>
        <row r="1">
          <c r="A1" t="inlineStr"><is><t>left</t></is></c>
          <c r="C1" t="inlineStr"><is><t>right</t></is></c>
        </row>
      </sheetData>
    </worksheet>
  XML
end

create_fixture(File.join(fixtures_dir, "shared_strings.xlsx")) do |zip|
  write_entry(zip, "[Content_Types].xml", content_types_xml(%(<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>)))
  write_entry(zip, "_rels/.rels", root_rels_xml)
  write_entry(zip, "xl/workbook.xml", workbook_xml("Strings"))
  write_entry(zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml(%(<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>)))
  write_entry(zip, "xl/styles.xml", default_styles_xml)
  write_entry(zip, "xl/sharedStrings.xml", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
      <si><t>alpha</t></si>
      <si><t></t></si>
    </sst>
  XML
  write_entry(zip, "xl/worksheets/sheet1.xml", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:A2"/>
      <sheetData>
        <row r="1"><c r="A1" t="s"><v>0</v></c></row>
        <row r="2"><c r="A2" t="s"><v>1</v></c></row>
      </sheetData>
    </worksheet>
  XML
end

create_fixture(File.join(fixtures_dir, "no_dimension.xlsx")) do |zip|
  write_entry(zip, "[Content_Types].xml", content_types_xml)
  write_entry(zip, "_rels/.rels", root_rels_xml)
  write_entry(zip, "xl/workbook.xml", workbook_xml("NoDimension"))
  write_entry(zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml)
  write_entry(zip, "xl/styles.xml", default_styles_xml)
  write_entry(zip, "xl/worksheets/sheet1.xml", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" t="inlineStr"><is><t>x</t></is></c>
          <c r="C1"><v>3</v></c>
        </row>
        <row r="2">
          <c r="B2"><v>2</v></c>
        </row>
      </sheetData>
    </worksheet>
  XML
end

create_fixture(File.join(fixtures_dir, "sparse_rows.xlsx")) do |zip|
  write_entry(zip, "[Content_Types].xml", content_types_xml)
  write_entry(zip, "_rels/.rels", root_rels_xml)
  write_entry(zip, "xl/workbook.xml", workbook_xml("SparseRows"))
  write_entry(zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml)
  write_entry(zip, "xl/styles.xml", default_styles_xml)
  write_entry(zip, "xl/worksheets/sheet1.xml", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:C2"/>
      <sheetData>
        <row r="1">
          <c r="A1" t="inlineStr"><is><t>top</t></is></c>
        </row>
        <row r="2">
          <c r="C2" t="inlineStr"><is><t>tail</t></is></c>
        </row>
      </sheetData>
    </worksheet>
  XML
end

create_fixture(File.join(fixtures_dir, "implicit_coordinates.xlsx")) do |zip|
  write_entry(zip, "[Content_Types].xml", content_types_xml)
  write_entry(zip, "_rels/.rels", root_rels_xml)
  write_entry(zip, "xl/workbook.xml", workbook_xml("Implicit"))
  write_entry(zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml)
  write_entry(zip, "xl/styles.xml", default_styles_xml)
  write_entry(zip, "xl/worksheets/sheet1.xml", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:C2"/>
      <sheetData>
        <row>
          <c t="inlineStr"><is><t>Test</t></is></c>
        </row>
        <row>
          <c t="inlineStr"><is><t>A2</t></is></c>
          <c t="inlineStr"><is><t>B2</t></is></c>
          <c t="inlineStr"><is><t>C2</t></is></c>
        </row>
      </sheetData>
    </worksheet>
  XML
end

create_fixture(File.join(fixtures_dir, "file_item_error.xlsx")) do |zip|
  write_entry(zip, "[Content_Types].xml", content_types_xml)
  write_entry(zip, "_rels/.rels", root_rels_xml)
  write_entry(zip, "xl/workbook.xml", workbook_xml("BrokenRels"))
  write_entry(zip, "xl/_rels/workbook.xml.rels", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="missing/calcChain.xml"/>
    </Relationships>
  XML
  write_entry(zip, "xl/styles.xml", default_styles_xml)
  write_entry(zip, "xl/worksheets/sheet1.xml", <<~XML.chomp)
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="A1:A1"/>
      <sheetData>
        <row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c></row>
      </sheetData>
    </worksheet>
  XML
end

puts "generated fixtures in #{fixtures_dir}"
