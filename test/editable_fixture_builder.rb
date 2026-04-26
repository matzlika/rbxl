# frozen_string_literal: true

require "zip"

module Rbxl
  module Test
    # Hand-builds a minimal +.xlsx+ fixture for the edit-mode tests so the
    # suite never depends on a binary file we can't fully account for. The
    # structure is the smallest one rbxl will open in edit mode:
    #
    # * +[Content_Types].xml+
    # * +_rels/.rels+
    # * +xl/workbook.xml+ + its rels
    # * +xl/styles.xml+ (one default xf, one styled xf)
    # * +xl/sharedStrings.xml+ (two entries)
    # * +xl/worksheets/sheet1.xml+ — exercises every read/edit code path
    # * +xl/worksheets/sheet2.xml+ — left untouched in tests so we can
    #   assert byte-for-byte pass-through on save
    module EditableFixtureBuilder
      MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
      PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
      WORKBOOK_REL_TYPE = "#{REL_NS}/officeDocument".freeze
      WORKSHEET_REL_TYPE = "#{REL_NS}/worksheet".freeze
      STYLES_REL_TYPE = "#{REL_NS}/styles".freeze
      SHARED_STRINGS_REL_TYPE = "#{REL_NS}/sharedStrings".freeze

      module_function

      def build(path)
        Zip::OutputStream.open(path) do |out|
          write_entry(out, "[Content_Types].xml", content_types_xml)
          write_entry(out, "_rels/.rels", root_rels_xml)
          write_entry(out, "xl/workbook.xml", workbook_xml)
          write_entry(out, "xl/_rels/workbook.xml.rels", workbook_rels_xml)
          write_entry(out, "xl/styles.xml", styles_xml)
          write_entry(out, "xl/sharedStrings.xml", shared_strings_xml)
          write_entry(out, "xl/worksheets/sheet1.xml", sheet1_xml)
          write_entry(out, "xl/worksheets/sheet2.xml", sheet2_xml)
        end
        path
      end

      def write_entry(out, name, body)
        out.put_next_entry(name)
        out.write(body)
      end

      def content_types_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Types xmlns="#{CONTENT_TYPES_NS}">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
            <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
            <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          </Types>
        XML
      end

      def root_rels_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PACKAGE_REL_NS}">
            <Relationship Id="rId1" Type="#{WORKBOOK_REL_TYPE}" Target="xl/workbook.xml"/>
          </Relationships>
        XML
      end

      def workbook_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <workbook xmlns="#{MAIN_NS}" xmlns:r="#{REL_NS}">
            <sheets>
              <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
              <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
            </sheets>
          </workbook>
        XML
      end

      def workbook_rels_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="#{PACKAGE_REL_NS}">
            <Relationship Id="rId1" Type="#{WORKSHEET_REL_TYPE}" Target="worksheets/sheet1.xml"/>
            <Relationship Id="rId2" Type="#{WORKSHEET_REL_TYPE}" Target="worksheets/sheet2.xml"/>
            <Relationship Id="rId3" Type="#{STYLES_REL_TYPE}" Target="styles.xml"/>
            <Relationship Id="rId4" Type="#{SHARED_STRINGS_REL_TYPE}" Target="sharedStrings.xml"/>
          </Relationships>
        XML
      end

      # Two xfs: index 0 is the default (no formatting), index 1 carries a
      # bold font reference so we can verify that overwriting an existing
      # cell preserves the +s="1"+ attribute.
      def styles_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <styleSheet xmlns="#{MAIN_NS}">
            <fonts count="2">
              <font><sz val="11"/><name val="Calibri"/></font>
              <font><b/><sz val="11"/><name val="Calibri"/></font>
            </fonts>
            <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
            <borders count="1"><border/></borders>
            <cellStyleXfs count="1"><xf numFmtId="0" fontId="0"/></cellStyleXfs>
            <cellXfs count="2">
              <xf numFmtId="0" fontId="0" xfId="0"/>
              <xf numFmtId="0" fontId="1" xfId="0" applyFont="1"/>
            </cellXfs>
          </styleSheet>
        XML
      end

      def shared_strings_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <sst xmlns="#{MAIN_NS}" count="2" uniqueCount="2">
            <si><t>Header A</t></si>
            <si><t>Header B</t></si>
          </sst>
        XML
      end

      # Sheet1 covers:
      #
      # * shared-string cell with a non-default style (A1: t="s" s="1")
      # * shared-string cell (B1: t="s")
      # * numeric cell (C1)
      # * numeric cell (A2)
      # * inline-string cell (B2: t="inlineStr")
      # * boolean cell (C2: t="b")
      # * sparse row (no row 3 / row 4) so we can test row insertion order
      # * lonely row 5 (A5)
      def sheet1_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <worksheet xmlns="#{MAIN_NS}" xmlns:r="#{REL_NS}">
            <dimension ref="A1:C5"/>
            <sheetData>
              <row r="1">
                <c r="A1" t="s" s="1"><v>0</v></c>
                <c r="B1" t="s"><v>1</v></c>
                <c r="C1"><v>100</v></c>
              </row>
              <row r="2">
                <c r="A2"><v>1</v></c>
                <c r="B2" t="inlineStr"><is><t>alpha</t></is></c>
                <c r="C2" t="b"><v>1</v></c>
              </row>
              <row r="5">
                <c r="A5"><v>99</v></c>
              </row>
            </sheetData>
          </worksheet>
        XML
      end

      def sheet2_xml
        <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <worksheet xmlns="#{MAIN_NS}" xmlns:r="#{REL_NS}">
            <dimension ref="A1:A1"/>
            <sheetData>
              <row r="1">
                <c r="A1"><v>42</v></c>
              </row>
            </sheetData>
          </worksheet>
        XML
      end
    end
  end
end
