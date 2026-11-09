require "minitest/autorun"
require "tmpdir"
require "zip"
require_relative "../lib/rbxl"
require_relative "editable_fixture_builder"

# Edit-mode features for template fill-in: styling, dates, write_rows,
# defined names, add_sheet, recalculation, and repeated saves.
class EditableTemplateTest < Minitest::Test
  B = Rbxl::Test::EditableFixtureBuilder
  MAIN_NS = B::MAIN_NS
  REL_NS = B::REL_NS
  CALC_CHAIN_REL_TYPE = "#{REL_NS}/calcChain".freeze

  # cellXfs index of the template's custom +yyyy/mm/dd+ style (see styles_xml).
  DATE_XF = 2

  # ---------- styles: patch on top of the template ----------

  def test_style_patch_keeps_template_font_and_changes_only_color
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        cell = book.sheet("Sheet1")["A1"] # s="1" → bold font
        assert_equal 1, cell.style_id
        cell.style(font: { color: "C00000" })
        assert_equal 3, cell.style_id, "derived xf is appended after the template's three"
        book.save(out)
      end

      styles = xml(out, "xl/styles.xml")
      fonts = styles.xpath("//m:fonts/m:font", "m" => MAIN_NS)
      assert_equal "3", styles.at_xpath("//m:fonts", "m" => MAIN_NS)["count"]
      new_font = fonts[2]
      refute_nil new_font.at_xpath("m:b", "m" => MAIN_NS), "template bold must survive"
      assert_equal "FFC00000", new_font.at_xpath("m:color", "m" => MAIN_NS)["rgb"]
      assert_equal "Calibri", new_font.at_xpath("m:name", "m" => MAIN_NS)["val"]

      xfs = styles.xpath("//m:cellXfs/m:xf", "m" => MAIN_NS)
      assert_equal 4, xfs.length
      assert_equal "4", styles.at_xpath("//m:cellXfs", "m" => MAIN_NS)["count"]
      assert_equal "2", xfs[3]["fontId"]
      assert_equal "1", xfs[3]["applyFont"]
      assert_equal "1", xfs[1]["fontId"], "existing xf must not change"
    end
  end

  def test_style_patch_is_shared_between_cells_with_same_base
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet["A1"].style(fill: { color: "FFF2CC" })
        sheet["A2"].style_id = 1
        sheet["A2"].style(fill: { color: "FFF2CC" })
        assert_equal sheet["A1"].style_id, sheet["A2"].style_id
        assert_equal 4, book.styles.xf_count
      end
    end
  end

  def test_style_patch_can_turn_bold_off_and_change_font_name
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A1"].style(font: { bold: false, name: "Yu Gothic UI" })
        book.save(out)
      end
      font = xml(out, "xl/styles.xml").xpath("//m:fonts/m:font", "m" => MAIN_NS).last
      assert_nil font.at_xpath("m:b", "m" => MAIN_NS)
      assert_equal "Yu Gothic UI", font.at_xpath("m:name", "m" => MAIN_NS)["val"]
      assert_nil font.at_xpath("m:scheme", "m" => MAIN_NS), "scheme would override the explicit name"
    end
  end

  def test_style_patch_merges_border_sides_and_alignment
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        cell = book.sheet("Sheet1")["B1"]
        cell.style_id = 2 # left thin border + centered
        cell.style(border: { bottom: { style: "double", color: "000000" } }, alignment: { vertical: "top" })
        book.save(out)
      end

      styles = xml(out, "xl/styles.xml")
      border = styles.xpath("//m:borders/m:border", "m" => MAIN_NS).last
      assert_equal "thin", border.at_xpath("m:left", "m" => MAIN_NS)["style"]
      assert_equal "double", border.at_xpath("m:bottom", "m" => MAIN_NS)["style"]
      assert_equal %w[left right top bottom diagonal], border.element_children.map(&:name)

      alignment = styles.xpath("//m:cellXfs/m:xf", "m" => MAIN_NS).last.at_xpath("m:alignment", "m" => MAIN_NS)
      assert_equal "center", alignment["horizontal"]
      assert_equal "top", alignment["vertical"]
    end
  end

  def test_number_format_reuses_builtin_then_existing_then_allocates
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet["C1"].style(number_format: "#,##0")
        sheet["C2"].style(number_format: '0.0"件"')
        sheet["C3"].style(number_format: "0.000")
        book.save(out)
      end

      styles = xml(out, "xl/styles.xml")
      xfs = styles.xpath("//m:cellXfs/m:xf", "m" => MAIN_NS)
      assert_equal %w[3 165 166], xfs.to_a.last(3).map { |x| x["numFmtId"] }
      fmts = styles.xpath("//m:numFmts/m:numFmt", "m" => MAIN_NS)
      assert_equal "3", styles.at_xpath("//m:numFmts", "m" => MAIN_NS)["count"]
      assert_equal "0.000", fmts.last["formatCode"]
    end
  end

  def test_add_style_and_style_id_assignment
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        id = book.add_style(font: { bold: true, color: "333333" }, fill: { color: "EEEEEE" })
        assert_equal 3, id
        book.sheet("Sheet1")["E9"].style_id = id
        assert_raises(ArgumentError) { book.sheet("Sheet1")["E9"].style_id = 99 }
        book.save(out)
      end

      assert_includes read_entry(out, "xl/worksheets/sheet1.xml"), %(<c r="E9" s="3"/>)
      xf = xml(out, "xl/styles.xml").xpath("//m:cellXfs/m:xf", "m" => MAIN_NS).last
      assert_equal "1", xf["applyFill"]
    end
  end

  def test_styles_part_untouched_without_style_operations
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = "x"
        book.save(out)
      end
      assert_equal read_entry(path, "xl/styles.xml"), read_entry(out, "xl/styles.xml")
    end
  end

  # ---------- dates ----------

  def test_date_round_trips_through_date_conversion
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet["A5"].value = Date.new(2026, 11, 10)
        sheet["B5"].value = Time.new(2026, 11, 10, 9, 30, 0)
        sheet["C5"].value = DateTime.new(1900, 3, 1, 12, 0, 0)
        book.save(out)
      end

      values = read_row(out, "Sheet1", 5, date_conversion: true)
      assert_equal Date.new(2026, 11, 10), values[0]
      assert_equal Time.new(2026, 11, 10, 9, 30, 0), values[1]
      assert_equal Time.new(1900, 3, 1, 12, 0, 0), values[2]
      assert_equal 46_336, read_row(out, "Sheet1", 5)[0]
    end
  end

  def test_date_in_1904_workbook
    with_template(date1904: true) do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A5"].value = Date.new(2026, 11, 10)
        book.save(out)
      end
      assert_equal 46_336 - 1462, read_row(out, "Sheet1", 5)[0]
      assert_equal Date.new(2026, 11, 10), read_row(out, "Sheet1", 5, date_conversion: true)[0]
    end
  end

  def test_date_keeps_template_date_style
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        cell = book.sheet("Sheet1")["D2"] # custom yyyy/mm/dd xf
        cell.value = Date.new(2026, 1, 2)
        assert_equal DATE_XF, cell.style_id
        book.save(out)
      end
      assert_equal read_entry(path, "xl/styles.xml"), read_entry(out, "xl/styles.xml")
    end
  end

  def test_date_on_styled_cell_derives_date_format_keeping_font
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        cell = book.sheet("Sheet1")["A1"] # bold, General
        cell.value = Date.new(2026, 1, 2)
        assert_equal 3, cell.style_id
        book.sheet("Sheet1")["A2"].value = Time.new(2026, 1, 2, 3, 4, 5)
        book.save(out)
      end
      xfs = xml(out, "xl/styles.xml").xpath("//m:cellXfs/m:xf", "m" => MAIN_NS)
      assert_equal "14", xfs[3]["numFmtId"]
      assert_equal "1", xfs[3]["fontId"], "template bold font must survive"
      assert_equal "22", xfs[4]["numFmtId"]
    end
  end

  def test_date_before_epoch_raises_without_touching_cell
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        cell = book.sheet("Sheet1")["A2"]
        assert_raises(Rbxl::EditableCellTypeError) { cell.value = Date.new(1899, 12, 1) }
        assert_equal 1, cell.value
        refute book.sheet("Sheet1").dirty?
      end
    end
  end

  # ---------- write_rows ----------

  def test_write_rows_fills_consecutive_cells_keeping_styles
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet.write_rows("A10:C12", [["作業費", 1, 50_000], ["交通費", 1, 3_200]])
        assert_equal 2, sheet["C10"].style_id, "template line-item style is kept"
        book.save(out)
      end
      assert_equal ["作業費", 1, 50_000], read_row(out, "Sheet1", 10)
      assert_equal ["交通費", 1, 3_200], read_row(out, "Sheet1", 11)
    end
  end

  def test_write_rows_with_plain_anchor_and_wide_columns
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet.write_rows("Y20", [[1, 2, 3]])
        assert_equal [1, 2, 3], [sheet["Y20"].value, sheet["Z20"].value, sheet["AA20"].value]
      end
    end
  end

  def test_write_rows_rejects_overflow_before_touching_anything
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        err = assert_raises(ArgumentError) { sheet.write_rows("A10:C11", [[1], [2], [3]]) }
        assert_includes err.message, "2 rows"
        assert_raises(ArgumentError) { sheet.write_rows("A10:B11", [[1, 2, 3]]) }
        assert_raises(ArgumentError) { sheet.write_rows("XFD1", [[1, 2]]) }
        assert_raises(ArgumentError) { sheet.write_rows("C3:A1", [[1]]) }
        assert_raises(ArgumentError) { sheet.write_rows("A1", [1, 2]) }
        refute sheet.dirty?
      end
    end
  end

  # ---------- defined names ----------

  def test_defined_name_resolves_global_name_case_insensitively
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        assert_includes book.defined_names, "顧客名"
        cell = book.defined_name("顧客名")
        assert_equal "B3", cell.coordinate
        cell.value = "株式会社Acme"
        book.defined_name("total").value = 1 # stored as "Total"
        book.save(out)
      end
      assert_equal "株式会社Acme", edit_value(out, "Sheet1", "B3")
      assert_equal 1, edit_value(out, "Sheet1", "E1")
    end
  end

  def test_defined_name_quoted_sheet_and_local_scope
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        assert_equal "A1", book.defined_name("note").coordinate
        assert_equal "O'Brien", book.sheets.to_a.last.name

        assert_equal "B2", book.defined_name("Local", sheet: "Sheet1").coordinate
        assert_equal "C2", book.defined_name("Local", sheet: 1).coordinate
        err = assert_raises(Rbxl::DefinedNameError) { book.defined_name("Local") }
        assert_includes err.message, "sheet:"
      end
    end
  end

  def test_defined_name_errors
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        assert_raises(Rbxl::DefinedNameError) { book.defined_name("missing") }
        err = assert_raises(Rbxl::DefinedNameError) { book.defined_name("明細") }
        assert_includes err.message, "not a single cell"
      end
    end
  end

  # ---------- add_sheet ----------

  def test_add_sheet_creates_readable_sheet
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        sheet = book.add_sheet("明細データ")
        sheet.write_rows("A1", [["品目", "金額"], ["作業費", 50_000]])
        sheet["C2"].value = Date.new(2026, 11, 10)
        assert_equal "明細データ", book.sheet_names.last
        assert_same sheet, book.sheet("明細データ")
        book.save(out)
      end

      Rbxl.open(out) do |book|
        assert_equal "明細データ", book.sheet_names.last
        rows = book.sheet("明細データ").rows(values_only: true).to_a
        assert_equal [["品目", "金額"], ["作業費", 50_000, 46_336]], rows
      end

      ct = read_entry(out, "[Content_Types].xml")
      assert_includes ct, %(PartName="/xl/worksheets/sheet3.xml")
      wb = read_entry(out, "xl/workbook.xml")
      assert_includes wb, %(<sheet name="明細データ" sheetId="4" r:id="rId10"/>)
      assert_includes read_entry(out, "xl/_rels/workbook.xml.rels"), %(Id="rId10")
    end
  end

  def test_add_sheet_validates_names
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        assert_raises(ArgumentError) { book.add_sheet("sheet1") }
        assert_raises(ArgumentError) { book.add_sheet("") }
        assert_raises(ArgumentError) { book.add_sheet("a" * 32) }
        assert_raises(ArgumentError) { book.add_sheet("a/b") }
        assert_raises(ArgumentError) { book.add_sheet("'quoted") }
        assert_raises(ArgumentError) { book.add_sheet(:sym) }
      end
    end
  end

  def test_add_sheet_survives_repeated_in_place_saves
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        book.add_sheet("New")["A1"].value = 1
        book.save
        book.sheet("New")["A2"].value = 2
        book.save
      end
      assert_equal [[1], [2]], Rbxl.open(path) { |b| b.sheet("New").rows(values_only: true).to_a }
      assert_equal 1, Zip::File.open(path) { |z| z.entries.count { |e| e.name == "xl/worksheets/sheet3.xml" } }
    end
  end

  # ---------- recalculation / calcChain ----------

  def test_edit_requests_full_calc_on_load_after_defined_names
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A2"].value = 5
        book.save(out)
      end
      root = xml(out, "xl/workbook.xml").root
      assert_equal %w[sheets definedNames calcPr], root.element_children.map(&:name)
      assert_equal "1", root.at_xpath("m:calcPr", "m" => MAIN_NS)["fullCalcOnLoad"]
    end
  end

  def test_reserialized_parts_keep_source_layout
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A2"].value = 5
        book.save(out)
      end
      expected = read_entry(path, "xl/workbook.xml").sub("</workbook>", %(<calcPr fullCalcOnLoad="1"/></workbook>))
      assert_equal expected, read_entry(out, "xl/workbook.xml")
      assert_equal read_entry(path, "xl/worksheets/sheet1.xml").sub("<c r=\"A2\"><v>1</v></c>", "<c r=\"A2\"><v>5</v></c>"),
                   read_entry(out, "xl/worksheets/sheet1.xml")
    end
  end

  def test_recalc_false_leaves_workbook_xml_untouched
    with_template do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A2"].value = 5
        book.save(out, recalc: false)
      end
      assert_equal read_entry(path, "xl/workbook.xml"), read_entry(out, "xl/workbook.xml")
    end
  end

  def test_existing_calc_pr_gets_flag_added
    with_template(calc_pr: true) do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A2"].value = 5
        book.save(out)
      end
      calc_pr = xml(out, "xl/workbook.xml").at_xpath("//m:calcPr", "m" => MAIN_NS)
      assert_equal "191029", calc_pr["calcId"]
      assert_equal "1", calc_pr["fullCalcOnLoad"]
    end
  end

  def test_calc_chain_dropped_only_when_formula_overwritten
    with_template do |path, dir|
      kept = File.join(dir, "kept.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A2"].value = 5
        book.save(kept)
      end
      assert Zip::File.open(kept) { |z| z.find_entry("xl/calcChain.xml") }

      dropped = File.join(dir, "dropped.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["D1"].value = 0 # was =SUM(A1:C1)
        book.save(dropped)
      end
      assert_nil Zip::File.open(dropped) { |z| z.find_entry("xl/calcChain.xml") }
      refute_includes read_entry(dropped, "[Content_Types].xml"), "calcChain"
      refute_includes read_entry(dropped, "xl/_rels/workbook.xml.rels"), "calcChain"
    end
  end

  # ---------- repeated saves (1.4.0 regressions) ----------

  def test_second_save_to_other_path_keeps_first_edits
    with_template do |path, dir|
      first = File.join(dir, "first.xlsx")
      second = File.join(dir, "second.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A2"].value = 111
        book.save(first)
        book.sheet(1)["A1"].value = "x"
        book.save(second)
      end
      assert_equal 111, read_row(second, "Sheet1", 2)[0]
      assert_equal ["x"], read_row(second, "Sheet2", 1)
    end
  end

  def test_in_place_save_then_touching_unloaded_sheet
    with_template do |path, _dir|
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A2"].value = 222
        book.save
        book.sheet(1)["A1"].value = "y"
        book.save
      end
      assert_equal 222, read_row(path, "Sheet1", 2)[0]
      assert_equal ["y"], read_row(path, "Sheet2", 1)
    end
  end

  # ---------- prefixed SpreadsheetML (OpenXML SDK style) ----------

  def test_inserted_nodes_take_the_documents_prefix
    with_template(prefixed: true) do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet["B7"].value = "new"
        sheet["B7"].style(font: { color: "FF0000" })
        book.save(out)
      end
      sheet_xml = read_entry(out, "xl/worksheets/sheet1.xml")
      assert_includes sheet_xml, %(<x:row r="7"><x:c r="B7" t="inlineStr")
      assert_includes sheet_xml, "<x:is><x:t>new</x:t></x:is>"
      refute_match(/<(row|c|is|t|v)[ >]/, sheet_xml)
      refute_match(/<(xf|font)[ >]/, read_entry(out, "xl/styles.xml"))
      assert_equal "new", edit_value(out, "Sheet1", "B7")
    end
  end

  # ---------- workbook without a styles part ----------

  def test_date_write_creates_styles_part_when_missing
    with_template(no_styles: true) do |path, dir|
      out = File.join(dir, "out.xlsx")
      Rbxl.open(path, edit: true) do |book|
        book.sheet("Sheet1")["A5"].value = Date.new(2026, 11, 10)
        book.save(out)
      end
      assert_includes read_entry(out, "[Content_Types].xml"), %(PartName="/xl/styles.xml")
      assert_includes read_entry(out, "xl/_rels/workbook.xml.rels"), %(Target="styles.xml")
      assert_equal Date.new(2026, 11, 10), read_row(out, "Sheet1", 5, date_conversion: true)[0]
    end
  end

  private

  def with_template(**opts)
    Dir.mktmpdir do |dir|
      path = File.join(dir, "template.xlsx")
      build_template(path, **opts)
      yield path, dir
    end
  end

  def build_template(path, date1904: false, calc_pr: false, prefixed: false, no_styles: false)
    parts = {
      "[Content_Types].xml" => content_types_xml(no_styles),
      "_rels/.rels" => B.root_rels_xml,
      "xl/workbook.xml" => workbook_xml(date1904, calc_pr),
      "xl/_rels/workbook.xml.rels" => workbook_rels_xml(no_styles),
      "xl/sharedStrings.xml" => B.shared_strings_xml,
      "xl/worksheets/sheet1.xml" => sheet1_xml,
      "xl/worksheets/sheet2.xml" => B.sheet2_xml,
      "xl/worksheets/sheet3.xml" => B.sheet2_xml,
      "xl/calcChain.xml" => %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="#{MAIN_NS}"><c r="D1" i="1"/></calcChain>)
    }
    parts["xl/styles.xml"] = styles_xml unless no_styles
    parts.transform_values! { |x| prefix_main(x) } if prefixed

    Zip::OutputStream.open(path) do |out|
      parts.each do |name, body|
        out.put_next_entry(name)
        out.write(body)
      end
    end
  end

  def prefix_main(xml)
    return xml unless xml.include?(%(xmlns="#{MAIN_NS}"))

    xml.gsub(%(xmlns="#{MAIN_NS}"), %(xmlns:x="#{MAIN_NS}"))
       .gsub(%r{<(/?)(?!\?)([A-Za-z]+)([ />])}) { "<#{$1}x:#{$2}#{$3}" }
  end

  def content_types_xml(no_styles)
    xml = B.content_types_xml.sub(
      "</Types>",
      %(  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n) +
      %(  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>\n</Types>)
    )
    no_styles ? xml.sub(%r{^.*/xl/styles.xml.*\n}, "") : xml
  end

  def workbook_rels_xml(no_styles)
    xml = B.workbook_rels_xml.sub(
      "</Relationships>",
      %(  <Relationship Id="rId5" Type="#{B::WORKSHEET_REL_TYPE}" Target="worksheets/sheet3.xml"/>\n) +
      %(  <Relationship Id="rId9" Type="#{CALC_CHAIN_REL_TYPE}" Target="calcChain.xml"/>\n</Relationships>)
    )
    xml = xml.sub("rId9", "rId3").sub(%r{^.*Id="rId3" Type="[^"]*/styles".*\n}, "") if no_styles
    xml
  end

  def workbook_xml(date1904, calc_pr)
    <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook xmlns="#{MAIN_NS}" xmlns:r="#{REL_NS}">#{date1904 ? '<workbookPr date1904="1"/>' : ''}
        <sheets>
          <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
          <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
          <sheet name="O'Brien" sheetId="3" r:id="rId5"/>
        </sheets>
        <definedNames>
          <definedName name="顧客名">Sheet1!$B$3</definedName>
          <definedName name="Total">Sheet1!$E$1:$E$1</definedName>
          <definedName name="note">'O''Brien'!$A$1</definedName>
          <definedName name="Local" localSheetId="0">Sheet1!$B$2</definedName>
          <definedName name="Local" localSheetId="1">Sheet2!$C$2</definedName>
          <definedName name="明細">Sheet1!$A$10:$C$12</definedName>
        </definedNames>#{calc_pr ? '<calcPr calcId="191029"/>' : ''}
      </workbook>
    XML
  end

  # xf 0: default; xf 1: bold font; xf 2: custom yyyy/mm/dd date + left
  # thin border + centered.
  def styles_xml
    <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <styleSheet xmlns="#{MAIN_NS}">
        <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy/mm/dd"/></numFmts>
        <fonts count="2">
          <font><sz val="11"/><name val="Calibri"/><scheme val="minor"/></font>
          <font><b/><sz val="11"/><name val="Calibri"/></font>
        </fonts>
        <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        <borders count="2">
          <border><left/><right/><top/><bottom/><diagonal/></border>
          <border><left style="thin"/><right/><top/><bottom/><diagonal/></border>
        </borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="3">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
          <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
          <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="center"/></xf>
        </cellXfs>
        <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
      </styleSheet>
    XML
  end

  def sheet1_xml
    <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="#{MAIN_NS}" xmlns:r="#{REL_NS}">
        <dimension ref="A1:D12"/>
        <sheetData>
          <row r="1">
            <c r="A1" t="s" s="1"><v>0</v></c>
            <c r="B1"><v>10</v></c>
            <c r="C1"><v>100</v></c>
            <c r="D1"><f>SUM(A1:C1)</f><v>110</v></c>
          </row>
          <row r="2">
            <c r="A2"><v>1</v></c>
            <c r="D2" s="2"/>
          </row>
          <row r="10"><c r="A10" s="1"/><c r="B10"/><c r="C10" s="2"/></row>
          <row r="11"><c r="A11" s="1"/><c r="B11"/><c r="C11" s="2"/></row>
        </sheetData>
      </worksheet>
    XML
  end

  def xml(path, entry)
    Nokogiri::XML(read_entry(path, entry))
  end

  def read_entry(path, name)
    Zip::File.open(path) { |z| z.find_entry(name).get_input_stream.read.force_encoding("UTF-8") }
  end

  def edit_value(path, sheet, coordinate)
    Rbxl.open(path, edit: true) { |book| book.sheet(sheet)[coordinate].value }
  end

  def read_row(path, sheet, row_number, date_conversion: false)
    Rbxl.open(path, date_conversion: date_conversion) do |book|
      row = book.sheet(sheet).rows.find { |r| r.index == row_number }
      row&.values
    end
  end
end
