require "minitest/autorun"
require "pathname"
require "tmpdir"
require "zip"
require_relative "../lib/rbxl"

class RbxlTest < Minitest::Test
  def test_open_defaults_to_read_only
    Dir.mktmpdir do |dir|
      path = File.join(dir, "report.xlsx")
      book = Rbxl.new
      book.add_sheet("S") << ["ok"]
      book.save(path)

      loaded = Rbxl.open(path)
      assert_equal [["ok"]], loaded.sheet("S").rows(values_only: true).to_a
      loaded.close
    end
  end

  def test_open_rejects_read_only_false
    assert_raises(NotImplementedError) { Rbxl.open("dummy.xlsx", read_only: false) }
  end

  def test_new_rejects_write_only_false
    assert_raises(NotImplementedError) { Rbxl.new(write_only: false) }
  end

  def test_write_only_then_read_only_round_trip
    Dir.mktmpdir do |dir|
      path = File.join(dir, "report.xlsx")

      book = Rbxl.new(write_only: true)
      sheet = book.add_sheet("Report")
      sheet.append(["id", "name", "active", "score"])
      sheet.append([1, "alice", true, 10.5])
      sheet.append([2, "bob", false, 7])
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      assert_equal ["Report"], loaded.sheet_names
      assert_equal 4, loaded.sheet("Report").max_column
      assert_equal 3, loaded.sheet("Report").max_row
      assert_equal "A1:D3", loaded.sheet("Report").calculate_dimension

      rows = loaded.sheet("Report").rows.map(&:values)
      assert_equal [
        ["id", "name", "active", "score"],
        [1, "alice", true, 10.5],
        [2, "bob", false, 7]
      ], rows

      values = loaded.sheet("Report").rows(values_only: true).to_a
      assert_equal rows, values

      loaded.close
      assert loaded.closed?
      assert_raises(Rbxl::ClosedWorkbookError) { loaded.sheet("Report") }
    end
  end

  def test_sheet_accepts_integer_index
    Dir.mktmpdir do |dir|
      path = File.join(dir, "multi.xlsx")

      book = Rbxl.new
      book.add_sheet("First").append(["a"])
      book.add_sheet("Second").append(["b"])
      book.add_sheet("Third").append(["c"])
      book.save(path)

      loaded = Rbxl.open(path)
      assert_equal "First", loaded.sheet(0).name
      assert_equal "Third", loaded.sheet(-1).name
      assert_equal "Second", loaded.sheet(1).name
      assert_equal [["a"]], loaded.sheet(0).rows(values_only: true).to_a
      loaded.close
    end
  end

  def test_sheets_iterates_in_workbook_order
    Dir.mktmpdir do |dir|
      path = File.join(dir, "multi.xlsx")

      book = Rbxl.new
      book.add_sheet("First").append(["a"])
      book.add_sheet("Second").append(["b"])
      book.add_sheet("Third").append(["c"])
      book.save(path)

      loaded = Rbxl.open(path)
      assert_equal ["First", "Second", "Third"], loaded.sheets.map(&:name)
      assert_equal "First", loaded.sheets.first.name
      assert_equal [["b"]], loaded.sheets.to_a[1].rows(values_only: true).to_a

      yielded = []
      loaded.sheets { |s| yielded << s.name }
      assert_equal ["First", "Second", "Third"], yielded

      loaded.close
    end
  end

  def test_sheets_raises_after_close
    Dir.mktmpdir do |dir|
      path = File.join(dir, "closed.xlsx")
      book = Rbxl.new
      book.add_sheet("Only").append(["x"])
      book.save(path)

      loaded = Rbxl.open(path)
      loaded.close

      assert_raises(Rbxl::ClosedWorkbookError) { loaded.sheets }
    end
  end

  def test_sheet_integer_index_out_of_range_raises
    Dir.mktmpdir do |dir|
      path = File.join(dir, "small.xlsx")

      book = Rbxl.new
      book.add_sheet("Only").append(["x"])
      book.save(path)

      loaded = Rbxl.open(path)
      err = assert_raises(Rbxl::SheetNotFoundError) { loaded.sheet(3) }
      assert_includes err.message, "out of range"
      assert_includes err.message, "3"
      loaded.close
    end
  end

  def test_missing_sheet_raises
    Dir.mktmpdir do |dir|
      path = File.join(dir, "book.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Only")
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      assert_raises(Rbxl::SheetNotFoundError) { loaded.sheet("Nope") }
      loaded.close
    end
  end

  def test_open_accepts_pathname
    Dir.mktmpdir do |dir|
      path = File.join(dir, "book.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Only").append([1])
      book.save(path)

      loaded = Rbxl.open(Pathname.new(path), read_only: true)
      assert_equal ["Only"], loaded.sheet_names
      loaded.close
    end
  end

  def test_dimensions_can_be_reset
    Dir.mktmpdir do |dir|
      path = File.join(dir, "sparse.xlsx")

      book = Rbxl.new(write_only: true)
      sheet = book.add_sheet("Sparse")
      sheet << ["a", "c"]
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      worksheet = loaded.sheet("Sparse")

      assert_equal 2, worksheet.max_column
      assert_equal 1, worksheet.max_row
      assert_equal "A1:B1", worksheet.calculate_dimension
      worksheet.reset_dimensions
      assert_nil worksheet.max_column
      assert_nil worksheet.max_row
      assert_raises(Rbxl::UnsizedWorksheetError) { worksheet.calculate_dimension }
      assert_equal "A1:B1", worksheet.calculate_dimension(force: true)
      loaded.close
    end
  end

  def test_write_only_workbook_can_only_save_once
    Dir.mktmpdir do |dir|
      path = File.join(dir, "report.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Report") << [Rbxl::WriteOnlyCell.new("x", style_id: 0)]
      book.save(path)

      assert_raises(Rbxl::WorkbookAlreadySavedError) { book.add_sheet("Another") }
      assert_raises(Rbxl::WorkbookAlreadySavedError) { book.save(path) }
    end
  end

  def test_append_rejects_non_row_iterables
    book = Rbxl.new(write_only: true)
    sheet = book.add_sheet("Report")

    assert_raises(TypeError) { sheet.append("not-a-row") }
  end

  def test_write_only_cell_round_trips_style_and_value
    Dir.mktmpdir do |dir|
      path = File.join(dir, "styled.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Styled") << [Rbxl::WriteOnlyCell.new("header", style_id: 3), 42]
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      row = loaded.sheet("Styled").each_row.first

      assert_equal ["header", 42], row.values
      assert_equal "A1", row[0].coordinate
      assert_instance_of Rbxl::ReadOnlyCell, row[0]
      loaded.close
    end
  end

  def test_writer_round_trips_escaped_strings
    Dir.mktmpdir do |dir|
      path = File.join(dir, "escaped.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Escaped").append(["&", "<", ">", "", "\"quoted\""])
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      assert_equal [["&", "<", ">", "", "\"quoted\""]], loaded.sheet("Escaped").rows(values_only: true).to_a
      loaded.close
    end
  end

  def test_writer_avoids_zip64_for_small_workbooks
    Dir.mktmpdir do |dir|
      path = File.join(dir, "small.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Bench").append(["a", "b", "c"])
      book.save(path)

      Zip::File.open(path) do |zip_file|
        zip_file.entries.each do |entry|
          refute_includes entry.extra.keys, :zip64
        end
      end
    end
  end

  def test_append_accepts_enumerator_rows
    Dir.mktmpdir do |dir|
      path = File.join(dir, "enum.xlsx")

      book = Rbxl.new(write_only: true)
      enum = ["a", "b", "c"].each
      book.add_sheet("Enum").append(enum)
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      assert_equal [["a", "b", "c"]], loaded.sheet("Enum").rows.map(&:values)
      loaded.close
    end
  end

  def test_values_only_rows_return_plain_values
    Dir.mktmpdir do |dir|
      path = File.join(dir, "values.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Values").append([1, "x", true])
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      row = loaded.sheet("Values").each_row(values_only: true).first

      assert_equal [1, "x", true], row
      loaded.close
    end
  end

  def test_values_only_with_pad_cells_returns_nil_for_missing_cells
    loaded = Rbxl.open(fixture_path("sparse.xlsx"), read_only: true)
    row = loaded.sheet("Sparse").each_row(pad_cells: true, values_only: true).first

    assert_equal ["left", nil, "right"], row
    loaded.close
  end

  def test_reader_supports_padded_cells_for_missing_coordinates
    loaded = Rbxl.open(fixture_path("sparse.xlsx"), read_only: true)
    row = loaded.sheet("Sparse").each_row(pad_cells: true).first

    assert_equal ["left", nil, "right"], row.values
    assert_instance_of Rbxl::EmptyCell, row[1]
    assert_equal "B1", row[1].coordinate
    loaded.close
  end

  def test_open_with_block_yields_workbook_and_auto_closes
    Dir.mktmpdir do |dir|
      path = File.join(dir, "block.xlsx")

      book = Rbxl.new
      book.add_sheet("S") << ["ok"]
      book.save(path)

      captured = nil
      Rbxl.open(path) do |loaded|
        captured = loaded
        refute loaded.closed?
        assert_equal [["ok"]], loaded.sheet("S").rows(values_only: true).to_a
      end

      assert_instance_of Rbxl::ReadOnlyWorkbook, captured
      assert captured.closed?
    end
  end

  def test_open_with_block_closes_on_exception
    Dir.mktmpdir do |dir|
      path = File.join(dir, "block_raise.xlsx")

      book = Rbxl.new
      book.add_sheet("S") << ["ok"]
      book.save(path)

      captured = nil
      assert_raises(RuntimeError) do
        Rbxl.open(path) do |loaded|
          captured = loaded
          raise "boom"
        end
      end

      assert captured.closed?
    end
  end

  def test_open_rejects_legacy_xls_with_helpful_message
    Dir.mktmpdir do |dir|
      path = File.join(dir, "legacy.xls")
      File.binwrite(path, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b + ("\x00".b * 512))

      err = assert_raises(Rbxl::UnsupportedFormatError) { Rbxl.open(path) }
      assert_includes err.message, path
      assert_includes err.message, "legacy .xls"
      assert_includes err.message, "--convert-to xlsx"
    end
  end

  def test_open_rejects_non_zip_with_helpful_message
    Dir.mktmpdir do |dir|
      path = File.join(dir, "garbage.xlsx")
      File.binwrite(path, "not an xlsx at all")

      err = assert_raises(Rbxl::UnsupportedFormatError) { Rbxl.open(path) }
      assert_includes err.message, path
      assert_includes err.message, "ZIP signature"
    end
  end

  def test_open_with_block_returns_block_value
    Dir.mktmpdir do |dir|
      path = File.join(dir, "block_return.xlsx")

      book = Rbxl.new
      book.add_sheet("S") << ["ok"]
      book.save(path)

      result = Rbxl.open(path) { |loaded| loaded.sheet_names }
      assert_equal ["S"], result
    end
  end

  def test_close_is_idempotent
    Dir.mktmpdir do |dir|
      path = File.join(dir, "report.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("Report").append([1])
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)
      loaded.close
      loaded.close

      assert loaded.closed?
    end
  end

  def test_multiple_sheets_preserve_order_and_can_be_read_individually
    Dir.mktmpdir do |dir|
      path = File.join(dir, "multi.xlsx")

      book = Rbxl.new(write_only: true)
      book.add_sheet("First").append(["a"])
      book.add_sheet("Second").append(["b"])
      book.add_sheet("Third").append(["c"])
      book.save(path)

      loaded = Rbxl.open(path, read_only: true)

      assert_equal ["First", "Second", "Third"], loaded.sheet_names
      assert_equal [["a"]], loaded.sheet("First").rows(values_only: true).to_a
      assert_equal [["b"]], loaded.sheet("Second").rows(values_only: true).to_a
      assert_equal [["c"]], loaded.sheet("Third").rows(values_only: true).to_a
      loaded.close
    end
  end

  def test_shared_strings_can_read_empty_string
    loaded = Rbxl.open(fixture_path("shared_strings.xlsx"), read_only: true)
    rows = loaded.sheet("Strings").rows(values_only: true).to_a

    assert_equal [["alpha"], [""]], rows
    loaded.close
  end

  def test_shared_strings_ignore_phonetic_runs
    loaded = Rbxl.open(fixture_path("phonetic_shared_strings.xlsx"), read_only: true)
    rows = loaded.sheet("Phonetic").rows(values_only: true).to_a

    assert_equal [["東京駅"], ["青空"]], rows
    refute_includes rows.flatten.join(" "), "トウキョウ"
    refute_includes rows.flatten.join(" "), "アオ"
    loaded.close
  end

  def test_calculate_dimension_force_handles_sheet_without_dimension_node
    loaded = Rbxl.open(fixture_path("no_dimension.xlsx"), read_only: true)
    sheet = loaded.sheet("NoDimension")

    assert_nil sheet.max_column
    assert_nil sheet.max_row
    assert_raises(Rbxl::UnsizedWorksheetError) { sheet.calculate_dimension }
    assert_equal "A1:C2", sheet.calculate_dimension(force: true)
    loaded.close
  end

  def test_reader_supports_padded_sparse_rows_over_multiple_rows
    loaded = Rbxl.open(fixture_path("sparse_rows.xlsx"), read_only: true)
    rows = loaded.sheet("SparseRows").rows(values_only: true).to_a
    padded = loaded.sheet("SparseRows").each_row(pad_cells: true, values_only: true).to_a

    assert_equal [["top"], ["tail"]], rows
    assert_equal [["top", nil, nil], [nil, nil, "tail"]], padded
    loaded.close
  end

  def test_reader_supports_implicit_coordinates
    loaded = Rbxl.open(fixture_path("implicit_coordinates.xlsx"), read_only: true)
    rows = loaded.sheet("Implicit").rows(values_only: true).to_a

    assert_equal [["Test"], ["A2", "B2", "C2"]], rows
    assert_equal "A1:C2", loaded.sheet("Implicit").calculate_dimension
    loaded.close
  end

  def test_open_tolerates_unrelated_broken_relationship_entries
    loaded = Rbxl.open(fixture_path("file_item_error.xlsx"), read_only: true)

    assert_equal ["BrokenRels"], loaded.sheet_names
    assert_equal [["ok"]], loaded.sheet("BrokenRels").rows(values_only: true).to_a
    loaded.close
  end

  def test_expand_merged_fills_values_only_rows_without_changing_default_behavior
    loaded = Rbxl.open(fixture_path("merged_cells.xlsx"), read_only: true)
    sheet = loaded.sheet("Merged")

    assert_equal [["group", "solo"], ["tail"]], sheet.rows(values_only: true).to_a
    assert_equal [["group", nil, "solo", nil], [nil, nil, nil, "tail"]], sheet.rows(values_only: true, pad_cells: true).to_a
    assert_equal [["group", "group", "solo", nil], ["group", "group", "solo", "tail"]], sheet.rows(values_only: true, pad_cells: true, expand_merged: true).to_a

    loaded.close
  end

  def test_expand_merged_fills_cell_objects
    loaded = Rbxl.open(fixture_path("merged_cells.xlsx"), read_only: true)
    row = loaded.sheet("Merged").each_row(pad_cells: true, expand_merged: true).to_a.last

    assert_equal ["group", "group", "solo", "tail"], row.values
    assert_equal "A2", row[0].coordinate
    assert_equal "B2", row[1].coordinate
    assert_equal "C2", row[2].coordinate

    loaded.close
  end

  def test_each_row_handles_self_closing_rows_and_cells_in_ruby_path
    Dir.mktmpdir do |dir|
      path = File.join(dir, "selfclose.xlsx")
      write_minimal_workbook(path, <<~XML)
        <dimension ref="A1:B3"/><sheetData>
          <row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c><c r="B1"/></row>
          <row r="2"/>
          <row r="3"><c r="A3" t="inlineStr"><is><t>y</t></is></c></row>
        </sheetData>
      XML

      loaded = Rbxl.open(path)
      sheet = loaded.sheet("Sheet1")
      sheet.instance_variable_set(:@disable_native, true)
      full_rows = sheet.each_row.to_a
      assert_equal [1, 2, 3], full_rows.map(&:index)
      assert_equal [["x", nil], [], ["y"]], full_rows.map(&:values)

      sheet = loaded.sheet("Sheet1")
      sheet.instance_variable_set(:@disable_native, true)
      value_rows = sheet.each_row(values_only: true).to_a
      assert_equal [["x", nil], [], ["y"]], value_rows
      refute_includes value_rows, nil

      loaded.close
    end
  end

  def test_date_conversion_returns_time_and_date_objects
    require "date"
    Dir.mktmpdir do |dir|
      path = File.join(dir, "dates.xlsx")
      styles = <<~XML
        <numFmts count="1"><numFmt numFmtId="200" formatCode="yyyy/mm/dd hh:mm:ss"/></numFmts>
        <cellXfs count="4">
          <xf numFmtId="0"/>
          <xf numFmtId="14"/>
          <xf numFmtId="22"/>
          <xf numFmtId="200"/>
        </cellXfs>
      XML
      sheet_xml = <<~XML
        <dimension ref="A1:D1"/><sheetData>
          <row r="1">
            <c r="A1" s="0"><v>3.14</v></c>
            <c r="B1" s="1"><v>44562</v></c>
            <c r="C1" s="2"><v>44562.5</v></c>
            <c r="D1" s="3"><v>44562.75</v></c>
          </row>
        </sheetData>
      XML
      write_minimal_workbook(path, sheet_xml, styles: styles)

      loaded = Rbxl.open(path, date_conversion: true)
      row = loaded.sheet("Sheet1").rows(values_only: true).first
      assert_equal 3.14, row[0]
      assert_equal Date.new(2022, 1, 1), row[1]
      assert_kind_of Time, row[2]
      assert_equal Time.new(2022, 1, 1, 12, 0, 0), row[3] - (6 * 3600)
      loaded.close

      loaded_default = Rbxl.open(path)
      default_row = loaded_default.sheet("Sheet1").rows(values_only: true).first
      assert_equal [3.14, 44562, 44562.5, 44562.75], default_row
      loaded_default.close
    end
  end

  def test_date_conversion_honors_1904_date_system
    Dir.mktmpdir do |dir|
      path = File.join(dir, "dates_1904.xlsx")
      styles = <<~XML
        <cellXfs count="2">
          <xf numFmtId="0"/>
          <xf numFmtId="14"/>
        </cellXfs>
      XML
      sheet_xml = <<~XML
        <dimension ref="A1:B1"/><sheetData>
          <row r="1">
            <c r="A1" s="1"><v>0</v></c>
            <c r="B1" s="1"><v>1.5</v></c>
          </row>
        </sheetData>
      XML
      write_minimal_workbook(path, sheet_xml, styles: styles, workbook_pr: '<workbookPr date1904="1"/>')

      loaded = Rbxl.open(path, date_conversion: true)
      row = loaded.sheet("Sheet1").rows(values_only: true).first

      assert_equal Date.new(1904, 1, 1), row[0]
      assert_equal Time.new(1904, 1, 2, 12, 0, 0), row[1]
      loaded.close
    end
  end

  def test_invalid_workbook_xml_reports_path_and_entry
    Dir.mktmpdir do |dir|
      path = File.join(dir, "broken_workbook.xlsx")
      write_minimal_workbook(path, "<sheetData/>")

      Zip::File.open(path) do |zf|
        zf.get_output_stream("xl/workbook.xml") { |s| s.write("<workbook") }
      end

      err = assert_raises(Rbxl::WorkbookFormatError) { Rbxl.open(path) }
      assert_includes err.message, path
      assert_includes err.message, "xl/workbook.xml"
    end
  end

  def test_invalid_worksheet_xml_reports_path_and_sheet
    Dir.mktmpdir do |dir|
      path = File.join(dir, "broken_sheet.xlsx")
      write_minimal_workbook(path, "<sheetData/>")

      Zip::File.open(path) do |zf|
        zf.get_output_stream("xl/worksheets/sheet1.xml") { |s| s.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row') }
      end

      loaded = Rbxl.open(path)
      err = assert_raises(Rbxl::WorksheetFormatError) do
        loaded.sheet("Sheet1").rows(values_only: true).to_a
      end
      assert_includes err.message, path
      assert_includes err.message, "Sheet1"
      loaded.close
    end
  end

  def test_add_style_returns_sequential_one_based_ids
    book = Rbxl.new
    first = book.add_style(font: { bold: true })
    second = book.add_style(font: { italic: true })
    third = book.add_style(font: { bold: true })

    assert_equal 1, first
    assert_equal 2, second
    assert_equal 3, third
  end

  def test_workbook_without_add_style_emits_unchanged_styles_xml
    Dir.mktmpdir do |dir|
      path = File.join(dir, "plain.xlsx")
      book = Rbxl.new
      book.add_sheet("S") << ["a"]
      book.save(path)

      styles = nil
      Zip::File.open(path) do |zf|
        styles = zf.read("xl/styles.xml")
      end

      assert_includes styles, '<fonts count="1">'
      assert_includes styles, '<fills count="1">'
      assert_includes styles, '<cellXfs count="1">'
      refute_includes styles, "gray125"
    end
  end

  def test_styles_xml_dedupes_fonts_across_styles
    Dir.mktmpdir do |dir|
      path = File.join(dir, "deduped.xlsx")
      book = Rbxl.new
      red = { name: "Calibri", color: "FF0000" }
      blue = { name: "Calibri", color: "0000FF" }
      book.add_style(font: red)
      book.add_style(font: blue)
      book.add_style(font: red)
      book.add_sheet("S") << ["a"]
      book.save(path)

      styles = nil
      Zip::File.open(path) do |zf|
        styles = zf.read("xl/styles.xml")
      end

      # 1 default + 2 unique custom fonts = 3
      assert_includes styles, '<fonts count="3">'
      assert_includes styles, '<cellXfs count="4">'
    end
  end

  def test_styles_xml_includes_gray125_when_custom_fill_present
    Dir.mktmpdir do |dir|
      path = File.join(dir, "filled.xlsx")
      book = Rbxl.new
      book.add_style(fill: { color: "EEEEEE" })
      book.add_sheet("S") << ["a"]
      book.save(path)

      styles = nil
      Zip::File.open(path) do |zf|
        styles = zf.read("xl/styles.xml")
      end

      assert_includes styles, '<fills count="3">'
      assert_includes styles, "gray125"
      assert_includes styles, '<fgColor rgb="FFEEEEEE"/>'
    end
  end

  def test_styled_workbook_round_trips_values
    Dir.mktmpdir do |dir|
      path = File.join(dir, "styled-rt.xlsx")
      book = Rbxl.new
      header = book.add_style(
        font: { name: "Yu Gothic UI", size: 11, bold: true, color: "333333" },
        fill: { color: "EEEEEE" },
        alignment: { horizontal: "center" },
      )
      mono = book.add_style(font: { name: "Consolas", size: 11, color: "4A90D9" })
      sheet = book.add_sheet("Report")
      sheet << [Rbxl::WriteOnlyCell.new("Path", style_id: header), Rbxl::WriteOnlyCell.new("Description", style_id: header)]
      sheet << [Rbxl::WriteOnlyCell.new("/bin/run.rb", style_id: mono), "entry point"]
      book.save(path)

      loaded = Rbxl.open(path)
      values = loaded.sheet("Report").rows(values_only: true).to_a
      assert_equal [["Path", "Description"], ["/bin/run.rb", "entry point"]], values
      loaded.close
    end
  end

  def test_add_style_after_save_raises
    Dir.mktmpdir do |dir|
      path = File.join(dir, "x.xlsx")
      book = Rbxl.new
      book.add_sheet("S") << ["a"]
      book.save(path)

      assert_raises(Rbxl::WorkbookAlreadySavedError) { book.add_style(font: { bold: true }) }
    end
  end

  def test_add_style_with_number_format_allocates_custom_numfmt_id
    Dir.mktmpdir do |dir|
      path = File.join(dir, "nf.xlsx")
      book = Rbxl.new
      book.add_style(number_format: "#,##0")
      book.add_sheet("S") << ["a"]
      book.save(path)

      styles = nil
      Zip::File.open(path) do |zf|
        styles = zf.read("xl/styles.xml")
      end

      assert_includes styles, '<numFmts count="1">'
      assert_includes styles, '<numFmt numFmtId="164" formatCode="#,##0"/>'
    end
  end

  def test_column_width_emits_cols_block
    Dir.mktmpdir do |dir|
      path = File.join(dir, "cols.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.column_width(0, 45.7)
      sheet.column_width(1, 60.0)
      sheet << ["a", "b"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<col min="1" max="1" width="45.7" customWidth="1"/>'
      assert_includes sheet_xml, '<col min="2" max="2" width="60" customWidth="1"/>'
    end
  end

  def test_column_width_with_range_merges_consecutive_same_widths
    Dir.mktmpdir do |dir|
      path = File.join(dir, "cols-range.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.column_width(1..3, 12)
      sheet << ["a", "b", "c", "d"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<col min="2" max="4" width="12" customWidth="1"/>'
    end
  end

  def test_column_width_rejects_invalid_arguments
    book = Rbxl.new
    sheet = book.add_sheet("S")

    assert_raises(ArgumentError) { sheet.column_width(-1, 10) }
    assert_raises(ArgumentError) { sheet.column_width(0, 0) }
    assert_raises(ArgumentError) { sheet.column_width(0, -1) }
    assert_raises(ArgumentError) { sheet.column_width(2..1, 10) }
    assert_raises(TypeError) { sheet.column_width("A", 10) }
  end

  def test_freeze_emits_pane_with_top_left_and_active_pane
    Dir.mktmpdir do |dir|
      path = File.join(dir, "freeze.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.freeze!(rows: 1, cols: 2)
      sheet << ["a", "b", "c"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<sheetViews>'
      assert_includes sheet_xml, '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>'
    end
  end

  def test_freeze_with_only_rows_uses_bottom_left_active_pane
    Dir.mktmpdir do |dir|
      path = File.join(dir, "freeze-row.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.freeze!(rows: 1)
      sheet << ["a"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
    end
  end

  def test_freeze_clears_when_both_zero
    book = Rbxl.new
    sheet = book.add_sheet("S")
    sheet.freeze!(rows: 1)
    sheet.freeze!(rows: 0, cols: 0)
    sheet << ["a"]

    refute_includes sheet.send(:to_xml), '<sheetViews>'
  end

  def test_merge_emits_merge_cells_after_sheet_data
    Dir.mktmpdir do |dir|
      path = File.join(dir, "merge.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.merge("A1:C1")
      sheet.merge("A2", "B2")
      sheet << ["heading"]
      sheet << ["sub"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_match %r{</sheetData>.*<mergeCells count="2">.*<mergeCell ref="A1:C1"/>.*<mergeCell ref="A2:B2"/>.*</mergeCells>}m, sheet_xml
    end
  end

  def test_merge_rejects_bad_arguments
    book = Rbxl.new
    sheet = book.add_sheet("S")

    assert_raises(ArgumentError) { sheet.merge }
    assert_raises(ArgumentError) { sheet.merge(:not_a_string) }
    assert_raises(ArgumentError) { sheet.merge("A1", "B1", "C1") }
  end

  def test_worksheet_without_meta_features_emits_unchanged_xml
    Dir.mktmpdir do |dir|
      path = File.join(dir, "plain-sheet.xlsx")
      book = Rbxl.new
      book.add_sheet("S") << ["a", "b"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      refute_includes sheet_xml, '<cols>'
      refute_includes sheet_xml, '<sheetViews>'
      refute_includes sheet_xml, '<mergeCells'
    end
  end

  def test_meta_features_round_trip_through_read_only
    Dir.mktmpdir do |dir|
      path = File.join(dir, "round.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("Report")
      sheet.column_width(0, 30)
      sheet.column_width(1, 50)
      sheet.freeze!(rows: 1)
      sheet.merge("A1:B1")
      sheet << ["title"]
      sheet << ["a", "b"]
      book.save(path)

      loaded = Rbxl.open(path)
      values = loaded.sheet("Report").rows(values_only: true).to_a
      assert_equal [["title"], ["a", "b"]], values
      loaded.close
    end
  end

  def test_row_height_emits_ht_and_custom_height_attribute
    Dir.mktmpdir do |dir|
      path = File.join(dir, "rh.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.row_height(0, 22)
      sheet << ["header"]
      sheet << ["body"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<row r="1" ht="22" customHeight="1">'
      assert_includes sheet_xml, '<row r="2">'
    end
  end

  def test_row_height_with_range_applies_to_each_row
    Dir.mktmpdir do |dir|
      path = File.join(dir, "rh-range.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.row_height(0..2, 18.5)
      4.times { |i| sheet << ["row#{i}"] }
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<row r="1" ht="18.5" customHeight="1">'
      assert_includes sheet_xml, '<row r="2" ht="18.5" customHeight="1">'
      assert_includes sheet_xml, '<row r="3" ht="18.5" customHeight="1">'
      assert_includes sheet_xml, '<row r="4">'
    end
  end

  def test_row_height_unmatched_index_does_not_emit
    Dir.mktmpdir do |dir|
      path = File.join(dir, "rh-future.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.row_height(99, 30)
      sheet << ["only one row"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<row r="1">'
      refute_includes sheet_xml, "ht="
      refute_includes sheet_xml, "customHeight"
    end
  end

  def test_row_height_rejects_invalid_arguments
    book = Rbxl.new
    sheet = book.add_sheet("S")

    assert_raises(ArgumentError) { sheet.row_height(-1, 20) }
    assert_raises(ArgumentError) { sheet.row_height(0, 0) }
    assert_raises(ArgumentError) { sheet.row_height(2..1, 20) }
    assert_raises(TypeError) { sheet.row_height("1", 20) }
  end

  def test_row_height_works_with_native_fast_path
    skip "native extension not loaded" unless defined?(Rbxl::Native)

    Dir.mktmpdir do |dir|
      path = File.join(dir, "rh-native.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.row_height(0, 22)
      sheet << ["a"]
      sheet << ["b"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<row r="1" ht="22" customHeight="1">'
    end
  end

  def test_row_height_combined_with_column_width_uses_ruby_path
    Dir.mktmpdir do |dir|
      path = File.join(dir, "rh-combined.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.row_height(0, 22)
      sheet.column_width(0, 30)
      sheet << ["a"]
      sheet << ["b"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<col min="1" max="1" width="30" customWidth="1"/>'
      assert_includes sheet_xml, '<row r="1" ht="22" customHeight="1">'
    end
  end

  def test_external_hyperlink_emits_rels_file_and_rId_reference
    Dir.mktmpdir do |dir|
      path = File.join(dir, "ext-link.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet << ["intro"]
      sheet << [Rbxl::WriteOnlyCell.new("repo", hyperlink: "https://github.com/matzlika/rbxl")]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      assert_includes sheet_xml, '<hyperlinks>'
      assert_includes sheet_xml, '<hyperlink ref="A2" r:id="rId1"/>'

      rels = nil
      Zip::File.open(path) { |zf| rels = zf.read("xl/worksheets/_rels/sheet1.xml.rels") }
      assert_includes rels, 'Id="rId1"'
      assert_includes rels, 'Target="https://github.com/matzlika/rbxl"'
      assert_includes rels, 'TargetMode="External"'
      assert_includes rels, "schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    end
  end

  def test_internal_hyperlink_uses_location_attribute_without_rels
    Dir.mktmpdir do |dir|
      path = File.join(dir, "int-link.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet << [Rbxl::WriteOnlyCell.new("see Detail", hyperlink: "#'Detail'!A1")]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, %(<hyperlink ref="A1" location="&apos;Detail&apos;!A1"/>).gsub("&apos;", "'") # accept either form below
      refute_includes sheet_xml, 'r:id='

      Zip::File.open(path) do |zf|
        refute(zf.entries.any? { |e| e.name == "xl/worksheets/_rels/sheet1.xml.rels" },
               "internal-only sheet should not have a rels file")
      end
    end
  end

  def test_internal_hyperlink_strips_leading_hash_from_location
    Dir.mktmpdir do |dir|
      path = File.join(dir, "int-link2.xlsx")
      book = Rbxl.new
      book.add_sheet("S") << [Rbxl::WriteOnlyCell.new("x", hyperlink: "#Sheet2!A1")]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<hyperlink ref="A1" location="Sheet2!A1"/>'
    end
  end

  def test_mixed_hyperlinks_emit_both_external_and_internal_in_one_block
    Dir.mktmpdir do |dir|
      path = File.join(dir, "mixed.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet << [
        Rbxl::WriteOnlyCell.new("ext", hyperlink: "https://example.com"),
        Rbxl::WriteOnlyCell.new("int", hyperlink: "#Other!A1"),
      ]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_includes sheet_xml, '<hyperlink ref="A1" r:id="rId1"/>'
      assert_includes sheet_xml, '<hyperlink ref="B1" location="Other!A1"/>'
    end
  end

  def test_hyperlinks_block_appears_after_merge_cells
    Dir.mktmpdir do |dir|
      path = File.join(dir, "order.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.merge("A1:B1")
      sheet << [Rbxl::WriteOnlyCell.new("link", hyperlink: "https://example.com")]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      assert_match %r{</sheetData>.*<mergeCells .*?</mergeCells>.*<hyperlinks>.*</hyperlinks>}m, sheet_xml
    end
  end

  def test_hyperlink_rId_namespace_is_per_sheet
    Dir.mktmpdir do |dir|
      path = File.join(dir, "two-sheets.xlsx")
      book = Rbxl.new
      a = book.add_sheet("A")
      a << [Rbxl::WriteOnlyCell.new("first", hyperlink: "https://a.example.com")]
      b = book.add_sheet("B")
      b << [Rbxl::WriteOnlyCell.new("second", hyperlink: "https://b.example.com")]
      book.save(path)

      Zip::File.open(path) do |zf|
        a_xml = zf.read("xl/worksheets/sheet1.xml")
        b_xml = zf.read("xl/worksheets/sheet2.xml")
        a_rels = zf.read("xl/worksheets/_rels/sheet1.xml.rels")
        b_rels = zf.read("xl/worksheets/_rels/sheet2.xml.rels")

        assert_includes a_xml, 'r:id="rId1"'
        assert_includes b_xml, 'r:id="rId1"'
        assert_includes a_rels, "a.example.com"
        assert_includes b_rels, "b.example.com"
      end
    end
  end

  def test_no_hyperlinks_means_no_rels_file_and_no_xmlns_r
    Dir.mktmpdir do |dir|
      path = File.join(dir, "plain.xlsx")
      book = Rbxl.new
      sheet = book.add_sheet("S")
      sheet.column_width(0, 30)
      sheet << ["a"]
      book.save(path)

      sheet_xml = read_sheet_xml(path)
      refute_includes sheet_xml, "xmlns:r"
      refute_includes sheet_xml, "<hyperlinks>"

      Zip::File.open(path) do |zf|
        refute(zf.entries.any? { |e| e.name.start_with?("xl/worksheets/_rels/") })
      end
    end
  end

  def test_external_hyperlink_target_is_xml_escaped_in_rels
    Dir.mktmpdir do |dir|
      path = File.join(dir, "escape.xlsx")
      book = Rbxl.new
      book.add_sheet("S") << [Rbxl::WriteOnlyCell.new("link", hyperlink: "https://example.com/?a=1&b=2")]
      book.save(path)

      rels = nil
      Zip::File.open(path) { |zf| rels = zf.read("xl/worksheets/_rels/sheet1.xml.rels") }
      assert_includes rels, "a=1&amp;b=2"
    end
  end

  private

  def fixture_path(name)
    File.join(__dir__, "fixtures", name)
  end

  def read_sheet_xml(path, sheet_index: 1)
    Zip::File.open(path) { |zf| zf.read("xl/worksheets/sheet#{sheet_index}.xml") }
  end

  def write_minimal_workbook(path, sheet_body, styles: nil, workbook_pr: nil)
    Zip::File.open(path, create: true) do |zf|
      zf.get_output_stream("[Content_Types].xml") do |s|
        s.write <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            #{styles ? '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' : ''}
          </Types>
        XML
      end
      zf.get_output_stream("_rels/.rels") do |s|
        s.write <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          </Relationships>
        XML
      end
      zf.get_output_stream("xl/workbook.xml") do |s|
        s.write <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            #{workbook_pr}
            <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
          </workbook>
        XML
      end
      zf.get_output_stream("xl/_rels/workbook.xml.rels") do |s|
        s.write <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          </Relationships>
        XML
      end
      zf.get_output_stream("xl/worksheets/sheet1.xml") do |s|
        s.write <<~XML
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">#{sheet_body}</worksheet>
        XML
      end
      if styles
        zf.get_output_stream("xl/styles.xml") do |s|
          s.write <<~XML
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">#{styles}</styleSheet>
          XML
        end
      end
    end
  end
end
