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

  private

  def fixture_path(name)
    File.join(__dir__, "fixtures", name)
  end

  def write_minimal_workbook(path, sheet_body, styles: nil, workbook_pr: nil)
    Zip::File.open(path, Zip::File::CREATE) do |zf|
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
