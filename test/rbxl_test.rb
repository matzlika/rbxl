require "minitest/autorun"
require "tmpdir"
require_relative "../lib/rbxl"

class RbxlTest < Minitest::Test
  def test_open_requires_read_only_mode
    assert_raises(ArgumentError) { Rbxl.open("dummy.xlsx") }
  end

  def test_new_requires_write_only_mode
    assert_raises(ArgumentError) { Rbxl.new }
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

      assert_raises(Rbxl::ClosedWorkbookError) { book.add_sheet("Another") }
      assert_raises(Rbxl::ClosedWorkbookError) { book.save(path) }
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

  def test_reader_supports_padded_cells_for_missing_coordinates
    Dir.mktmpdir do |dir|
      path = File.join(dir, "sparse_manual.xlsx")
      build_sparse_xlsx(path)

      loaded = Rbxl.open(path, read_only: true)
      row = loaded.sheet("Sparse").each_row(pad_cells: true).first

      assert_equal ["left", nil, "right"], row.values
      assert_instance_of Rbxl::EmptyCell, row[1]
      assert_equal "B1", row[1].coordinate
      loaded.close
    end
  end

  private

  def build_sparse_xlsx(path)
    Zip::OutputStream.open(path) do |zip|
      write_entry(zip, "[Content_Types].xml", <<~XML.chomp)
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
      XML

      write_entry(zip, "_rels/.rels", <<~XML.chomp)
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      XML

      write_entry(zip, "xl/workbook.xml", <<~XML.chomp)
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="Sparse" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
      XML

      write_entry(zip, "xl/_rels/workbook.xml.rels", <<~XML.chomp)
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
      XML

      write_entry(zip, "xl/styles.xml", <<~XML.chomp)
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
  end

  def write_entry(zip, name, content)
    zip.put_next_entry(name)
    zip.write(content)
  end
end
