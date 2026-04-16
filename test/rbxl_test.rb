require "minitest/autorun"
require "pathname"
require "tmpdir"
require "zip"
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

  private

  def fixture_path(name)
    File.join(__dir__, "fixtures", name)
  end
end
