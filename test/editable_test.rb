require "minitest/autorun"
require "fileutils"
require "tmpdir"
require "zip"
require_relative "../lib/rbxl"
require_relative "editable_fixture_builder"

class EditableTest < Minitest::Test
  FIXTURE_DIR = File.expand_path("fixtures", __dir__)
  FIXTURE_PATH = File.join(FIXTURE_DIR, "editable.xlsx")
  FileUtils.mkdir_p(FIXTURE_DIR)
  Rbxl::Test::EditableFixtureBuilder.build(FIXTURE_PATH)

  # ---------- mode wiring ----------

  def test_open_with_edit_returns_editable_workbook
    book = Rbxl.open(FIXTURE_PATH, edit: true)
    assert_kind_of Rbxl::EditableWorkbook, book
    book.close
  end

  def test_open_with_edit_block_form_auto_closes
    captured = nil
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      captured = book
      refute book.closed?
    end
    assert captured.closed?
  end

  def test_open_with_edit_block_form_closes_on_exception
    captured = nil
    assert_raises(RuntimeError) do
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        captured = book
        raise "boom"
      end
    end
    assert captured.closed?
  end

  def test_edit_rejects_streaming_option
    assert_raises(ArgumentError) { Rbxl.open(FIXTURE_PATH, edit: true, streaming: true) }
  end

  def test_edit_rejects_date_conversion_option
    assert_raises(ArgumentError) { Rbxl.open(FIXTURE_PATH, edit: true, date_conversion: true) }
  end

  def test_open_default_still_returns_read_only
    book = Rbxl.open(FIXTURE_PATH)
    assert_kind_of Rbxl::ReadOnlyWorkbook, book
    book.close
  end

  # ---------- input validation ----------

  def test_open_rejects_non_zip_with_helpful_message
    Dir.mktmpdir do |dir|
      path = File.join(dir, "garbage.xlsx")
      File.binwrite(path, "definitely not a zip")

      err = assert_raises(Rbxl::UnsupportedFormatError) { Rbxl.open(path, edit: true) }
      assert_includes err.message, path
      assert_includes err.message, "ZIP signature"
    end
  end

  def test_open_rejects_legacy_xls_with_conversion_hint
    Dir.mktmpdir do |dir|
      path = File.join(dir, "legacy.xls")
      File.binwrite(path, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1".b + ("\x00".b * 512))

      err = assert_raises(Rbxl::UnsupportedFormatError) { Rbxl.open(path, edit: true) }
      assert_includes err.message, path
      assert_includes err.message, "legacy .xls"
      assert_includes err.message, "convert-to xlsx"
    end
  end

  # ---------- structural reads ----------

  def test_sheet_names_match_workbook_order
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_equal %w[Sheet1 Sheet2], book.sheet_names
    end
  end

  def test_sheet_lookup_by_name_and_index
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_equal "Sheet1", book.sheet("Sheet1").name
      assert_equal "Sheet1", book.sheet(0).name
      assert_equal "Sheet2", book.sheet(-1).name
    end
  end

  def test_sheet_returns_same_instance_across_calls
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_same book.sheet("Sheet1"), book.sheet("Sheet1")
      assert_same book.sheet("Sheet1"), book.sheet(0)
    end
  end

  def test_sheets_iterates_in_workbook_order
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_equal %w[Sheet1 Sheet2], book.sheets.map(&:name)
    end
  end

  def test_sheet_lookup_raises_for_missing_name
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_raises(Rbxl::SheetNotFoundError) { book.sheet("Bogus") }
    end
  end

  def test_sheet_lookup_raises_for_out_of_range_index
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_raises(Rbxl::SheetNotFoundError) { book.sheet(99) }
    end
  end

  # ---------- cell reads ----------

  def test_shared_string_cell_reads_through_sst
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      sheet = book.sheet("Sheet1")
      assert_equal "Header A", sheet["A1"].value
      assert_equal "Header B", sheet["B1"].value
    end
  end

  def test_inline_string_cell_reads_text
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_equal "alpha", book.sheet("Sheet1")["B2"].value
    end
  end

  def test_numeric_cell_reads_as_integer_when_whole
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      sheet = book.sheet("Sheet1")
      assert_equal 100, sheet["C1"].value
      assert_equal 1, sheet["A2"].value
      assert_equal 99, sheet["A5"].value
    end
  end

  def test_boolean_cell_reads_as_true_or_false
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_equal true, book.sheet("Sheet1")["C2"].value
    end
  end

  def test_missing_cell_reads_as_nil
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_nil book.sheet("Sheet1")["Z99"].value
      assert_nil book.sheet("Sheet1")["D2"].value
    end
  end

  def test_invalid_coordinate_raises
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_raises(ArgumentError) { book.sheet("Sheet1")["123"] }
      assert_raises(ArgumentError) { book.sheet("Sheet1")[""] }
    end
  end

  # ---------- save: byte-for-byte round-trip ----------

  def test_no_op_save_preserves_every_entry_byte_for_byte
    Dir.mktmpdir do |dir|
      out = File.join(dir, "copy.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) { |book| book.save(out) }

      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      assert_equal original.keys.sort, saved.keys.sort
      original.each do |name, bytes|
        assert_equal bytes, saved.fetch(name), "entry #{name} differs after no-op save"
      end
    end
  end

  def test_no_op_save_after_reading_does_not_dirty_sheets
    Dir.mktmpdir do |dir|
      out = File.join(dir, "copy.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        # Read a cell to force the worksheet DOM to load. A pure read must
        # not set the dirty flag — otherwise we'd re-serialize unchanged
        # sheets and lose byte-for-byte round trip.
        _ = book.sheet("Sheet1")["A1"].value
        refute book.sheet("Sheet1").dirty?
        book.save(out)
      end

      assert_equal File.binread(FIXTURE_PATH).bytesize > 0, true
      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      original.each do |name, bytes|
        assert_equal bytes, saved.fetch(name), "entry #{name} differs after read-only no-op save"
      end
    end
  end

  # ---------- save: surgical edits ----------

  def test_edit_marks_only_owning_sheet_dirty
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      book.sheet("Sheet1")["A1"].value = "新タイトル"
      assert book.sheet("Sheet1").dirty?
      refute book.sheet("Sheet2").dirty?
    end
  end

  def test_edit_clears_dirty_flag_after_save
    Dir.mktmpdir do |dir|
      out = File.join(dir, "edited.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = "X"
        assert book.sheet("Sheet1").dirty?
        book.save(out)
        refute book.sheet("Sheet1").dirty?, "dirty flag should clear after successful save"
      end
    end
  end

  def test_untouched_sheet_passes_through_byte_for_byte_after_edit
    Dir.mktmpdir do |dir|
      out = File.join(dir, "edited.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = "edited"
        book.save(out)
      end

      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      assert_equal original["xl/worksheets/sheet2.xml"], saved.fetch("xl/worksheets/sheet2.xml"),
                   "sheet2 should be unchanged after editing only sheet1"
      refute_equal original["xl/worksheets/sheet1.xml"], saved.fetch("xl/worksheets/sheet1.xml"),
                   "sheet1 should differ after edit"
    end
  end

  def test_shared_strings_xml_is_not_mutated_by_edits
    Dir.mktmpdir do |dir|
      out = File.join(dir, "edited.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        # Overwrite a previously-shared-string cell with a new string
        # value. The new value must end up inline; sharedStrings.xml must
        # round-trip byte-for-byte.
        book.sheet("Sheet1")["A1"].value = "完全に新しい文字列"
        book.save(out)
      end

      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      assert_equal original["xl/sharedStrings.xml"], saved.fetch("xl/sharedStrings.xml"),
                   "sharedStrings.xml must not be mutated by edits"
    end
  end

  def test_styles_xml_is_not_mutated_by_edits
    Dir.mktmpdir do |dir|
      out = File.join(dir, "edited.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = "x"
        book.save(out)
      end

      original = zip_entries(FIXTURE_PATH)
      saved = zip_entries(out)
      assert_equal original["xl/styles.xml"], saved.fetch("xl/styles.xml"),
                   "styles.xml must not be mutated by edits"
    end
  end

  # ---------- save: round-trip values ----------

  def test_string_value_round_trips_as_inline_string
    round_trip_value("A1", "hello world") do |reloaded|
      assert_equal "hello world", reloaded.sheet("Sheet1")["A1"].value
    end
  end

  def test_string_value_with_leading_whitespace_uses_xml_space_preserve
    Dir.mktmpdir do |dir|
      out = File.join(dir, "ws.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = "  trailing  "
        book.save(out)
      end

      sheet_xml = read_zip_entry(out, "xl/worksheets/sheet1.xml")
      assert_match(/xml:space="preserve"/, sheet_xml,
                   "leading/trailing whitespace requires xml:space=\"preserve\" to round-trip")

      Rbxl.open(out, edit: true) do |reloaded|
        assert_equal "  trailing  ", reloaded.sheet("Sheet1")["A1"].value
      end
    end
  end

  def test_string_value_escapes_xml_special_characters
    round_trip_value("A1", '<tag attr="x">&amp;</tag>') do |reloaded|
      assert_equal '<tag attr="x">&amp;</tag>', reloaded.sheet("Sheet1")["A1"].value
    end
  end

  def test_integer_value_round_trips_as_numeric
    round_trip_value("A1", 42) do |reloaded|
      assert_equal 42, reloaded.sheet("Sheet1")["A1"].value
    end
  end

  def test_float_value_round_trips_as_numeric
    round_trip_value("A1", 3.14) do |reloaded|
      assert_in_delta 3.14, reloaded.sheet("Sheet1")["A1"].value, 1e-9
    end
  end

  def test_boolean_true_round_trips
    round_trip_value("A1", true) do |reloaded|
      assert_equal true, reloaded.sheet("Sheet1")["A1"].value
    end
  end

  def test_boolean_false_round_trips
    round_trip_value("A1", false) do |reloaded|
      assert_equal false, reloaded.sheet("Sheet1")["A1"].value
    end
  end

  def test_nil_value_clears_cell_and_preserves_style_index
    Dir.mktmpdir do |dir|
      out = File.join(dir, "nilled.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = nil
        book.save(out)
      end

      sheet_xml = read_zip_entry(out, "xl/worksheets/sheet1.xml")
      # The cell should still exist and still carry s="1", but no value.
      assert_match(/<c r="A1"[^>]*s="1"[^>]*\/?>/, sheet_xml,
                   "style index should survive a nil clear")

      Rbxl.open(out, edit: true) do |reloaded|
        assert_nil reloaded.sheet("Sheet1")["A1"].value
      end
    end
  end

  # ---------- save: style preservation on overwrite ----------

  def test_existing_style_index_survives_value_overwrite
    Dir.mktmpdir do |dir|
      out = File.join(dir, "styled.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        # A1 was: <c r="A1" t="s" s="1"><v>0</v></c> (style 1 = bold)
        book.sheet("Sheet1")["A1"].value = "Replaced"
        book.save(out)
      end

      sheet_xml = read_zip_entry(out, "xl/worksheets/sheet1.xml")
      assert_match(/<c r="A1"[^>]*s="1"[^>]*>/, sheet_xml,
                   "s=\"1\" must survive a value overwrite")
      assert_includes sheet_xml, "Replaced"
    end
  end

  def test_overwrite_drops_t_attribute_for_numeric_assignment
    Dir.mktmpdir do |dir|
      out = File.join(dir, "numeric.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        # B2 was <c r="B2" t="inlineStr">. After numeric assignment it
        # should have no t attribute (numeric is the OOXML default).
        book.sheet("Sheet1")["B2"].value = 7
        book.save(out)
      end

      Rbxl.open(out, edit: true) do |reloaded|
        assert_equal 7, reloaded.sheet("Sheet1")["B2"].value
      end
    end
  end

  # ---------- save: insertion order ----------

  def test_new_cell_inserts_in_column_sorted_position_within_existing_row
    Dir.mktmpdir do |dir|
      out = File.join(dir, "insert.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        # Row 2 has A2, B2, C2 — write D2 (later) and AA2 (way later) to
        # test that the new cell lands in column-sorted position.
        sheet = book.sheet("Sheet1")
        sheet["D2"].value = "d"
        sheet["AA2"].value = "aa"
        book.save(out)
      end

      Rbxl.open(out, edit: true) do |reloaded|
        sheet = reloaded.sheet("Sheet1")
        assert_equal 1, sheet["A2"].value
        assert_equal "alpha", sheet["B2"].value
        assert_equal true, sheet["C2"].value
        assert_equal "d", sheet["D2"].value
        assert_equal "aa", sheet["AA2"].value
      end

      sheet_xml = read_zip_entry(out, "xl/worksheets/sheet1.xml")
      # Column order within row 2: A2 → B2 → C2 → D2 → AA2.
      indices = %w[A2 B2 C2 D2 AA2].map { |c| sheet_xml.index(%(r="#{c}")) }
      assert(indices.each_cons(2).all? { |a, b| a < b },
             "row 2 cells should remain in column-sorted order, got positions #{indices.inspect}")
    end
  end

  def test_new_row_inserts_in_row_sorted_position
    Dir.mktmpdir do |dir|
      out = File.join(dir, "row.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        # Existing rows: 1, 2, 5. Write into rows 3 and 7 so we get one
        # insert between existing rows and one append.
        sheet = book.sheet("Sheet1")
        sheet["A3"].value = "between"
        sheet["A7"].value = "after"
        book.save(out)
      end

      sheet_xml = read_zip_entry(out, "xl/worksheets/sheet1.xml")
      row_positions = [1, 2, 3, 5, 7].map { |n| sheet_xml.index(%(r="#{n}")) }
      assert(row_positions.each_cons(2).all? { |a, b| a < b },
             "rows must serialize in numeric order, got positions #{row_positions.inspect}")

      Rbxl.open(out, edit: true) do |reloaded|
        assert_equal "between", reloaded.sheet("Sheet1")["A3"].value
        assert_equal "after",   reloaded.sheet("Sheet1")["A7"].value
      end
    end
  end

  def test_writing_to_brand_new_row_creates_it_with_sibling_cells_intact
    Dir.mktmpdir do |dir|
      out = File.join(dir, "newrow.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet["B100"].value = "b"
        sheet["A100"].value = "a"
        book.save(out)
      end

      Rbxl.open(out, edit: true) do |reloaded|
        sheet = reloaded.sheet("Sheet1")
        assert_equal "a", sheet["A100"].value
        assert_equal "b", sheet["B100"].value
        # Existing row 1 cells must still be readable.
        assert_equal "Header A", sheet["A1"].value
      end
    end
  end

  # ---------- save: in-place ----------

  def test_save_in_place_overwrites_atomically
    Dir.mktmpdir do |dir|
      target = File.join(dir, "deck.xlsx")
      FileUtils.cp(FIXTURE_PATH, target)

      Rbxl.open(target, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = "上書き"
        book.save                         # no path → in place
      end

      Rbxl.open(target, edit: true) do |reloaded|
        assert_equal "上書き", reloaded.sheet("Sheet1")["A1"].value
      end

      refute Dir.children(dir).any? { |f| f.include?("rbxl-tmp") }, "temp files should not remain"
    end
  end

  def test_edited_file_round_trips_through_read_only_workbook
    Dir.mktmpdir do |dir|
      out = File.join(dir, "for_read_only.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        sheet = book.sheet("Sheet1")
        sheet["A1"].value = "edit-mode value"
        sheet["D2"].value = 42
        sheet["A100"].value = "far row"
        book.save(out)
      end

      Rbxl.open(out) do |reloaded|
        rows = reloaded.sheet("Sheet1").rows(values_only: true).to_a
        assert_equal "edit-mode value", rows[0][0]
        assert_equal 42, rows[1][3]
        # Last row carries the inserted A100; existing rows in between
        # remain readable through the streaming row reader.
        assert_includes rows.flatten, "far row"
      end
    end
  end

  def test_save_explicit_path_does_not_modify_original
    Dir.mktmpdir do |dir|
      target = File.join(dir, "deck.xlsx")
      FileUtils.cp(FIXTURE_PATH, target)
      original_bytes = File.binread(target)

      Rbxl.open(target, edit: true) do |book|
        book.sheet("Sheet1")["A1"].value = "新規"
        book.save(File.join(dir, "out.xlsx"))
      end

      assert_equal original_bytes, File.binread(target),
                   "save(other_path) must not touch the original load path"
    end
  end

  # ---------- type rejection ----------

  def test_date_value_raises_editable_cell_type_error
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      err = assert_raises(Rbxl::EditableCellTypeError) do
        book.sheet("Sheet1")["A1"].value = Date.new(2026, 4, 27)
      end
      assert_includes err.message, "Date"
    end
  end

  def test_arbitrary_object_raises_editable_cell_type_error
    Rbxl.open(FIXTURE_PATH, edit: true) do |book|
      assert_raises(Rbxl::EditableCellTypeError) do
        book.sheet("Sheet1")["A1"].value = Object.new
      end
    end
  end

  # ---------- lifecycle ----------

  def test_close_is_idempotent
    book = Rbxl.open(FIXTURE_PATH, edit: true)
    assert_equal true, book.close
    assert_equal false, book.close
    assert book.closed?
  end

  def test_sheet_lookup_after_close_raises
    book = Rbxl.open(FIXTURE_PATH, edit: true)
    book.close
    assert_raises(Rbxl::ClosedWorkbookError) { book.sheet("Sheet1") }
  end

  def test_save_after_close_raises
    book = Rbxl.open(FIXTURE_PATH, edit: true)
    book.close
    Dir.mktmpdir do |dir|
      assert_raises(Rbxl::ClosedWorkbookError) { book.save(File.join(dir, "x.xlsx")) }
    end
  end

  private

  def round_trip_value(coordinate, value)
    Dir.mktmpdir do |dir|
      out = File.join(dir, "rt.xlsx")
      Rbxl.open(FIXTURE_PATH, edit: true) do |book|
        book.sheet("Sheet1")[coordinate].value = value
        book.save(out)
      end

      Rbxl.open(out, edit: true) { |reloaded| yield reloaded }
    end
  end

  def zip_entries(path)
    Zip::File.open(path) do |zf|
      zf.each_with_object({}) do |entry, h|
        next if entry.directory?

        h[entry.name] = entry.get_input_stream.read
      end
    end
  end

  def read_zip_entry(path, name)
    Zip::File.open(path) { |zf| zf.find_entry(name).get_input_stream.read.force_encoding("UTF-8") }
  end
end
