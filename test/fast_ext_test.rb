require "minitest/autorun"
require "tmpdir"
require_relative "../lib/rbxl"
require_relative "../lib/rbxl/native"

class FastExtTest < Minitest::Test
  # -----------------------------------------------------------
  # Native reader tests
  # -----------------------------------------------------------

  def test_reader_values_only_matches_nokogiri
    with_test_file do |path|
      assert_equal read_values_nokogiri(path), read_values_fast(path)
    end
  end

  def test_reader_values_only_streaming_matches_default
    with_test_file do |path|
      assert_equal read_values_fast(path), read_values_fast_streaming(path)
    end
  end

  def test_reader_full_matches_nokogiri
    with_test_file do |path|
      expected = read_full_nokogiri(path)
      actual = read_full_fast(path)

      assert_equal expected.size, actual.size
      expected.zip(actual).each_with_index do |(e_row, a_row), i|
        assert_equal e_row.index, a_row.index, "row #{i} index"
        e_row.cells.zip(a_row.cells).each_with_index do |(e_cell, a_cell), j|
          assert_equal e_cell.coordinate, a_cell.coordinate, "row #{i} cell #{j} coordinate"
          assert_equal e_cell.value, a_cell.value, "row #{i} cell #{j} value"
        end
      end
    end
  end

  def test_reader_full_streaming_matches_default
    with_test_file do |path|
      expected = read_full_fast(path)
      actual = read_full_fast_streaming(path)

      assert_equal expected.size, actual.size
      expected.zip(actual).each_with_index do |(e_row, a_row), i|
        assert_equal e_row.index, a_row.index, "row #{i} index"
        e_row.cells.zip(a_row.cells).each_with_index do |(e_cell, a_cell), j|
          assert_equal e_cell.coordinate, a_cell.coordinate, "row #{i} cell #{j} coordinate"
          assert_equal e_cell.value, a_cell.value, "row #{i} cell #{j} value"
        end
      end
    end
  end

  def test_reader_unicode_values
    round_trip_values(["日本語", "émojis 🎉", "中文测试", "Ü∞ß"]) do |row|
      assert_equal "日本語", row[0]
      assert_equal "émojis 🎉", row[1]
      assert_equal "中文测试", row[2]
      assert_equal "Ü∞ß", row[3]
      row.each { |v| assert_equal Encoding::UTF_8, v.encoding }
    end
  end

  def test_reader_multibyte_with_special_chars
    round_trip_values(["<日本語>&\"引用\"", "foo&bar<baz>"]) do |row|
      assert_equal "<日本語>&\"引用\"", row[0]
      assert_equal "foo&bar<baz>", row[1]
    end
  end

  def test_reader_empty_and_nil
    round_trip_values(["", nil, "a"]) do |row|
      assert_equal "", row[0]
      assert_nil row[1]
      assert_equal "a", row[2]
    end
  end

  def test_reader_shared_strings_unicode
    Dir.mktmpdir do |dir|
      path = File.join(dir, "shared_unicode.xlsx")
      book = Rbxl.new(write_only: true)
      sheet = book.add_sheet("Sheet")
      sheet.append(["こんにちは", "世界"])
      sheet.append(["こんにちは", "世界"]) # reuse triggers shared strings
      book.save(path)

      values = read_values_fast(path)
      assert_equal ["こんにちは", "世界"], values[0]
      assert_equal ["こんにちは", "世界"], values[1]
      values.flatten.each { |v| assert_equal Encoding::UTF_8, v.encoding }
    end
  end

  # -----------------------------------------------------------
  # Native writer tests
  # -----------------------------------------------------------

  def test_writer_output_matches_ruby
    rows = build_test_rows
    ruby_xml = ruby_to_xml(rows)
    fast_xml = Rbxl::Native.generate_sheet(rows)

    assert_equal ruby_xml, fast_xml
  end

  def test_writer_unicode_round_trip
    Dir.mktmpdir do |dir|
      path = File.join(dir, "unicode.xlsx")
      book = Rbxl.new(write_only: true)
      book.add_sheet("Sheet").append(["日本語テスト", "émojis 🎉🚀", "中文", "Ü∞ß"])
      book.save(path)

      values = read_values_fast(path)
      assert_equal ["日本語テスト", "émojis 🎉🚀", "中文", "Ü∞ß"], values.first
      values.first.each { |v| assert_equal Encoding::UTF_8, v.encoding }
    end
  end

  def test_writer_escaping
    Dir.mktmpdir do |dir|
      path = File.join(dir, "escape.xlsx")
      book = Rbxl.new(write_only: true)
      book.add_sheet("Sheet").append(["&amp;", "<tag>", "a\"b", "日<本>&語"])
      book.save(path)

      values = read_values_fast(path)
      assert_equal ["&amp;", "<tag>", "a\"b", "日<本>&語"], values.first
    end
  end

  def test_writer_numeric_types
    Dir.mktmpdir do |dir|
      path = File.join(dir, "nums.xlsx")
      book = Rbxl.new(write_only: true)
      book.add_sheet("Sheet").append([42, -7, 3.14, 0, 0.0])
      book.save(path)

      values = read_values_fast(path)
      row = values.first
      assert_equal 42, row[0]
      assert_equal(-7, row[1])
      assert_in_delta 3.14, row[2], 1e-10
      assert_equal 0, row[3]
      assert_in_delta 0.0, row[4], 1e-10
    end
  end

  def test_writer_boolean_and_nil
    Dir.mktmpdir do |dir|
      path = File.join(dir, "bool.xlsx")
      book = Rbxl.new(write_only: true)
      book.add_sheet("Sheet").append([true, false, nil])
      book.save(path)

      values = read_values_fast(path)
      assert_equal [true, false, nil], values.first
    end
  end

  def test_writer_empty_string
    Dir.mktmpdir do |dir|
      path = File.join(dir, "empty.xlsx")
      book = Rbxl.new(write_only: true)
      book.add_sheet("Sheet").append(["", "a", ""])
      book.save(path)

      values = read_values_fast(path)
      assert_equal ["", "a", ""], values.first
    end
  end

  def test_writer_write_only_cell_with_style
    Dir.mktmpdir do |dir|
      path = File.join(dir, "styled.xlsx")
      book = Rbxl.new(write_only: true)
      book.add_sheet("Sheet").append([Rbxl::WriteOnlyCell.new("styled", style_id: 1)])
      book.save(path)

      wb = Rbxl.open(path, read_only: true)
      row = wb.sheet("Sheet").rows.first
      assert_equal "styled", row.cells[0].value
      wb.close
    end
  end

  def test_writer_output_encoding_is_utf8
    xml = Rbxl::Native.generate_sheet([["hello"]])
    assert_equal Encoding::UTF_8, xml.encoding
  end

  private

  def with_test_file
    Dir.mktmpdir do |dir|
      path = File.join(dir, "test.xlsx")
      book = Rbxl.new(write_only: true)
      sheet = book.add_sheet("Bench")
      sheet.append(Array.new(10) { |i| "col_#{i}" })
      100.times do |row|
        sheet.append(Array.new(10) { |col|
          case col % 4
          when 0 then row
          when 1 then "row-#{row}-col-#{col}"
          when 2 then (row + col).odd?
          else ((row * 100) + col) / 10.0
          end
        })
      end
      book.save(path)
      yield path
    end
  end

  def round_trip_values(input)
    Dir.mktmpdir do |dir|
      path = File.join(dir, "rt.xlsx")
      book = Rbxl.new(write_only: true)
      book.add_sheet("Sheet").append(input)
      book.save(path)

      values = read_values_fast(path)
      yield values.first
    end
  end

  def build_test_rows
    [
      ["hello", 42, 3.14, true, false, nil],
      ["日本語", -1, 0.0, false, true, nil],
    ]
  end

  def ruby_to_xml(rows)
    ws = Rbxl::WriteOnlyWorksheet.new(name: "test")
    # Temporarily remove Native to force Ruby path
    native = Rbxl.const_get(:Native)
    Rbxl.send(:remove_const, :Native)
    rows.each { |r| ws.append(r) }
    result = ws.to_xml
    Rbxl.const_set(:Native, native)
    result
  end

  def read_values_nokogiri(path)
    wb = Rbxl.open(path, read_only: true)
    ws = wb.sheet(wb.sheet_names.first)
    ws.instance_variable_set(:@disable_native, true)
    result = ws.rows(values_only: true).to_a
    wb.close
    result
  end

  def read_values_fast(path)
    wb = Rbxl.open(path, read_only: true)
    result = wb.sheet(wb.sheet_names.first).rows(values_only: true).to_a
    wb.close
    result
  end

  def read_full_nokogiri(path)
    wb = Rbxl.open(path, read_only: true)
    ws = wb.sheet(wb.sheet_names.first)
    ws.instance_variable_set(:@disable_native, true)
    result = ws.rows.to_a
    wb.close
    result
  end

  def read_full_fast(path)
    wb = Rbxl.open(path, read_only: true)
    result = wb.sheet(wb.sheet_names.first).rows.to_a
    wb.close
    result
  end

  def read_values_fast_streaming(path)
    wb = Rbxl.open(path, read_only: true, streaming: true)
    result = wb.sheet(wb.sheet_names.first).rows(values_only: true).to_a
    wb.close
    result
  end

  def read_full_fast_streaming(path)
    wb = Rbxl.open(path, read_only: true, streaming: true)
    result = wb.sheet(wb.sheet_names.first).rows.to_a
    wb.close
    result
  end
end
