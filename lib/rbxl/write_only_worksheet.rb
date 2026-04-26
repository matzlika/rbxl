module Rbxl
  # Worksheet builder used by {Rbxl::WriteOnlyWorkbook}.
  #
  # Rows are appended in order and later serialized as SpreadsheetML by
  # {#to_xml} when the workbook is saved. The builder never rewrites a
  # previously appended row, so the worksheet's in-memory footprint scales
  # linearly with the number of appended rows.
  #
  # == Row values
  #
  # Each element of an appended row may be one of:
  #
  # * +nil+ — serialized as an empty cell
  # * +true+ / +false+ — serialized as a boolean cell
  # * +Integer+ / +Numeric+ — serialized as a numeric cell
  # * +Date+ / +DateTime+ / +Time+ — serialized as an ISO-8601 inline string
  # * any other object — serialized as +value.to_s+ in an inline string
  # * {Rbxl::WriteOnlyCell} — same as the above, but with an optional style id
  class WriteOnlyWorksheet
    # @return [String] visible sheet name
    attr_reader :name

    # @param name [String] visible sheet name
    def initialize(name:)
      @name = name
      @rows = []
      @column_name_cache = []
      @column_widths = {}
      @row_heights = {}
      @freeze = nil
      @merges = []
    end

    # Sets the display width of one or more columns.
    #
    #   sheet.column_width(0, 45.7)        # column A
    #   sheet.column_width(1..3, 12.0)     # columns B-D
    #
    # The argument is a 0-based column index or an ascending Range of
    # indexes. Calling +column_width+ on the same column twice keeps the
    # last value. Width values must be positive.
    #
    # @param index_or_range [Integer, Range<Integer>] 0-based column index or Range
    # @param width [Numeric] column width in Excel units
    # @return [Rbxl::WriteOnlyWorksheet] +self+ for chaining
    # @raise [ArgumentError] if the index is negative or the width is not positive
    # @raise [TypeError] if the first argument is neither Integer nor Range
    def column_width(index_or_range, width)
      width = Float(width)
      raise ArgumentError, "column width must be positive (got #{width})" unless width.positive?

      case index_or_range
      when Integer
        raise ArgumentError, "column index must be non-negative (got #{index_or_range})" if index_or_range.negative?

        @column_widths[index_or_range] = width
      when Range
        first = index_or_range.first
        last = index_or_range.last
        raise ArgumentError, "column range must not be negative (got #{index_or_range})" if first.negative?
        raise ArgumentError, "column range must be ascending (got #{index_or_range})" if first > last

        index_or_range.each { |i| @column_widths[i] = width }
      else
        raise TypeError, "expected Integer or Range, got #{index_or_range.class}"
      end

      self
    end

    # Sets the height of one or more rows.
    #
    #   sheet.row_height(0, 22)          # row 1 (header)
    #   sheet.row_height(0..2, 18)       # rows 1-3
    #
    # The argument is a 0-based row index or an ascending Range. Indexes
    # may target rows that have not yet been appended; the height is
    # applied if and when a matching row is emitted at save time.
    #
    # @param index_or_range [Integer, Range<Integer>] 0-based row index or Range
    # @param height [Numeric] row height in points
    # @return [Rbxl::WriteOnlyWorksheet] +self+ for chaining
    # @raise [ArgumentError] if the index is negative or the height is not positive
    # @raise [TypeError] if the first argument is neither Integer nor Range
    def row_height(index_or_range, height)
      height = Float(height)
      raise ArgumentError, "row height must be positive (got #{height})" unless height.positive?

      case index_or_range
      when Integer
        raise ArgumentError, "row index must be non-negative (got #{index_or_range})" if index_or_range.negative?

        @row_heights[index_or_range] = height
      when Range
        first = index_or_range.first
        last = index_or_range.last
        raise ArgumentError, "row range must not be negative (got #{index_or_range})" if first.negative?
        raise ArgumentError, "row range must be ascending (got #{index_or_range})" if first > last

        index_or_range.each { |i| @row_heights[i] = height }
      else
        raise TypeError, "expected Integer or Range, got #{index_or_range.class}"
      end

      self
    end

    # Freezes the top +rows+ rows and/or left +cols+ columns so they remain
    # visible while the rest of the sheet scrolls.
    #
    #   sheet.freeze!(rows: 1)              # header row only
    #   sheet.freeze!(rows: 1, cols: 2)     # header row + leftmost two columns
    #   sheet.freeze!(rows: 0, cols: 0)     # clears any prior freeze
    #
    # @param rows [Integer] number of rows to freeze from the top (>= 0)
    # @param cols [Integer] number of columns to freeze from the left (>= 0)
    # @return [Rbxl::WriteOnlyWorksheet] +self+ for chaining
    # @raise [ArgumentError] if either count is negative
    def freeze!(rows: 0, cols: 0)
      raise ArgumentError, "rows must be non-negative (got #{rows})" if rows.negative?
      raise ArgumentError, "cols must be non-negative (got #{cols})" if cols.negative?

      @freeze = (rows.zero? && cols.zero?) ? nil : { rows: rows, cols: cols }
      self
    end

    # Merges a rectangular range of cells. Excel keeps the value of the
    # top-left ("anchor") cell and visually spans it across the range.
    #
    #   sheet.merge("A1:C1")
    #   sheet.merge("A1", "C1")
    #
    # The range string is stored as written; rbxl does not validate that
    # the range is rectangular, ascending, or non-overlapping with other
    # merges — Excel will resolve the file as long as the syntax is valid.
    #
    # @overload merge(range_string)
    #   @param range_string [String] OOXML range like <tt>"A1:C1"</tt>
    # @overload merge(top_left, bottom_right)
    #   @param top_left [String] top-left cell coordinate
    #   @param bottom_right [String] bottom-right cell coordinate
    # @return [Rbxl::WriteOnlyWorksheet] +self+ for chaining
    # @raise [ArgumentError] if the arguments do not match either form
    def merge(*args)
      range =
        case args
        in [String => single]
          single
        in [String => from, String => to]
          "#{from}:#{to}"
        else
          raise ArgumentError, %(merge expects "A1:B2" or "A1", "B2")
        end

      @merges << range
      self
    end

    # Appends a row of values. Equivalent to {#append} so that the shell
    # operator reads naturally at the call site:
    #
    #   sheet << ["id", "name"]
    #
    # @param values [Array, Enumerator] row values
    # @return [Rbxl::WriteOnlyWorksheet] +self+ for chaining
    # @raise [TypeError] if +values+ is not Array-like
    def <<(values)
      append(values)
    end

    # Appends a row of values.
    #
    # @param values [Array, Enumerator] row values; each element is serialized
    #   according to the rules documented on the class
    # @return [Rbxl::WriteOnlyWorksheet] +self+ for chaining
    # @raise [TypeError] if +values+ is neither an Array nor an Enumerator
    def append(values)
      unless values.is_a?(Array) || values.is_a?(Enumerator)
        raise TypeError, "row must be an Array or Enumerator, got #{values.class}"
      end

      @rows << Array(values)
      self
    end

    # Serializes the worksheet to SpreadsheetML.
    #
    # When <tt>require "rbxl/native"</tt> has been loaded the native
    # extension handles serialization for a significant speedup; otherwise
    # a pure-Ruby implementation is used. Both paths produce equivalent
    # output.
    #
    # @return [String] worksheet XML
    def to_xml
      if column_meta_empty? && defined?(Rbxl::Native)
        return Rbxl::Native.generate_sheet(@rows, formatted_row_heights)
      end

      dimension_ref = @rows.empty? ? "A1" : "A1:#{column_name(max_columns)}#{@rows.length}"
      buf = +""
      buf << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      buf << "\n"
      buf << '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      buf << "\n"
      buf << '  <dimension ref="'
      buf << dimension_ref
      buf << '"/>'
      buf << "\n"

      append_sheet_views_to(buf) if @freeze
      append_cols_to(buf) unless @column_widths.empty?

      buf << '  <sheetData>'

      @rows.each_with_index do |row_values, row_index|
        row_num_str = (row_index + 1).to_s
        buf << '<row r="'
        buf << row_num_str
        buf << '"'
        if (height = @row_heights[row_index])
          buf << %( ht="#{format_number(height)}" customHeight="1")
        end
        buf << '>'
        row_values.each_with_index do |value, col_index|
          serialize_cell_to(buf, column_name(col_index + 1), row_num_str, value)
        end
        buf << '</row>'
      end

      buf << '</sheetData>'
      append_merge_cells_to(buf) unless @merges.empty?
      buf << "\n</worksheet>"
      buf
    end

    private

    def column_meta_empty?
      @column_widths.empty? && @freeze.nil? && @merges.empty?
    end

    def formatted_row_heights
      return nil if @row_heights.empty?

      @row_heights.each_with_object({}) { |(idx, h), out| out[idx] = format_number(h) }
    end

    def format_number(n)
      n == n.to_i ? n.to_i.to_s : n.to_s
    end

    def append_sheet_views_to(buf)
      rows = @freeze[:rows]
      cols = @freeze[:cols]
      top_left = "#{column_name(cols + 1)}#{rows + 1}"
      active_pane =
        if cols.positive? && rows.positive?
          "bottomRight"
        elsif cols.positive?
          "topRight"
        else
          "bottomLeft"
        end

      buf << '  <sheetViews><sheetView workbookViewId="0"><pane'
      buf << %( xSplit="#{cols}") if cols.positive?
      buf << %( ySplit="#{rows}") if rows.positive?
      buf << %( topLeftCell="#{top_left}")
      buf << %( activePane="#{active_pane}")
      buf << ' state="frozen"/></sheetView></sheetViews>'
      buf << "\n"
    end

    def append_cols_to(buf)
      sorted = @column_widths.sort
      buf << '  <cols>'
      i = 0
      while i < sorted.length
        start_idx, width = sorted[i]
        end_idx = start_idx
        j = i + 1
        while j < sorted.length && sorted[j][0] == end_idx + 1 && sorted[j][1] == width
          end_idx = sorted[j][0]
          j += 1
        end
        buf << %(<col min="#{start_idx + 1}" max="#{end_idx + 1}" width="#{format_number(width)}" customWidth="1"/>)
        i = j
      end
      buf << '</cols>'
      buf << "\n"
    end

    def append_merge_cells_to(buf)
      buf << %(<mergeCells count="#{@merges.length}">)
      @merges.each { |range| buf << %(<mergeCell ref="#{range}"/>) }
      buf << '</mergeCells>'
    end

    def serialize_cell_to(buf, col_name, row_num_str, value)
      if value.is_a?(WriteOnlyCell)
        serialize_write_only_cell_to(buf, col_name, row_num_str, value)
        return
      end

      buf << '<c r="'
      buf << col_name
      buf << row_num_str
      case value
      when nil
        buf << '"/>'
      when true
        buf << '" t="b"><v>1</v></c>'
      when false
        buf << '" t="b"><v>0</v></c>'
      when Integer
        buf << '"><v>'
        buf << value.to_s
        buf << '</v></c>'
      when Numeric
        buf << '"><v>'
        buf << value.to_s
        buf << '</v></c>'
      when Date, DateTime, Time
        buf << '" t="inlineStr"><is><t>'
        escape_to(buf, value.iso8601)
        buf << '</t></is></c>'
      else
        buf << '" t="inlineStr"><is><t>'
        escape_to(buf, value.to_s)
        buf << '</t></is></c>'
      end
    end

    def escape_to(buf, str)
      i = 0
      len = str.bytesize
      start = 0

      while i < len
        byte = str.getbyte(i)
        case byte
        when 38 # &
          buf << str.byteslice(start, i - start) if i > start
          buf << '&amp;'
          start = i + 1
        when 60 # <
          buf << str.byteslice(start, i - start) if i > start
          buf << '&lt;'
          start = i + 1
        when 62 # >
          buf << str.byteslice(start, i - start) if i > start
          buf << '&gt;'
          start = i + 1
        when 34 # "
          buf << str.byteslice(start, i - start) if i > start
          buf << '&quot;'
          start = i + 1
        end
        i += 1
      end

      if start == 0
        buf << str
      elsif start < len
        buf << str.byteslice(start, len - start)
      end
    end

    def column_name(index)
      @column_name_cache[index] ||= begin
        name = +""
        current = index
        while current.positive?
          current -= 1
          name.prepend((65 + (current % 26)).chr)
          current /= 26
        end
        name.freeze
      end
    end

    def max_columns
      max = 0
      @rows.each { |r| max = r.length if r.length > max }
      max > 0 ? max : 1
    end

    def serialize_write_only_cell_to(buf, col_name, row_num_str, cell)
      buf << '<c r="'
      buf << col_name
      buf << row_num_str
      buf << '"'
      if cell.style_id
        buf << ' s="'
        buf << cell.style_id.to_s
        buf << '"'
      end

      case cell.value
      when nil
        buf << '/>'
      when true
        buf << ' t="b"><v>1</v></c>'
      when false
        buf << ' t="b"><v>0</v></c>'
      when Numeric
        buf << '><v>'
        buf << cell.value.to_s
        buf << '</v></c>'
      when Date, DateTime, Time
        buf << ' t="inlineStr"><is><t>'
        escape_to(buf, cell.value.iso8601)
        buf << '</t></is></c>'
      else
        buf << ' t="inlineStr"><is><t>'
        escape_to(buf, cell.value.to_s)
        buf << '</t></is></c>'
      end
    end
  end
end
