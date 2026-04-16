module Rbxl
  class WriteOnlyWorksheet
    attr_reader :name

    def initialize(name:)
      @name = name
      @rows = []
      @column_name_cache = []
    end

    def <<(values)
      append(values)
    end

    def append(values)
      unless values.is_a?(Array) || values.is_a?(Enumerator)
        raise TypeError, "row must be an Array or Enumerator, got #{values.class}"
      end

      @rows << Array(values)
      self
    end

    def to_xml
      if defined?(Rbxl::Native)
        return Rbxl::Native.generate_sheet(@rows)
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
      buf << '  <sheetData>'

      @rows.each_with_index do |row_values, row_index|
        row_num_str = (row_index + 1).to_s
        buf << '<row r="'
        buf << row_num_str
        buf << '">'
        row_values.each_with_index do |value, col_index|
          serialize_cell_to(buf, column_name(col_index + 1), row_num_str, value)
        end
        buf << '</row>'
      end

      buf << "</sheetData>\n</worksheet>"
      buf
    end

    private

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
