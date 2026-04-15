module Rbxl
  class WriteOnlyWorksheet
    attr_reader :name

    def initialize(name:)
      @name = name
      @rows = []
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
      dimension_ref = @rows.empty? ? "A1" : "A1:#{column_name(max_columns)}#{@rows.length}"
      row_nodes = @rows.each_with_index.map do |row_values, row_index|
        cells = row_values.each_with_index.map do |value, col_index|
          reference = "#{column_name(col_index + 1)}#{row_index + 1}"
          serialize_cell(reference, value)
        end.join

        %(<row r="#{row_index + 1}">#{cells}</row>)
      end.join

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="#{dimension_ref}"/>
          <sheetData>#{row_nodes}</sheetData>
        </worksheet>
      XML
    end

    private

    def serialize_cell(reference, value)
      if value.is_a?(WriteOnlyCell)
        return serialize_write_only_cell(reference, value)
      end

      case value
      when nil
        %(<c r="#{reference}"/>)
      when true
        %(<c r="#{reference}" t="b"><v>1</v></c>)
      when false
        %(<c r="#{reference}" t="b"><v>0</v></c>)
      when Numeric
        %(<c r="#{reference}"><v>#{value}</v></c>)
      when Date, DateTime, Time
        %(<c r="#{reference}" t="inlineStr"><is><t>#{escape(value.iso8601)}</t></is></c>)
      else
        %(<c r="#{reference}" t="inlineStr"><is><t>#{escape(value)}</t></is></c>)
      end
    end

    def escape(value)
      CGI.escapeHTML(value.to_s)
    end

    def column_name(index)
      name = +""
      current = index

      while current.positive?
        current -= 1
        name.prepend((65 + (current % 26)).chr)
        current /= 26
      end

      name
    end

    def max_columns
      @rows.map(&:length).max || 1
    end

    def serialize_write_only_cell(reference, cell)
      style_attr = cell.style_id ? %( s="#{cell.style_id}") : ""
      serialized = serialize_scalar_value(cell.value)

      %(<c r="#{reference}"#{style_attr}#{serialized}>)
    end

    def serialize_scalar_value(value)
      case value
      when nil
        ""
      when true
        %( t="b"><v>1</v></c>)
      when false
        %( t="b"><v>0</v></c>)
      when Numeric
        %(<v>#{value}</v></c>)
      when Date, DateTime, Time
        %( t="inlineStr"><is><t>#{escape(value.iso8601)}</t></is></c>)
      else
        %( t="inlineStr"><is><t>#{escape(value)}</t></is></c>)
      end
    end
  end
end
