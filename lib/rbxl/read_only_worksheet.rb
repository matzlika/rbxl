module Rbxl
  class ReadOnlyWorksheet
    ELEMENT_NODE = Nokogiri::XML::Reader::TYPE_ELEMENT
    TEXT_NODE = Nokogiri::XML::Reader::TYPE_TEXT
    CDATA_NODE = Nokogiri::XML::Reader::TYPE_CDATA
    END_ELEMENT_NODE = Nokogiri::XML::Reader::TYPE_END_ELEMENT

    attr_reader :name, :dimensions

    def initialize(zip:, entry_path:, shared_strings:, name:)
      @zip = zip
      @entry_path = entry_path
      @shared_strings = shared_strings
      @name = name
      @dimensions = extract_dimensions
    end

    def each_row(pad_cells: false, values_only: false)
      return enum_for(:each_row, pad_cells: pad_cells, values_only: values_only) unless block_given?

      current_row_index = nil
      last_row_index = 0
      current_cells = nil
      cell_ref = nil
      cell_type = nil
      current_col_index = 0
      collecting_value = false
      raw_value = nil
      value_buffer = +""

      with_sheet_reader do |reader|
        reader.each do |node|
          case node.node_type
          when ELEMENT_NODE
            case node.local_name
            when "row"
              current_row_index = attribute_int(node, "r") || (last_row_index + 1)
              current_col_index = 0
              current_cells = []
            when "c"
              cell_ref = node.attribute("r")
              if cell_ref
                col_index, = split_coordinate(cell_ref)
                current_col_index = col_index || current_col_index
              else
                current_col_index += 1
                cell_ref = "#{column_name(current_col_index)}#{current_row_index}"
              end
              cell_type = node.attribute("t")
              raw_value = nil
            when "v", "t"
              collecting_value = true
              value_buffer.clear
            end
          when TEXT_NODE, CDATA_NODE
            value_buffer << node.value if collecting_value
          when END_ELEMENT_NODE
            case node.local_name
            when "v"
              raw_value = value_buffer.dup
              collecting_value = false
            when "t"
              raw_value = raw_value ? raw_value << value_buffer : value_buffer.dup
              collecting_value = false
            when "c"
              current_cells << build_row_entry(cell_ref, coerce_value(raw_value, cell_type), values_only)
              cell_ref = nil
              cell_type = nil
              raw_value = nil
            when "row"
              current_cells = pad_row(current_cells, current_row_index, values_only: values_only) if pad_cells
              yield values_only ? extract_values(current_cells).freeze : Row.new(index: current_row_index, cells: current_cells)
              last_row_index = current_row_index
              current_row_index = nil
              current_cells = nil
            end
          end
        end
      end
    end

    def rows(values_only: false)
      each_row(values_only: values_only)
    end

    def max_column
      return nil unless dimensions

      dimensions[:max_col]
    end

    def max_row
      return nil unless dimensions

      dimensions[:max_row]
    end

    def reset_dimensions
      @dimensions = nil
    end

    def calculate_dimension(force: false)
      if dimensions
        return dimensions[:ref]
      end

      raise UnsizedWorksheetError, "worksheet is unsized, use force: true" unless force

      @dimensions = scan_dimensions
      dimensions ? dimensions[:ref] : "A1:A1"
    end

    private

    def with_sheet_reader
      io = @zip.get_entry(@entry_path).get_input_stream
      reader = Nokogiri::XML::Reader(io)
      yield reader
    ensure
      io&.close
    end

    def extract_dimensions
      with_sheet_reader do |reader|
        reader.each do |node|
          next unless node.node_type == ELEMENT_NODE && node.local_name == "dimension"

          return parse_range(node.attribute("ref"))
        end
      end

      nil
    end

    def scan_dimensions
      max_col = nil
      max_row = nil

      with_sheet_reader do |reader|
        reader.each do |node|
          next unless node.node_type == ELEMENT_NODE && node.local_name == "c"

          coordinate = node.attribute("r")
          col, row = split_coordinate(coordinate)
          next unless col && row

          max_col = col if max_col.nil? || col > max_col
          max_row = row if max_row.nil? || row > max_row
        end
      end

      return nil unless max_col && max_row

      { ref: "A1:#{column_name(max_col)}#{max_row}", max_col: max_col, max_row: max_row }
    end

    def parse_range(reference)
      return nil if reference.nil? || reference.empty?

      start_ref, finish_ref = reference.split(":", 2)
      finish_ref ||= start_ref
      _, _, max_col, max_row = *range_bounds(start_ref, finish_ref)
      { ref: reference, max_col: max_col, max_row: max_row }
    end

    def range_bounds(start_ref, finish_ref)
      start_col, start_row = split_coordinate(start_ref)
      finish_col, finish_row = split_coordinate(finish_ref)
      [start_col, start_row, finish_col, finish_row]
    end

    def split_coordinate(reference)
      match = reference.match(/\A([A-Z]+)(\d+)\z/)
      return [nil, nil] unless match

      [column_index(match[1]), match[2].to_i]
    end

    def column_index(label)
      label.each_char.reduce(0) { |sum, char| (sum * 26) + (char.ord - 64) }
    end

    def pad_row(cells, row_index, values_only:)
      return cells unless dimensions && dimensions[:max_col]

      by_column = cells.each_with_object({}) do |cell, acc|
        coordinate =
          if cell.respond_to?(:coordinate)
            cell.coordinate
          elsif values_only
            cell[0]
          end
        next unless coordinate

        acc[column_index(coordinate[/\A[A-Z]+/])] = cell
      end

      (1..dimensions[:max_col]).map do |col|
        by_column[col] || (values_only ? [nil, nil] : EmptyCell.new(coordinate: "#{column_name(col)}#{row_index}"))
      end
    end

    def coerce_value(raw_value, type)
      case type
      when "s"
        @shared_strings[raw_value.to_i]
      when "inlineStr", "str"
        raw_value
      when "b"
        raw_value == "1"
      else
        infer_scalar(raw_value)
      end
    end

    def infer_scalar(raw_value)
      return nil if raw_value.nil? || raw_value.empty?
      return raw_value.to_i if raw_value.match?(/\A-?\d+\z/)
      return raw_value.to_f if raw_value.match?(/\A-?(?:\d+\.\d+|\d+\.|\.\d+)\z/)

      raw_value
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

    def attribute_int(node, key)
      value = node.attribute(key)
      value&.to_i
    end

    def build_row_entry(coordinate, value, values_only)
      return [coordinate, value] if values_only

      ReadOnlyCell.new(coordinate, value)
    end

    def extract_values(cells)
      cells.map { |cell| cell.is_a?(Array) ? cell[1] : cell }
    end
  end
end
