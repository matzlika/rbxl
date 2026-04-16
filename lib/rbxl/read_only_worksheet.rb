module Rbxl
  # Streaming worksheet reader for a single sheet of a read-only workbook.
  #
  # Instances are produced by {Rbxl::ReadOnlyWorkbook#sheet} and must not be
  # constructed directly; their lifecycle is bound to the workbook's ZIP
  # handle. Rows can be consumed as {Rbxl::Row} objects or as plain value
  # arrays depending on the iteration options.
  #
  # == Iteration modes
  #
  #   # Default: yield Rbxl::Row with cell wrappers.
  #   sheet.each_row { |row| row.values }
  #
  #   # Fast path: yield plain Array<Object> of decoded values.
  #   sheet.each_row(values_only: true) { |values| ... }
  #
  #   # Pad missing cells in sparse rows up to max_column.
  #   sheet.each_row(pad_cells: true) { |row| ... }
  #
  #   # Replicate anchor values across merged ranges.
  #   sheet.each_row(expand_merged: true) { |row| ... }
  #
  # Iteration without a block returns an +Enumerator+.
  #
  # == Dimensions
  #
  # The worksheet dimension (the <tt>A1:C10</tt>-style range) is read from
  # the sheet's +<dimension>+ element when present. When absent or when the
  # caller wants to recompute it, {#calculate_dimension} with
  # <tt>force: true</tt> scans the sheet for actual cell coordinates.
  class ReadOnlyWorksheet
    # @private Nokogiri reader node-type shortcuts.
    ELEMENT_NODE = Nokogiri::XML::Reader::TYPE_ELEMENT
    # @private
    TEXT_NODE = Nokogiri::XML::Reader::TYPE_TEXT
    # @private
    CDATA_NODE = Nokogiri::XML::Reader::TYPE_CDATA
    # @private
    END_ELEMENT_NODE = Nokogiri::XML::Reader::TYPE_END_ELEMENT

    # @return [String] visible sheet name
    attr_reader :name

    # Parsed dimension metadata, +nil+ when the sheet has no +<dimension>+
    # element and no scan has been forced. When present the hash has keys
    # +:ref+, +:max_col+, and +:max_row+.
    #
    # @return [Hash{Symbol => Object}, nil]
    attr_reader :dimensions

    # @param zip [Zip::File] open archive shared with the workbook
    # @param entry_path [String] ZIP entry path for this sheet's XML
    # @param shared_strings [Array<String>] pre-decoded shared strings table
    # @param name [String] visible sheet name
    def initialize(zip:, entry_path:, shared_strings:, name:)
      @zip = zip
      @entry_path = entry_path
      @shared_strings = shared_strings
      @name = name
      @dimensions = extract_dimensions
      @merge_ranges_by_row = nil
      @merge_anchor_values = {}
    end

    # Iterates rows in worksheet order.
    #
    # With +values_only+ and neither +pad_cells+ nor +expand_merged+ set,
    # iteration takes a tighter path that yields frozen +Array<Object>+
    # rows and skips allocating cell wrappers.
    #
    # @param pad_cells [Boolean] pad sparse rows with {Rbxl::EmptyCell} (or
    #   <tt>[coordinate, nil]</tt> pairs in +values_only+ mode) up to the
    #   worksheet's +max_column+
    # @param values_only [Boolean] yield plain value arrays instead of
    #   {Rbxl::Row} instances
    # @param expand_merged [Boolean] propagate the anchor value of every
    #   merged range across the range's cells
    # @yieldparam row [Rbxl::Row, Array<Object>]
    # @return [Enumerator, void] enumerator when called without a block
    def each_row(pad_cells: false, values_only: false, expand_merged: false, &block)
      return enum_for(:each_row, pad_cells: pad_cells, values_only: values_only, expand_merged: expand_merged) unless block

      if values_only && !pad_cells && !expand_merged
        each_row_values_only(&block)
      else
        each_row_full(pad_cells: pad_cells, values_only: values_only, expand_merged: expand_merged, &block)
      end
    end

    # Enumerator-returning alias for {#each_row} that reads more naturally
    # when the call site chains further enumerable operations.
    #
    #   sheet.rows(values_only: true).take(10)
    #
    # @param values_only [Boolean] see {#each_row}
    # @param pad_cells [Boolean] see {#each_row}
    # @param expand_merged [Boolean] see {#each_row}
    # @return [Enumerator]
    def rows(values_only: false, pad_cells: false, expand_merged: false)
      each_row(values_only: values_only, pad_cells: pad_cells, expand_merged: expand_merged)
    end

    # @return [Integer, nil] rightmost column index (1-based) from the
    #   worksheet dimension, or +nil+ when dimensions are unknown
    def max_column
      return nil unless dimensions

      dimensions[:max_col]
    end

    # @return [Integer, nil] bottom row index (1-based) from the worksheet
    #   dimension, or +nil+ when dimensions are unknown
    def max_row
      return nil unless dimensions

      dimensions[:max_row]
    end

    # Clears cached dimension metadata so that the next call to
    # {#calculate_dimension} recomputes it.
    #
    # @return [nil]
    def reset_dimensions
      @dimensions = nil
    end

    # Returns the worksheet dimension reference (e.g. <tt>"A1:C10"</tt>).
    #
    # When the sheet lacks a +<dimension>+ element the default is to raise
    # {Rbxl::UnsizedWorksheetError}. Passing <tt>force: true</tt> scans the
    # sheet for the actual cell bounds instead; a sheet with no cells at
    # all falls back to <tt>"A1:A1"</tt>.
    #
    # @param force [Boolean] scan the sheet when no stored dimension exists
    # @return [String] Excel-style range reference
    # @raise [Rbxl::UnsizedWorksheetError] if the sheet is unsized and
    #   +force+ is +false+
    def calculate_dimension(force: false)
      if dimensions
        return dimensions[:ref]
      end

      raise UnsizedWorksheetError, "worksheet is unsized, use force: true" unless force

      @dimensions = scan_dimensions
      dimensions ? dimensions[:ref] : "A1:A1"
    end

    private

    def each_row_values_only(&block)
      if defined?(Rbxl::Native) && !@disable_native
        xml = @zip.get_entry(@entry_path).get_input_stream.read
        Rbxl::Native.parse_sheet(xml, @shared_strings, &block)
        return
      end

      cell_type = nil
      collecting_value = false
      in_v = false
      raw_value = nil
      value_buffer = +""
      current_values = nil
      row_depth = nil

      with_sheet_reader do |reader|
        reader.each do |node|
          case node.node_type
          when ELEMENT_NODE
            case node.local_name
            when "row"
              current_values = []
              row_depth = node.depth
            when "c"
              cell_type = node.attribute("t")
              raw_value = nil
            when "v"
              collecting_value = true
              in_v = true
              value_buffer.clear
            when "t"
              collecting_value = true
              value_buffer.clear
            end
          when TEXT_NODE, CDATA_NODE
            value_buffer << node.value if collecting_value
          when END_ELEMENT_NODE
            if collecting_value
              if in_v
                raw_value = value_buffer.dup
                collecting_value = false
                in_v = false
              else
                raw_value = raw_value ? raw_value << value_buffer : value_buffer.dup
                collecting_value = false
              end
            elsif node.depth == row_depth
              yield current_values.freeze
              current_values = nil
            elsif current_values && node.depth == row_depth + 1
              current_values << coerce_value(raw_value, cell_type)
              cell_type = nil
              raw_value = nil
            end
          end
        end
      end
    end

    def each_row_full(pad_cells:, values_only:, expand_merged:, &block)
      if defined?(Rbxl::Native) && !@disable_native && !pad_cells && !expand_merged && !values_only
        xml = @zip.get_entry(@entry_path).get_input_stream.read
        Rbxl::Native.parse_sheet_full(xml, @shared_strings, &block)
        return
      end

      current_row_index = nil
      last_row_index = 0
      current_cells = nil
      cell_ref = nil
      cell_type = nil
      current_col_index = 0
      collecting_value = false
      in_v = false
      raw_value = nil
      value_buffer = +""
      row_depth = nil

      with_sheet_reader do |reader|
        reader.each do |node|
          case node.node_type
          when ELEMENT_NODE
            case node.local_name
            when "row"
              current_row_index = attribute_int(node, "r") || (last_row_index + 1)
              current_col_index = 0
              current_cells = []
              row_depth = node.depth
            when "c"
              cell_ref = node.attribute("r")
              if cell_ref
                current_col_index = split_col_index(cell_ref)
              else
                current_col_index += 1
                cell_ref = "#{column_name(current_col_index)}#{current_row_index}"
              end
              cell_type = node.attribute("t")
              raw_value = nil
            when "v"
              collecting_value = true
              in_v = true
              value_buffer.clear
            when "t"
              collecting_value = true
              value_buffer.clear
            end
          when TEXT_NODE, CDATA_NODE
            value_buffer << node.value if collecting_value
          when END_ELEMENT_NODE
            if collecting_value
              if in_v
                raw_value = value_buffer.dup
                collecting_value = false
                in_v = false
              else
                raw_value = raw_value ? raw_value << value_buffer : value_buffer.dup
                collecting_value = false
              end
            elsif node.depth == row_depth
              current_cells = pad_row(current_cells, current_row_index, values_only: values_only) if pad_cells
              current_cells = expand_merged_cells(current_cells, current_row_index, values_only: values_only) if expand_merged
              yield values_only ? extract_values(current_cells).freeze : Row.new(index: current_row_index, cells: current_cells)
              last_row_index = current_row_index
              current_row_index = nil
              current_cells = nil
            elsif current_cells && node.depth == row_depth + 1
              current_cells << build_row_entry(cell_ref, coerce_value(raw_value, cell_type), values_only)
              cell_ref = nil
              cell_type = nil
              raw_value = nil
            end
          end
        end
      end
    end

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

    def extract_merge_ranges_by_row
      ranges_by_row = Hash.new { |hash, key| hash[key] = [] }

      with_sheet_reader do |reader|
        reader.each do |node|
          next unless node.node_type == ELEMENT_NODE && node.local_name == "mergeCell"

          range = parse_merge_range(node.attribute("ref"))
          next unless range

          (range[:start_row]..range[:end_row]).each do |row|
            ranges_by_row[row] << range
          end
        end
      end

      ranges_by_row
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

    def parse_merge_range(reference)
      return nil if reference.nil? || reference.empty?

      start_ref, finish_ref = reference.split(":", 2)
      finish_ref ||= start_ref
      start_col, start_row, end_col, end_row = *range_bounds(start_ref, finish_ref)
      return nil unless start_col && start_row && end_col && end_row

      {
        start_col: start_col,
        start_row: start_row,
        end_col: end_col,
        end_row: end_row
      }
    end

    def range_bounds(start_ref, finish_ref)
      start_col, start_row = split_coordinate(start_ref)
      finish_col, finish_row = split_coordinate(finish_ref)
      [start_col, start_row, finish_col, finish_row]
    end

    def split_coordinate(reference)
      col = 0
      i = 0
      len = reference.length

      while i < len
        byte = reference.getbyte(i)
        break unless byte >= 65 && byte <= 90 # A-Z

        col = (col * 26) + (byte - 64)
        i += 1
      end

      return [nil, nil] if i == 0 || i == len

      row = 0
      while i < len
        byte = reference.getbyte(i)
        return [nil, nil] unless byte >= 48 && byte <= 57 # 0-9

        row = (row * 10) + (byte - 48)
        i += 1
      end

      [col, row]
    end

    def column_index(label)
      col = 0
      i = 0
      len = label.length
      while i < len
        col = (col * 26) + (label.getbyte(i) - 64)
        i += 1
      end
      col
    end

    def split_col_index(reference)
      col = 0
      i = 0
      len = reference.length

      while i < len
        byte = reference.getbyte(i)
        break unless byte >= 65 && byte <= 90

        col = (col * 26) + (byte - 64)
        i += 1
      end

      col
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

    def expand_merged_cells(cells, row_index, values_only:)
      merge_ranges = merge_ranges_by_row[row_index]
      return cells if merge_ranges.empty?

      expanded_cells = cells.dup

      merge_ranges.each do |range|
        if row_index == range[:start_row]
          @merge_anchor_values[range] = value_at(expanded_cells, range[:start_col], values_only: values_only)
        end

        anchor_value = @merge_anchor_values[range]
        next if anchor_value.nil?

        (range[:start_col]..range[:end_col]).each do |col|
          next if row_index == range[:start_row] && col == range[:start_col]

          expanded_cells = set_value_at(expanded_cells, row_index, col, anchor_value, values_only: values_only)
        end
      end

      expanded_cells
    end

    def value_at(cells, col_index, values_only:)
      cell = cells[col_index - 1]
      return nil unless cell

      if values_only
        cell[1]
      elsif cell.is_a?(EmptyCell)
        nil
      else
        cell.value
      end
    end

    def set_value_at(cells, row_index, col_index, value, values_only:)
      if values_only
        coordinate = "#{column_name(col_index)}#{row_index}"
        cells[col_index - 1] = [coordinate, value]
      else
        coordinate = "#{column_name(col_index)}#{row_index}"
        cells[col_index - 1] = ReadOnlyCell.new(coordinate, value)
      end

      cells
    end

    def merge_ranges_by_row
      @merge_ranges_by_row ||= extract_merge_ranges_by_row
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

      numeric_kind = detect_numeric_kind(raw_value)
      return raw_value.to_i if numeric_kind == :integer
      return raw_value.to_f if numeric_kind == :float

      raw_value
    end

    def detect_numeric_kind(value)
      index = 0
      length = value.length
      saw_digit = false
      saw_dot = false

      if value.getbyte(0) == 45
        index = 1
        return nil if length == 1
      end

      while index < length
        byte = value.getbyte(index)

        if byte >= 48 && byte <= 57
          saw_digit = true
        elsif byte == 46
          return nil if saw_dot

          saw_dot = true
        else
          return nil
        end

        index += 1
      end

      return nil unless saw_digit

      saw_dot ? :float : :integer
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
