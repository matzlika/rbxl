module Rbxl
  # Immutable row wrapper yielded by {Rbxl::ReadOnlyWorksheet#each_row}.
  #
  # A row holds its 1-based worksheet index and a frozen array of cell
  # objects. The cell array may contain {Rbxl::Cell}, {Rbxl::ReadOnlyCell},
  # or {Rbxl::EmptyCell} instances depending on the iteration options
  # (+pad_cells+, +expand_merged+) and the parser backend in use.
  #
  #   sheet.each_row do |row|
  #     row.index       # => 2
  #     row.size        # => 3
  #     row.values      # => ["alice", 100, true]
  #     row[0].value    # => "alice"
  #   end
  class Row
    # @return [Integer] 1-based worksheet row number
    attr_reader :index

    # @return [Array<Rbxl::Cell, Rbxl::ReadOnlyCell, Rbxl::EmptyCell>]
    #   frozen array of cell objects
    attr_reader :cells

    # @param index [Integer] 1-based worksheet row number
    # @param cells [Array<Rbxl::Cell, Rbxl::ReadOnlyCell, Rbxl::EmptyCell>]
    #   cell objects in column order; the array is frozen in place
    def initialize(index:, cells:)
      @index = index
      @cells = cells.freeze
      @values = nil
    end

    # Returns the cell at a zero-based offset within the row.
    #
    # No bounds checking is performed beyond Array semantics: an offset
    # outside the cell range simply returns +nil+.
    #
    # @param offset [Integer] zero-based position within the row
    # @return [Rbxl::Cell, Rbxl::ReadOnlyCell, Rbxl::EmptyCell, nil]
    def [](offset)
      cells[offset]
    end

    # Returns the row as plain Ruby values, memoized and frozen so that
    # repeated calls are allocation-free.
    #
    # @return [Array<Object>] decoded cell values in column order
    def values
      @values ||= cells.map(&:value).freeze
    end

    # @return [Integer] number of cells in the row
    def size
      cells.size
    end
  end
end
