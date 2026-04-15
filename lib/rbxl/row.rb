module Rbxl
  class Row
    attr_reader :index, :cells

    def initialize(index:, cells:)
      @index = index
      @cells = cells.freeze
      @values = nil
    end

    def [](offset)
      cells[offset]
    end

    def values
      @values ||= cells.map(&:value).freeze
    end

    def size
      cells.size
    end
  end
end
