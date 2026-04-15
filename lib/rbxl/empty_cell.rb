module Rbxl
  class EmptyCell
    attr_reader :coordinate

    def initialize(coordinate:)
      @coordinate = coordinate
    end

    def value
      nil
    end
  end
end
