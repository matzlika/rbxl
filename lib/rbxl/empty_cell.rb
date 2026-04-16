module Rbxl
  # Placeholder cell returned when a coordinate in a padded row has no data.
  #
  # Used only when {Rbxl::ReadOnlyWorksheet#each_row} is called with
  # <tt>pad_cells: true</tt>. The object carries the synthetic coordinate so
  # that downstream code can still locate the slot in the worksheet grid.
  #
  #   cell = Rbxl::EmptyCell.new(coordinate: "C5")
  #   cell.coordinate # => "C5"
  #   cell.value      # => nil
  class EmptyCell
    # @return [String] Excel-style coordinate such as +"C5"+
    attr_reader :coordinate

    # @param coordinate [String] Excel-style coordinate
    def initialize(coordinate:)
      @coordinate = coordinate
    end

    # Always +nil+; exposed so callers can treat {EmptyCell} like any other
    # cell object without a type check.
    #
    # @return [nil]
    def value
      nil
    end
  end
end
