module Rbxl
  # Generic value-object cell used by the pure-Ruby reader path.
  #
  # Yielded as an element of {Rbxl::Row#cells} when a worksheet is iterated
  # without +values_only+. Cells are keyword-constructed and expose the
  # decoded Ruby value plus the Excel-style coordinate.
  #
  #   cell = Rbxl::Cell.new(value: 42, coordinate: "B3")
  #   cell.value      # => 42
  #   cell.coordinate # => "B3"
  #
  # @!attribute [rw] value
  #   @return [Object] decoded Ruby value for the cell (String, Numeric,
  #     Boolean, or +nil+)
  # @!attribute [rw] coordinate
  #   @return [String, nil] Excel-style coordinate such as +"B3"+
  Cell = Struct.new(:value, :coordinate, keyword_init: true)
end
