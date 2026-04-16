module Rbxl
  # Immutable cell value object used by the read-only worksheet path.
  #
  # Produced during streaming iteration when cells are yielded without
  # +values_only+. Implemented as a +Data+ class so instances are frozen and
  # hash-equal by value.
  #
  # @!attribute [r] coordinate
  #   @return [String] Excel-style coordinate such as +"A1"+
  # @!attribute [r] value
  #   @return [Object, nil] decoded Ruby value (String, Numeric, Boolean, or +nil+)
  ReadOnlyCell = Data.define(:coordinate, :value)
end
