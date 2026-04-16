module Rbxl
  # Wraps a write-side cell value so that a style id can be associated with
  # it without widening every call site to a Hash or Array.
  #
  # Instances are passed transparently to {Rbxl::WriteOnlyWorksheet#append}
  # (or +<<+) in place of a plain value:
  #
  #   cell = Rbxl::WriteOnlyCell.new(42, style_id: 1)
  #   sheet << ["id", cell]
  #
  # The value is serialized using the same type rules as a bare value; the
  # +style_id+, when present, is emitted as the cell's +s+ attribute.
  class WriteOnlyCell
    # @return [Object] underlying Ruby value (String, Numeric, Boolean,
    #   Date/DateTime/Time, or +nil+)
    attr_reader :value

    # @return [Integer, nil] style index into the workbook's +cellXfs+ table
    attr_reader :style_id

    # @param value [Object] Ruby value to serialize into the cell
    # @param style_id [Integer, nil] optional style index
    def initialize(value, style_id: nil)
      @value = value
      @style_id = style_id
    end
  end
end
