module Rbxl
  # Wraps a write-side cell value so that a style id and/or hyperlink can be
  # associated with it without widening every call site to a Hash or Array.
  #
  # Instances are passed transparently to {Rbxl::WriteOnlyWorksheet#append}
  # (or +<<+) in place of a plain value:
  #
  #   styled = Rbxl::WriteOnlyCell.new(42, style_id: 1)
  #   linked = Rbxl::WriteOnlyCell.new("repo", style_id: 1, hyperlink: "https://...")
  #   sheet << ["id", styled, linked]
  #
  # The value is serialized using the same type rules as a bare value; the
  # +style_id+, when present, is emitted as the cell's +s+ attribute. The
  # +hyperlink+, when present, is collected into the worksheet's
  # +<hyperlinks>+ element at save time. Targets starting with +#+ are
  # treated as internal sheet anchors (e.g. +"#'Detail'!A1"+); anything
  # else is treated as an external URL and gets a per-sheet
  # +_rels/sheetN.xml.rels+ entry.
  class WriteOnlyCell
    # @return [Object] underlying Ruby value (String, Numeric, Boolean,
    #   Date/DateTime/Time, or +nil+)
    attr_reader :value

    # @return [Integer, nil] style index into the workbook's +cellXfs+ table
    attr_reader :style_id

    # @return [String, nil] hyperlink target. Strings beginning with +#+
    #   are internal anchors; everything else is treated as an external URL.
    attr_reader :hyperlink

    # @param value [Object] Ruby value to serialize into the cell
    # @param style_id [Integer, nil] optional style index
    # @param hyperlink [String, nil] optional hyperlink target
    def initialize(value, style_id: nil, hyperlink: nil)
      @value = value
      @style_id = style_id
      @hyperlink = hyperlink
    end
  end
end
