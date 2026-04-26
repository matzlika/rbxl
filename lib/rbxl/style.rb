module Rbxl
  # Frozen value object describing a cell style registered via
  # {Rbxl::WriteOnlyWorkbook#add_style}.
  #
  # Each component is normalized to a frozen Hash with symbol keys (or
  # +nil+ when not specified) so that two styles built from equivalent input
  # hash equal and can share +fonts+/+fills+/+borders+/+numFmts+ entries in
  # the emitted +styles.xml+.
  #
  #   Rbxl::Style.new(
  #     font: { name: "Yu Gothic UI", size: 11, bold: true, color: "333333" },
  #     fill: { color: "EEEEEE" },
  #   )
  #
  # Colors are 6-character RRGGBB hex strings without the leading +#+.
  class Style
    # @return [Hash{Symbol=>Object}, nil] font attributes (+name+, +size+, +bold+, +italic+, +underline+, +color+)
    attr_reader :font

    # @return [Hash{Symbol=>Object}, nil] fill attributes (+color+ for solid pattern, or +pattern_type+)
    attr_reader :fill

    # @return [Hash{Symbol=>Object}, nil] border attributes; supports +all+ for four-sided
    #   shorthand or per-side +top+/+right+/+bottom+/+left+ keys, each as
    #   <tt>{ style:, color: }</tt>
    attr_reader :border

    # @return [Hash{Symbol=>Object}, nil] alignment attributes (+horizontal+, +vertical+, +wrap_text+, +text_rotation+)
    attr_reader :alignment

    # @return [String, nil] Excel number format code
    attr_reader :number_format

    # @param font [Hash, nil]
    # @param fill [Hash, nil]
    # @param border [Hash, nil]
    # @param alignment [Hash, nil]
    # @param number_format [String, nil]
    def initialize(font: nil, fill: nil, border: nil, alignment: nil, number_format: nil)
      @font = normalize_hash(font, :font)
      @fill = normalize_hash(fill, :fill)
      @border = normalize_border(border)
      @alignment = normalize_hash(alignment, :alignment)
      @number_format = number_format.nil? ? nil : number_format.to_s.dup.freeze
      freeze
    end

    private

    def normalize_hash(value, label)
      return nil if value.nil?
      raise ArgumentError, "#{label}: expected Hash, got #{value.class}" unless value.is_a?(Hash)

      value.each_with_object({}) { |(k, v), out| out[k.to_sym] = v }.freeze
    end

    def normalize_border(value)
      return nil if value.nil?
      raise ArgumentError, "border: expected Hash, got #{value.class}" unless value.is_a?(Hash)

      sym_keyed = value.each_with_object({}) { |(k, v), out| out[k.to_sym] = v }

      sides =
        if sym_keyed.key?(:all)
          spec = sym_keyed[:all]
          %i[top right bottom left].each_with_object({}) { |side, out| out[side] = spec }
        else
          %i[top right bottom left].each_with_object({}) do |side, out|
            out[side] = sym_keyed[side] if sym_keyed.key?(side)
          end
        end

      sides.transform_values { |spec| spec.nil? ? nil : normalize_hash(spec, :border).freeze }.freeze
    end
  end
end
