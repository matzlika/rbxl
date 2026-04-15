module Rbxl
  class WriteOnlyCell
    attr_reader :value, :style_id

    def initialize(value, style_id: nil)
      @value = value
      @style_id = style_id
    end
  end
end
