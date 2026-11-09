module Rbxl
  # @api private
  #
  # Renders the individual +styles.xml+ entries (+<font>+, +<fill>+,
  # +<border>+, +<xf>+) for a normalized {Rbxl::Style} component. Shared by
  # the write-only {StyleRegistry} (which emits a whole +styles.xml+) and the
  # edit-mode {EditableStyles} (which appends entries to an existing one), so
  # both modes produce identical markup for identical input.
  module StyleXml
    module_function

    def render_font(font)
      attrs = +'<font>'
      attrs << '<b val="1"/>' if font[:bold]
      attrs << '<i val="1"/>' if font[:italic]
      attrs << '<u/>' if font[:underline]
      attrs << %(<sz val="#{font[:size]}"/>) if font[:size]
      attrs << %(<color rgb="#{escape_attr(argb(font[:color]))}"/>) if font[:color]
      attrs << %(<name val="#{escape_attr(font[:name])}"/>) if font[:name]
      attrs << '</font>'
      attrs
    end

    def render_fill(fill)
      pattern = fill[:pattern_type] || (fill[:color] ? 'solid' : 'none')
      buf = +'<fill><patternFill patternType="'
      buf << escape_attr(pattern)
      buf << '"'
      if fill[:color] && pattern != 'none'
        buf << '><fgColor rgb="'
        buf << escape_attr(argb(fill[:color]))
        buf << '"/></patternFill></fill>'
      else
        buf << '/></fill>'
      end
      buf
    end

    def render_border_side(side, spec)
      return %(<#{side}/>) unless spec

      buf = %(<#{side} style="#{escape_attr(spec[:style] || 'thin')}">)
      buf << %(<color rgb="#{escape_attr(argb(spec[:color]))}"/>) if spec[:color]
      buf << %(</#{side}>)
      buf
    end

    def render_border(border)
      buf = +'<border>'
      %i[left right top bottom].each { |side| buf << render_border_side(side, border[side]) }
      buf << '<diagonal/></border>'
      buf
    end

    def render_xf(xf)
      attrs = +'<xf'
      attrs << %( numFmtId="#{xf[:numfmt_id]}")
      attrs << %( fontId="#{xf[:font_id]}")
      attrs << %( fillId="#{xf[:fill_id]}")
      attrs << %( borderId="#{xf[:border_id]}")
      attrs << ' xfId="0"'
      attrs << ' applyFont="1"' if xf[:font_id].positive?
      attrs << ' applyFill="1"' if xf[:fill_id].positive?
      attrs << ' applyBorder="1"' if xf[:border_id].positive?
      attrs << ' applyNumberFormat="1"' if xf[:numfmt_id].positive?
      attrs << ' applyAlignment="1"' if xf[:alignment]

      if xf[:alignment]
        attrs << '>'
        attrs << render_alignment(xf[:alignment])
        attrs << '</xf>'
      else
        attrs << '/>'
      end
      attrs
    end

    def render_alignment(alignment)
      buf = +'<alignment'
      alignment_attributes(alignment).each { |name, value| buf << %( #{name}="#{escape_attr(value)}") }
      buf << '/>'
      buf
    end

    # @return [Array<Array(String, String)>] OOXML attribute name/value pairs
    #   for an alignment Hash, in a stable order
    def alignment_attributes(alignment)
      pairs = []
      pairs << ["horizontal", alignment[:horizontal].to_s] if alignment[:horizontal]
      pairs << ["vertical", alignment[:vertical].to_s] if alignment[:vertical]
      pairs << ["wrapText", "1"] if alignment[:wrap_text]
      pairs << ["textRotation", alignment[:text_rotation].to_i.to_s] if alignment[:text_rotation]
      pairs
    end

    # Normalizes a user color to the 8-digit ARGB form OOXML specifies:
    # a leading +#+ is dropped and 6-digit +RRGGBB+ gains an opaque +FF+
    # alpha. Other lengths pass through unchanged.
    def argb(color)
      hex = color.to_s.upcase.delete_prefix("#")
      hex.length == 6 ? "FF#{hex}" : hex
    end

    def escape_attr(value)
      str = value.to_s
      str.gsub(/[&<>"]/) { |c| { '&' => '&amp;', '<' => '&lt;', '>' => '&gt;', '"' => '&quot;' }[c] }
    end
  end
end
