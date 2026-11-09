module Rbxl
  # Collects {Rbxl::Style} entries registered via
  # {Rbxl::WriteOnlyWorkbook#add_style} and emits the corresponding
  # +xl/styles.xml+ document.
  #
  # Each call to {#register} appends a new <tt>cellXf</tt> entry and returns
  # the 1-based +style_id+ to use on cells (index 0 is reserved for the
  # default xf so authored +style_id+ values match what Excel sees on disk).
  # Sub-tables (+fonts+, +fills+, +borders+, +numFmts+) are deduplicated by
  # content while keeping the +cellXfs+ ordering deterministic.
  class StyleRegistry
    # First custom +numFmtId+ allocated to user-supplied +number_format+
    # strings; OOXML reserves ids below 164 for built-in formats.
    NUMFMT_CUSTOM_BASE = 164

    def initialize
      @styles = []
      @fonts = []
      @fills = []
      @borders = []
      @numfmts = []
      @font_index = {}
      @fill_index = {}
      @border_index = {}
      @numfmt_index = {}
    end

    # @return [Boolean] true if no styles have been registered
    def empty?
      @styles.empty?
    end

    # Registers a style and returns its +cellXfs+ index.
    #
    # @param style [Rbxl::Style]
    # @return [Integer] 1-based style id
    def register(style)
      @styles << resolve(style)
      @styles.length
    end

    # @return [String] full +xl/styles.xml+ document
    def to_xml
      buf = +'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      buf << "\n"
      buf << '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'

      append_numfmts(buf)
      append_fonts(buf)
      append_fills(buf)
      append_borders(buf)

      buf << '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'

      append_cell_xfs(buf)

      buf << '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      buf << '</styleSheet>'
      buf
    end

    private

    def resolve(style)
      {
        font_id: style.font ? intern(@fonts, @font_index, style.font) : 0,
        fill_id: style.fill ? intern_fill(style.fill) : 0,
        border_id: style.border ? intern_border(style.border) : 0,
        numfmt_id: style.number_format ? intern_numfmt(style.number_format) : 0,
        alignment: style.alignment
      }
    end

    def intern(table, index, key)
      cached = index[key]
      return cached if cached

      table << key
      index[key] = table.length
    end

    def intern_fill(fill_hash)
      cached = @fill_index[fill_hash]
      return cached if cached

      @fills << fill_hash
      # +2 because fillId 0 (none) and fillId 1 (gray125) are conventionally reserved
      slot = @fills.length + 1
      @fill_index[fill_hash] = slot
    end

    def intern_border(border_hash)
      cached = @border_index[border_hash]
      return cached if cached

      @borders << border_hash
      @border_index[border_hash] = @borders.length
    end

    def intern_numfmt(format_code)
      cached = @numfmt_index[format_code]
      return cached if cached

      @numfmts << format_code
      slot = NUMFMT_CUSTOM_BASE + @numfmts.length - 1
      @numfmt_index[format_code] = slot
    end

    def append_numfmts(buf)
      return if @numfmts.empty?

      buf << %(<numFmts count="#{@numfmts.length}">)
      @numfmts.each_with_index do |code, i|
        id = NUMFMT_CUSTOM_BASE + i
        buf << %(<numFmt numFmtId="#{id}" formatCode="#{escape_attr(code)}"/>)
      end
      buf << '</numFmts>'
    end

    def append_fonts(buf)
      total = @fonts.length + 1 # +1 for default font 0
      buf << %(<fonts count="#{total}">)
      buf << '<font><sz val="11"/><name val="Calibri"/></font>'
      @fonts.each { |f| buf << render_font(f) }
      buf << '</fonts>'
    end

    def append_fills(buf)
      total = @fills.length + 2 # +2 for built-in none and gray125
      buf << %(<fills count="#{total}">)
      buf << '<fill><patternFill patternType="none"/></fill>'
      buf << '<fill><patternFill patternType="gray125"/></fill>'
      @fills.each { |f| buf << render_fill(f) }
      buf << '</fills>'
    end

    def append_borders(buf)
      total = @borders.length + 1 # +1 for default empty border 0
      buf << %(<borders count="#{total}">)
      buf << '<border><left/><right/><top/><bottom/><diagonal/></border>'
      @borders.each { |b| buf << render_border(b) }
      buf << '</borders>'
    end

    def append_cell_xfs(buf)
      total = @styles.length + 1 # +1 for default xf 0
      buf << %(<cellXfs count="#{total}">)
      buf << '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      @styles.each { |xf| buf << render_xf(xf) }
      buf << '</cellXfs>'
    end

    def render_font(font) = StyleXml.render_font(font)
    def render_fill(fill) = StyleXml.render_fill(fill)
    def render_border(border) = StyleXml.render_border(border)
    def render_xf(xf) = StyleXml.render_xf(xf)
    def escape_attr(value) = StyleXml.escape_attr(value)
  end
end
