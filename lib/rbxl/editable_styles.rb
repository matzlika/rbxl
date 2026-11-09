module Rbxl
  # @api private
  #
  # Append-only view of an existing workbook's +xl/styles.xml+, used by edit
  # mode to register new cell styles without disturbing the template's own.
  #
  # Every operation only <em>appends</em> entries to +numFmts+ / +fonts+ /
  # +fills+ / +borders+ / +cellXfs+ and bumps their +count+ attributes, so the
  # indices that existing cells reference never change. The part is
  # re-serialized only when something was actually appended; otherwise it is
  # copied from the source ZIP byte-for-byte like any other untouched part.
  #
  # Two ways to obtain a new +cellXfs+ index:
  #
  # * {#register} — a style built from scratch, exactly like
  #   {WriteOnlyWorkbook#add_style}; unspecified components use entry 0
  # * {#patch} — a copy of an existing xf with only the given components
  #   overridden (font attributes merged, border sides merged, alignment
  #   attributes merged, fill and number format replaced)
  class EditableStyles
    # Child-element sequence of +<styleSheet>+ (ECMA-376 part 1 §18.8.39).
    STYLESHEET_ORDER = %w[numFmts fonts fills borders cellStyleXfs cellXfs
                          cellStyles dxfs tableStyles colors extLst].freeze

    # Conventional element order inside +<font>+, as written by Excel.
    FONT_ORDER = %w[b i strike condense extend outline shadow u vertAlign
                    sz color name family charset scheme].freeze

    # Child-element sequence of +<border>+.
    BORDER_ORDER = %w[start left end right top bottom diagonal vertical horizontal].freeze

    # Child-element sequence of +<xf>+.
    XF_ORDER = %w[alignment protection extLst].freeze

    # Locale-independent built-in number formats a +number_format+ string is
    # mapped to before a custom +<numFmt>+ is allocated. Date formats are
    # deliberately absent: built-in ids 14–22 render per the viewer's locale,
    # so mapping a literal code onto them would change what the user asked for.
    BUILTIN_NUMFMTS = {
      "General" => 0, "0" => 1, "0.00" => 2, "#,##0" => 3, "#,##0.00" => 4,
      "0%" => 9, "0.00%" => 10, "0.00E+00" => 11, "@" => 49
    }.freeze

    # First id available for custom number formats.
    NUMFMT_CUSTOM_BASE = 164

    # Minimal stylesheet used when a workbook ships without a styles part.
    MINIMAL_XML = <<~XML.freeze
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>
    XML

    # @return [String] ZIP entry path of the styles part
    attr_reader :part_name

    # @param xml [String] source +styles.xml+
    # @param part_name [String] ZIP entry path of the styles part
    # @param workbook_path [String] used in error messages
    def initialize(xml, part_name:, workbook_path:)
      @part_name = part_name
      @doc = Nokogiri::XML(xml)
      unless @doc.errors.empty? && @doc.root&.name == "styleSheet"
        detail = @doc.errors.first || "root element is not <styleSheet>"
        raise WorkbookFormatError, "#{workbook_path}: #{part_name}: #{detail}"
      end

      @root = @doc.root
      @modified = false
      @entry_cache = {}
      @patch_cache = {}
      @font_patch_cache = {}
      @border_patch_cache = {}
    end

    # @return [Boolean] whether any entry has been appended since load
    def modified?
      @modified
    end

    # @return [String] the stylesheet XML including appended entries
    def to_xml
      XmlEdit.serialize(@doc)
    end

    # @return [Integer] number of +cellXfs+ entries (valid style ids are
    #   +0...xf_count+)
    def xf_count
      entries("cellXfs", "xf").length
    end

    # Registers a style from scratch. Like {StyleRegistry#register}, each
    # call appends a new xf (no deduplication at the xf level), while
    # identical fonts / fills / borders appended in this session are shared.
    #
    # @param style [Rbxl::Style]
    # @return [Integer] the new +cellXfs+ index
    def register(style)
      xf = {
        font_id: style.font ? append_shared("fonts", "font", StyleXml.render_font(style.font)) : 0,
        fill_id: style.fill ? append_shared("fills", "fill", StyleXml.render_fill(style.fill)) : 0,
        border_id: style.border ? append_shared("borders", "border", StyleXml.render_border(style.border)) : 0,
        numfmt_id: style.number_format ? numfmt_id_for(style.number_format) : 0,
        alignment: style.alignment
      }
      append_entry("cellXfs", "xf", StyleXml.render_xf(xf))
    end

    # Derives a new xf from +base_index+ with the components of +style+
    # layered on top. Identical (base, style) pairs return the same index.
    #
    # @param base_index [Integer] existing +cellXfs+ index to start from
    # @param style [Rbxl::Style]
    # @param numfmt_id [Integer, nil] internal override for the number format
    #   id (used for date writes), taking precedence over +style.number_format+
    # @return [Integer] the derived +cellXfs+ index
    def patch(base_index, style, numfmt_id: nil)
      key = [base_index, style.font, style.fill, style.border, style.alignment, style.number_format, numfmt_id]
      @patch_cache[key] ||= build_patched_xf(base_index, style, numfmt_id)
    end

    # @param index [Integer] +cellXfs+ index
    # @return [Boolean] whether that xf renders numbers as a date/time
    def date_style?(index)
      xf = entries("cellXfs", "xf")[index]
      fmt_id = xf ? xf["numFmtId"].to_i : 0
      return true if DateFormat::BUILTIN_IDS.include?(fmt_id)

      code = entries("numFmts", "numFmt").find { |n| n["numFmtId"].to_i == fmt_id }&.[]("formatCode")
      code ? DateFormat.date_format_code?(code) : false
    end

    private

    def build_patched_xf(base_index, style, numfmt_id)
      base = entries("cellXfs", "xf")[base_index]
      source = base ? XmlEdit.serialize(base) : '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

      # Resolve sub-entries first so the xf appended below is the last one.
      font_id = style.font ? patched_font(base&.[]("fontId").to_i, style.font) : nil
      fill_id = style.fill ? append_shared("fills", "fill", StyleXml.render_fill(style.fill)) : nil
      border_id = style.border ? patched_border(base&.[]("borderId").to_i, style.border) : nil
      fmt_id = numfmt_id || (style.number_format && numfmt_id_for(style.number_format))

      index = append_entry("cellXfs", "xf", source)
      xf = entries("cellXfs", "xf")[index]

      set_component(xf, "fontId", "applyFont", font_id)
      set_component(xf, "fillId", "applyFill", fill_id)
      set_component(xf, "borderId", "applyBorder", border_id)
      set_component(xf, "numFmtId", "applyNumberFormat", fmt_id)
      merge_alignment(xf, style.alignment) if style.alignment
      index
    end

    def set_component(xf, id_attr, apply_attr, value)
      return if value.nil?

      xf[id_attr] = value.to_s
      xf[apply_attr] = "1"
    end

    def merge_alignment(xf, alignment)
      node = XmlEdit.child(xf, "alignment") ||
             XmlEdit.insert_in_order(xf, "alignment", "<alignment/>", XF_ORDER)
      StyleXml.alignment_attributes(alignment).each { |name, value| node[name] = value }
      node.delete("wrapText") if alignment.key?(:wrap_text) && !alignment[:wrap_text]
      xf["applyAlignment"] = "1"
    end

    def patched_font(base_id, font)
      @font_patch_cache[[base_id, font]] ||= begin
        base = entries("fonts", "font")[base_id]
        index = append_entry("fonts", "font", base ? XmlEdit.serialize(base) : "<font/>")
        node = entries("fonts", "font")[index]
        apply_font_attributes(node, font)
        index
      end
    end

    def apply_font_attributes(node, font)
      set_flag(node, "b", font[:bold]) if font.key?(:bold)
      set_flag(node, "i", font[:italic]) if font.key?(:italic)
      if font.key?(:underline)
        u = font[:underline]
        set_valued(node, "u", u.is_a?(String) || u.is_a?(Symbol) ? u.to_s : nil, present: !!u)
      end
      set_valued(node, "sz", font[:size].to_s, present: true) if font[:size]
      if font[:color]
        replace_child(node, "color", %(<color rgb="#{StyleXml.escape_attr(StyleXml.argb(font[:color]))}"/>), FONT_ORDER)
      end
      return unless font[:name]

      set_valued(node, "name", font[:name].to_s, present: true)
      # A <scheme> (major/minor) makes Excel substitute the theme font and
      # ignore <name>, so an explicit name must drop it.
      XmlEdit.children(node, "scheme").each(&:unlink)
    end

    def set_flag(node, name, on)
      XmlEdit.children(node, name).each(&:unlink)
      XmlEdit.insert_in_order(node, name, "<#{name}/>", FONT_ORDER) if on
    end

    def set_valued(node, name, val, present:)
      XmlEdit.children(node, name).each(&:unlink)
      return unless present

      xml = val ? %(<#{name} val="#{StyleXml.escape_attr(val)}"/>) : "<#{name}/>"
      XmlEdit.insert_in_order(node, name, xml, FONT_ORDER)
    end

    def replace_child(node, name, xml, order)
      existing = XmlEdit.children(node, name)
      if existing.empty?
        XmlEdit.insert_in_order(node, name, xml, order)
      else
        XmlEdit.insert_before(existing.first, xml)
        existing.each(&:unlink)
      end
    end

    def patched_border(base_id, border)
      @border_patch_cache[[base_id, border]] ||= begin
        base = entries("borders", "border")[base_id]
        index = append_entry("borders", "border", base ? XmlEdit.serialize(base) : "<border/>")
        node = entries("borders", "border")[index]
        border.each do |side, spec|
          replace_child(node, side.to_s, StyleXml.render_border_side(side, spec), BORDER_ORDER)
        end
        index
      end
    end

    def numfmt_id_for(code)
      code = code.to_s
      builtin = BUILTIN_NUMFMTS[code]
      return builtin if builtin

      existing = entries("numFmts", "numFmt").find { |n| n["formatCode"] == code }
      return existing["numFmtId"].to_i if existing

      used = entries("numFmts", "numFmt").map { |n| n["numFmtId"].to_i }
      id = [used.max || 0, NUMFMT_CUSTOM_BASE - 1].max + 1
      append_entry("numFmts", "numFmt", %(<numFmt numFmtId="#{id}" formatCode="#{StyleXml.escape_attr(code)}"/>))
      id
    end

    # Appends +xml+ to +container+ unless the identical markup was already
    # appended in this session, in which case the earlier index is reused.
    def append_shared(container, entry, xml)
      @entry_cache[[container, xml]] ||= append_entry(container, entry, xml)
    end

    def append_entry(container_name, entry_name, xml)
      container = container(container_name)
      index = XmlEdit.children(container, entry_name).length
      XmlEdit.append(container, xml)
      container["count"] = (index + 1).to_s
      @modified = true
      index
    end

    def container(name)
      XmlEdit.child(@root, name) ||
        XmlEdit.insert_in_order(@root, name, %(<#{name} count="0"/>), STYLESHEET_ORDER)
    end

    def entries(container_name, entry_name)
      container = XmlEdit.child(@root, container_name)
      container ? XmlEdit.children(container, entry_name) : []
    end
  end
end
