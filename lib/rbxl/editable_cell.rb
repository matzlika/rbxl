module Rbxl
  # A view onto a single +<c>+ element inside an {EditableWorksheet}.
  #
  # Cells are not stored — each call to {EditableWorksheet#cell} returns a
  # fresh {EditableCell} that resolves the underlying +<c>+ node on demand.
  # Reads decode the current XML; writes mutate the worksheet's DOM and
  # mark the sheet dirty so the next {EditableWorkbook#save} re-serializes
  # it.
  #
  # == Type matrix on write
  #
  # * +nil+ — clears the cell's value (children + +t+ attribute removed),
  #   leaving an empty +<c>+ that retains its +s+ (style index)
  # * +true+ / +false+ — boolean cell (+t="b"+)
  # * +Integer+ / +Float+ — number cell (no +t+ attribute); other +Numeric+
  #   values are written as +Float+
  # * +String+ — inline string cell (+t="inlineStr"+); +xl/sharedStrings.xml+
  #   is never mutated, so this round-trips deterministically without a
  #   second pass over the SST
  # * +Date+ / +Time+ / +DateTime+ — Excel serial number (respecting the
  #   workbook's 1900/1904 date system). If the cell's current style already
  #   has a date/time number format it is kept as-is; otherwise a copy of the
  #   current style is derived with the built-in short-date format (+Date+)
  #   or date-time format (+Time+ / +DateTime+), so fonts, fills, and borders
  #   from the template survive
  #
  # When overwriting an existing cell, the +s+ (style index) attribute is
  # preserved so template formatting (number format, font, fill, alignment)
  # carries through to the new value. Any +<f>+ (formula) and cached +<v>+
  # are dropped — assigning a value means the cell is no longer a formula.
  #
  # == Styling
  #
  # {#style} layers changes on top of the cell's current style (only the
  # attributes you pass change); {#style_id=} replaces the style wholesale
  # with an index from {EditableWorkbook#add_style}.
  class EditableCell
    # Namespace for the main SpreadsheetML schema.
    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".freeze

    # @return [String] Excel-style coordinate, e.g. +"B5"+
    attr_reader :coordinate

    # @api private
    # Construct via {EditableWorksheet#cell}; not for direct use.
    #
    # @param worksheet [EditableWorksheet]
    # @param coordinate [String] already-normalized +A1+-style coordinate
    def initialize(worksheet:, coordinate:)
      @worksheet = worksheet
      @coordinate = coordinate
    end

    # Decodes the current value of the cell.
    #
    # @return [String, Integer, Float, true, false, nil] the cell's value, or
    #   +nil+ if the cell does not exist or has no stored value. Boolean
    #   cells return +true+/+false+; numeric cells return +Integer+ when the
    #   stored value is integer-shaped, +Float+ otherwise; +t="s"+ cells
    #   resolve through the workbook's shared strings table; +t="inlineStr"+
    #   and +t="str"+ cells return the literal text
    def value
      node = @worksheet.find_or_create_cell_node(@coordinate, create: false)
      return nil unless node

      decode(node)
    end

    # Sets the cell's value. See the class-level "Type matrix on write"
    # documentation for accepted Ruby types and how each is serialized.
    #
    # @param new_value [String, Numeric, true, false, Date, Time, DateTime, nil]
    # @return [Object] +new_value+
    # @raise [Rbxl::EditableCellTypeError] for unsupported types, or for
    #   dates before the workbook's date epoch
    def value=(new_value)
      reject_unsupported_type!(new_value)
      serial = date_serial(new_value) if date_like?(new_value)

      node = @worksheet.find_or_create_cell_node(@coordinate, create: true)
      @worksheet.note_formula_removed! if node.element_children.any? { |c| c.name == "f" }

      if serial
        apply_value(node, serial)
        ensure_date_style(node, new_value)
      else
        apply_value(node, new_value)
      end
      @worksheet.mark_dirty!
      new_value
    end

    # @return [Integer] the cell's style index (+s+ attribute); +0+ when the
    #   cell has none or does not exist
    def style_id
      node = @worksheet.find_or_create_cell_node(@coordinate, create: false)
      node ? node["s"].to_i : 0
    end

    # Replaces the cell's style with +id+, typically a value returned by
    # {EditableWorkbook#add_style}. The cell's value is left untouched; the
    # cell is created (empty) if it does not exist yet.
    #
    # @param id [Integer] a +cellXfs+ index valid for this workbook
    # @return [Integer] +id+
    # @raise [ArgumentError] if +id+ is not an existing style index
    def style_id=(id)
      count = @worksheet.styles.xf_count
      unless id.is_a?(Integer) && id >= 0 && id < count
        raise ArgumentError, "style_id must be an Integer in 0...#{count}, got #{id.inspect}"
      end

      node = @worksheet.find_or_create_cell_node(@coordinate, create: true)
      if id.zero?
        node.delete("s")
      else
        node["s"] = id.to_s
      end
      @worksheet.mark_dirty!
      id
    end

    # Changes part of the cell's formatting while keeping the rest of its
    # current style (typically the template's). Accepts the same keywords as
    # {EditableWorkbook#add_style}, with merge semantics:
    #
    # * +font:+ — only the given attributes change (e.g. <tt>{ color: "C00000" }</tt>
    #   keeps the template's font face and size); +false+ turns +bold+ /
    #   +italic+ / +underline+ off
    # * +border:+ — only the given sides change; <tt>{ bottom: nil }</tt>
    #   removes the bottom line
    # * +alignment:+ — only the given attributes change
    # * +fill:+, +number_format:+ — replaced
    #
    # Identical changes on cells sharing a style reuse one derived style.
    #
    #   sheet["F20"].style(font: { color: "C00000", bold: true })
    #
    # @return [Rbxl::EditableCell] +self+
    def style(font: nil, fill: nil, border: nil, alignment: nil, number_format: nil)
      patch = Style.new(font: font, fill: fill, border: border, alignment: alignment, number_format: number_format)
      self.style_id = @worksheet.styles.patch(style_id, patch)
      self
    end

    private

    WHITESPACE_BYTES = [" ".ord, "\t".ord, "\n".ord, "\r".ord].freeze
    private_constant :WHITESPACE_BYTES

    def reject_unsupported_type!(value)
      case value
      when nil, true, false, Numeric, String, Date, Time
        # supported (DateTime is a Date)
      else
        raise EditableCellTypeError,
              "unsupported cell value type: #{value.class}"
      end
    end

    def date_like?(value)
      value.is_a?(Date) || value.is_a?(Time)
    end

    def date_serial(value)
      DateFormat.to_serial(value, date_1904: @worksheet.date_1904?)
    rescue ArgumentError => e
      raise EditableCellTypeError, "#{@coordinate}: #{e.message}"
    end

    def ensure_date_style(node, value)
      current = node["s"].to_i
      styles = @worksheet.styles
      return if styles.date_style?(current)

      fmt = value.is_a?(Time) || value.is_a?(DateTime) ? DateFormat::BUILTIN_DATETIME_ID : DateFormat::BUILTIN_DATE_ID
      node["s"] = styles.patch(current, Style.new, numfmt_id: fmt).to_s
    end

    def apply_value(node, value)
      node.children.unlink
      node.delete("t")

      case value
      when nil
        # empty cell — preserve <c r="..." s="..."/>
      when true
        node["t"] = "b"
        XmlEdit.append(node, "<v>1</v>")
      when false
        node["t"] = "b"
        XmlEdit.append(node, "<v>0</v>")
      when Integer
        XmlEdit.append(node, "<v>#{value}</v>")
      when Float
        # Ruby's Float#to_s gives the shortest round-trippable form. Excel
        # accepts standard decimal and scientific notation as <v> text.
        XmlEdit.append(node, "<v>#{value}</v>")
      when String
        node["t"] = "inlineStr"
        text = CGI.escapeHTML(value)
        space_attr = preserve_whitespace?(value) ? ' xml:space="preserve"' : ""
        XmlEdit.append(node, "<is><t#{space_attr}>#{text}</t></is>")
      when Numeric
        XmlEdit.append(node, "<v>#{value.to_f}</v>")
      end
    end

    def decode(node)
      type = node["t"]
      case type
      when "s"
        text = first_text_at(node, "v")
        text ? @worksheet.shared_string_at(text.to_i) : nil
      when "inlineStr"
        decode_inline_string(node)
      when "str"
        first_text_at(node, "v")
      when "b"
        first_text_at(node, "v") == "1"
      when "e"
        first_text_at(node, "v")
      else
        raw = first_text_at(node, "v")
        decode_numeric(raw)
      end
    end

    def first_text_at(node, local_name)
      child = node.at_xpath("./main:#{local_name}", "main" => MAIN_NS)
      child&.text
    end

    def decode_inline_string(node)
      is_node = node.at_xpath("./main:is", "main" => MAIN_NS)
      return nil unless is_node

      is_node.xpath(".//main:t", "main" => MAIN_NS).map(&:text).join
    end

    def decode_numeric(raw)
      return nil if raw.nil? || raw.empty?

      if raw.match?(/\A-?\d+\z/)
        raw.to_i
      elsif raw.match?(/\A-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\z/)
        raw.to_f
      else
        raw
      end
    end

    def preserve_whitespace?(string)
      return false if string.empty?

      WHITESPACE_BYTES.include?(string.getbyte(0)) ||
        WHITESPACE_BYTES.include?(string.getbyte(string.bytesize - 1))
    end
  end
end
