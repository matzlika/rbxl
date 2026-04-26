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
  # * +Integer+ / +Float+ — number cell (no +t+ attribute)
  # * +String+ — inline string cell (+t="inlineStr"+); +xl/sharedStrings.xml+
  #   is never mutated, so this round-trips deterministically without a
  #   second pass over the SST
  # * +Date+ / +Time+ / +DateTime+ — raises {EditableCellTypeError}; convert
  #   to a numeric serial yourself if you need a date cell. Date support is
  #   intentionally deferred so 1.4.0 doesn't ship a half-baked numFmt write
  #
  # When overwriting an existing cell, the +s+ (style index) attribute is
  # preserved so template formatting (number format, font, fill, alignment)
  # carries through to the new value. Any +<f>+ (formula) and cached +<v>+
  # are dropped — assigning a value means the cell is no longer a formula.
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
    # @param new_value [String, Integer, Float, true, false, nil]
    # @return [Object] +new_value+
    # @raise [Rbxl::EditableCellTypeError] for unsupported types
    #   (+Date+/+Time+, arbitrary objects)
    def value=(new_value)
      reject_unsupported_type!(new_value)

      node = @worksheet.find_or_create_cell_node(@coordinate, create: true)
      apply_value(node, new_value)
      @worksheet.mark_dirty!
      new_value
    end

    private

    WHITESPACE_BYTES = [" ".ord, "\t".ord, "\n".ord, "\r".ord].freeze
    private_constant :WHITESPACE_BYTES

    def reject_unsupported_type!(value)
      case value
      when nil, true, false, Integer, Float, String
        # supported
      when Date, Time, DateTime
        raise EditableCellTypeError,
              "Date/Time/DateTime are not supported by EditableCell in 1.4.0; " \
              "convert to a numeric Excel serial yourself if you need a date cell"
      when Numeric
        # other Numerics (Rational, BigDecimal) — coerce to Float on apply
      else
        raise EditableCellTypeError,
              "unsupported cell value type: #{value.class}"
      end
    end

    def apply_value(node, value)
      node.children.unlink
      node.delete("t")

      case value
      when nil
        # empty cell — preserve <c r="..." s="..."/>
      when true
        node["t"] = "b"
        node.add_child("<v>1</v>")
      when false
        node["t"] = "b"
        node.add_child("<v>0</v>")
      when Integer
        node.add_child("<v>#{value}</v>")
      when Float
        # Ruby's Float#to_s gives the shortest round-trippable form. Excel
        # accepts standard decimal and scientific notation as <v> text.
        node.add_child("<v>#{value}</v>")
      when String
        node["t"] = "inlineStr"
        text = CGI.escapeHTML(value)
        space_attr = preserve_whitespace?(value) ? ' xml:space="preserve"' : ""
        node.add_child("<is><t#{space_attr}>#{text}</t></is>")
      when Numeric
        node.add_child("<v>#{value.to_f}</v>")
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
