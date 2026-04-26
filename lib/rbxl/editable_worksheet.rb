module Rbxl
  # A single worksheet inside an {EditableWorkbook}.
  #
  # The worksheet's XML payload is parsed lazily — calling {#cell} for the
  # first time triggers a single Nokogiri DOM parse of the sheet entry, and
  # subsequent edits mutate that in-memory tree. Worksheets that are never
  # touched are never parsed; on save they pass through the ZIP unchanged.
  #
  # Cell access is openpyxl-style:
  #
  #   sheet["B5"].value = "company name"
  #   sheet.cell("B5").value      # => "company name"
  #
  # See {EditableWorkbook} for the design contract these edits live inside.
  class EditableWorksheet
    # Namespace for the main SpreadsheetML schema.
    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".freeze

    # @return [String] visible sheet name
    attr_reader :name

    # @return [String] ZIP entry path of the worksheet's XML part
    attr_reader :entry_path

    # @param zip [Zip::File] open archive shared with the workbook
    # @param entry_path [String] ZIP entry path for this sheet's XML
    # @param workbook_path [String] filesystem path the workbook was opened from
    # @param shared_strings [Array<String>] pre-decoded shared strings table
    # @param name [String] visible sheet name
    def initialize(zip:, entry_path:, workbook_path:, shared_strings:, name:)
      @zip = zip
      @entry_path = entry_path
      @workbook_path = workbook_path
      @shared_strings = shared_strings
      @name = name
      @doc = nil
      @sheet_data = nil
      @row_index = nil
      @dirty = false
    end

    # Returns the {EditableCell} view for +coordinate+. Cells not present in
    # the sheet's XML are addressable too — reading their value yields +nil+,
    # writing creates the +<c>+ (and its enclosing +<row>+ if needed) in
    # column-sorted position. Repeated calls for the same coordinate may
    # return different {EditableCell} objects but the underlying XML is the
    # same, so reads are consistent.
    #
    # @param coordinate [String] Excel-style coordinate (e.g. +"A1"+, +"B5"+)
    # @return [Rbxl::EditableCell]
    # @raise [ArgumentError] if +coordinate+ is not a valid +A1+-style ref
    def cell(coordinate)
      EditableCell.new(worksheet: self, coordinate: normalize_coordinate(coordinate))
    end

    alias [] cell

    # @return [Boolean] whether any cell on this sheet has been mutated since
    #   load (or since the last successful save)
    def dirty?
      @dirty
    end

    # Marks the sheet dirty. Called by {EditableCell#value=}; not part of
    # the public API.
    #
    # @api private
    def mark_dirty!
      @dirty = true
    end

    # @api private
    def clear_dirty!
      @dirty = false
    end

    # @return [String] the worksheet's XML, reflecting any in-memory edits.
    #   The XML declaration and original namespace bindings are preserved.
    def to_xml
      ensure_doc_loaded!
      @doc.to_xml
    end

    # @api private
    # Resolves a shared-string index against the table loaded from
    # +xl/sharedStrings.xml+. Used by {EditableCell} when decoding +t="s"+
    # cells.
    def shared_string_at(index)
      @shared_strings[index]
    end

    # @api private
    # Locates the +<c>+ node for +coordinate+. With +create: true+ the
    # node — and its enclosing +<row>+ — are inserted in sorted position
    # when missing. Returns +nil+ when +create+ is false and the cell does
    # not exist.
    def find_or_create_cell_node(coordinate, create:)
      ensure_doc_loaded!
      col, row = parse_coordinate(coordinate)
      raise ArgumentError, "invalid coordinate: #{coordinate.inspect}" unless col && row

      row_node = find_or_create_row(row, create: create)
      return nil unless row_node

      existing = row_node.element_children.find { |c| c["r"] == coordinate }
      return existing if existing
      return nil unless create

      insert_cell_in_order(row_node, coordinate, col)
    end

    # @api private
    # Returns the document for in-place mutation. Loads the XML on first
    # access.
    def document
      ensure_doc_loaded!
      @doc
    end

    private

    def ensure_doc_loaded!
      return if @doc

      entry = @zip.find_entry(@entry_path)
      unless entry
        raise WorksheetFormatError,
              "worksheet #{@name.inspect} is missing XML entry #{@entry_path.inspect} in #{@workbook_path}"
      end

      parsed = Nokogiri::XML(entry.get_input_stream.read)
      unless parsed.errors.empty?
        raise WorksheetFormatError,
              "invalid worksheet XML for sheet #{@name.inspect} in #{@workbook_path}: #{parsed.errors.first}"
      end

      sheet_data = parsed.at_xpath("/main:worksheet/main:sheetData", "main" => MAIN_NS)
      unless sheet_data
        raise WorksheetFormatError,
              "worksheet #{@name.inspect} in #{@workbook_path} is missing <sheetData>"
      end

      @doc = parsed
      @sheet_data = sheet_data
      @row_index = sheet_data.xpath("./main:row", "main" => MAIN_NS).each_with_object({}) do |row, h|
        idx = row["r"]&.to_i
        h[idx] = row if idx
      end
    end

    def find_or_create_row(row_num, create:)
      existing = @row_index[row_num]
      return existing if existing
      return nil unless create

      row_node = insert_row_in_order(@sheet_data, row_num)
      @row_index[row_num] = row_node
      row_node
    end

    # Insertion is done by parsing an XML fragment in the parent's context
    # so the new element inherits the SpreadsheetML default namespace
    # binding from its surroundings rather than landing in +xmlns=""+ jail.
    def insert_row_in_order(parent, row_num)
      following = parent.element_children.find do |child|
        child.name == "row" && (child["r"]&.to_i || 0) > row_num
      end
      xml = %(<row r="#{row_num}"/>)
      added = following ? following.add_previous_sibling(xml) : parent.add_child(xml)
      first_node(added)
    end

    def insert_cell_in_order(parent, coordinate, col_index)
      following = parent.element_children.find do |child|
        next false unless child.name == "c"

        child_col, _ = parse_coordinate(child["r"])
        child_col && child_col > col_index
      end
      xml = %(<c r="#{coordinate}"/>)
      added = following ? following.add_previous_sibling(xml) : parent.add_child(xml)
      first_node(added)
    end

    def first_node(result)
      result.is_a?(Nokogiri::XML::NodeSet) ? result.first : result
    end

    COORDINATE_RE = /\A([A-Z]+)([1-9]\d*)\z/.freeze
    private_constant :COORDINATE_RE

    def normalize_coordinate(coordinate)
      raise ArgumentError, "coordinate cannot be nil" if coordinate.nil?

      str = coordinate.to_s.upcase
      raise ArgumentError, "invalid coordinate: #{coordinate.inspect}" unless str.match?(COORDINATE_RE)

      str
    end

    def parse_coordinate(coordinate)
      return [nil, nil] unless coordinate

      m = coordinate.match(COORDINATE_RE)
      return [nil, nil] unless m

      [column_index(m[1]), m[2].to_i]
    end

    def column_index(label)
      col = 0
      label.each_byte { |b| col = (col * 26) + (b - 64) }
      col
    end
  end
end
