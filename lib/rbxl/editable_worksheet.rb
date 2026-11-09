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
  # Table-shaped regions (invoice line items, report rows) can be filled in
  # one call with {#write_rows}.
  #
  # See {EditableWorkbook} for the design contract these edits live inside.
  class EditableWorksheet
    # Namespace for the main SpreadsheetML schema.
    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".freeze

    # @return [String] visible sheet name
    attr_reader :name

    # @return [String] ZIP entry path of the worksheet's XML part
    attr_reader :entry_path

    # Largest column index Excel accepts (+XFD+).
    MAX_COLUMN = 16_384

    # Largest row number Excel accepts.
    MAX_ROW = 1_048_576

    # @api private
    # Construct via {EditableWorkbook#sheet} / {EditableWorkbook#add_sheet}.
    #
    # @param workbook [Rbxl::EditableWorkbook] owning workbook
    # @param zip [Zip::File] open archive shared with the workbook
    # @param entry_path [String] ZIP entry path for this sheet's XML
    # @param workbook_path [String] filesystem path the workbook was opened from
    # @param shared_strings [Array<String>] pre-decoded shared strings table
    # @param name [String] visible sheet name
    # @param source_xml [String, nil] initial XML for a sheet that does not
    #   exist in the source ZIP yet (see {EditableWorkbook#add_sheet})
    def initialize(workbook:, zip:, entry_path:, workbook_path:, shared_strings:, name:, source_xml: nil)
      @workbook = workbook
      @zip = zip
      @entry_path = entry_path
      @workbook_path = workbook_path
      @shared_strings = shared_strings
      @name = name
      @source_xml = source_xml
      @doc = nil
      @sheet_data = nil
      @row_index = nil
      @dirty = false
      @modified = !source_xml.nil?
      @formula_removed = false
    end

    # Writes a 2D array of values into consecutive cells, top-left first.
    # Each value goes through {EditableCell#value=}, so the template's
    # per-cell styles are kept and +Date+ / +Time+ get date formatting.
    # +nil+ entries clear the corresponding cell.
    #
    # Pass a range (e.g. <tt>"B10:D19"</tt>) instead of a single anchor to
    # declare the template's capacity: writing more rows or wider rows than
    # the range holds raises before any cell is touched. Rows are never
    # inserted — growing a table past the template's prepared rows is out
    # of scope.
    #
    #   sheet.write_rows("B10:D19", [
    #     ["作業費", 1, 50_000],
    #     ["交通費", 1,  3_200],
    #   ])
    #
    # @param anchor [String] top-left coordinate, or an +A1:B2+-style range
    # @param rows [Array<Array>] row-major values
    # @return [Rbxl::EditableWorksheet] +self+
    # @raise [ArgumentError] if the data does not fit the range or the sheet
    def write_rows(anchor, rows)
      raise ArgumentError, "rows must be an Array of Arrays" unless rows.is_a?(Array) && rows.all? { |r| r.is_a?(Array) }

      start_col, start_row, max_rows, max_cols = parse_anchor(anchor)
      width = rows.map(&:length).max || 0

      if max_rows && rows.length > max_rows
        raise ArgumentError, "#{rows.length} rows do not fit #{anchor} (#{max_rows} rows)"
      end
      if max_cols && width > max_cols
        raise ArgumentError, "a #{width}-column row does not fit #{anchor} (#{max_cols} columns)"
      end
      if start_row + rows.length - 1 > MAX_ROW || start_col + width - 1 > MAX_COLUMN
        raise ArgumentError, "data starting at #{anchor} extends past the sheet's bounds"
      end

      rows.each_with_index do |values, r|
        values.each_with_index do |value, c|
          cell("#{column_label(start_col + c)}#{start_row + r}").value = value
        end
      end
      self
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
      @modified = true
    end

    # @api private
    # Whether the in-memory XML differs from the source ZIP. Unlike {#dirty?}
    # this survives {EditableWorkbook#save}: the source ZIP is never
    # rewritten, so every later save must keep emitting the edited XML.
    def modified?
      @modified
    end

    # @api private
    def note_formula_removed!
      @formula_removed = true
    end

    # @api private
    def formula_removed?
      @formula_removed
    end

    # @api private
    # Points lazy loads at a freshly reopened archive (after an in-place
    # save replaced the file the old handle was reading by name).
    def rebind_zip!(zip)
      @zip = zip
    end

    # @api private
    def styles
      @workbook.styles
    end

    # @api private
    def date_1904?
      @workbook.date_1904?
    end

    # @api private
    def clear_dirty!
      @dirty = false
    end

    # @return [String] the worksheet's XML, reflecting any in-memory edits.
    #   The XML declaration and original namespace bindings are preserved.
    def to_xml
      ensure_doc_loaded!
      XmlEdit.serialize(@doc)
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

      parsed = Nokogiri::XML(@source_xml || read_entry)
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

    def read_entry
      entry = @zip.find_entry(@entry_path)
      unless entry
        raise WorksheetFormatError,
              "worksheet #{@name.inspect} is missing XML entry #{@entry_path.inspect} in #{@workbook_path}"
      end

      entry.get_input_stream.read
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
      following ? XmlEdit.insert_before(following, xml) : XmlEdit.append(parent, xml)
    end

    def insert_cell_in_order(parent, coordinate, col_index)
      following = parent.element_children.find do |child|
        next false unless child.name == "c"

        child_col, _ = parse_coordinate(child["r"])
        child_col && child_col > col_index
      end
      xml = %(<c r="#{coordinate}"/>)
      following ? XmlEdit.insert_before(following, xml) : XmlEdit.append(parent, xml)
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

    RANGE_RE = /\A([A-Z]+)([1-9]\d*)(?::([A-Z]+)([1-9]\d*))?\z/.freeze
    private_constant :RANGE_RE

    # @return [Array(Integer, Integer, Integer|nil, Integer|nil)]
    #   start column, start row, and the range's row / column capacity
    def parse_anchor(anchor)
      m = anchor.to_s.upcase.match(RANGE_RE)
      raise ArgumentError, "invalid anchor: #{anchor.inspect}" unless m

      col1 = column_index(m[1])
      row1 = m[2].to_i
      return [col1, row1, nil, nil] unless m[3]

      col2 = column_index(m[3])
      row2 = m[4].to_i
      raise ArgumentError, "range must run top-left to bottom-right: #{anchor.inspect}" if col2 < col1 || row2 < row1

      [col1, row1, row2 - row1 + 1, col2 - col1 + 1]
    end

    def column_label(index)
      label = +""
      while index.positive?
        index, rem = (index - 1).divmod(26)
        label.prepend((65 + rem).chr)
      end
      label
    end

    def column_index(label)
      col = 0
      label.each_byte { |b| col = (col * 26) + (b - 64) }
      col
    end
  end
end
