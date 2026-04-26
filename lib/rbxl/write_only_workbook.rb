module Rbxl
  # Write-only workbook for single-pass XLSX generation.
  #
  # The workbook accumulates rows per worksheet and emits the full
  # <tt>.xlsx</tt> package in a single pass when {#save} is called. By
  # design a write-only workbook can only be saved once: {#save} calls
  # {#close} on success, and any subsequent call raises
  # {Rbxl::WorkbookAlreadySavedError}.
  #
  #   book  = Rbxl.new(write_only: true)
  #   sheet = book.add_sheet("Report")
  #   sheet.append(["id", "name"])
  #   sheet.append([1, "alice"])
  #   book.save("report.xlsx")
  #
  # Cell styles can be registered up-front via {#add_style}; the returned
  # +style_id+ is then attached to {Rbxl::WriteOnlyCell} instances. Workbooks
  # that never call +add_style+ emit the same minimal +xl/styles.xml+ as
  # earlier releases, byte-for-byte.
  class WriteOnlyWorkbook
    # @return [Array<Rbxl::WriteOnlyWorksheet>] worksheets in insertion order
    attr_reader :worksheets

    # Creates an empty write-only workbook with no worksheets.
    def initialize
      @worksheets = []
      @closed = false
      @saved = false
      @style_registry = StyleRegistry.new
    end

    # Registers a cell style and returns its +style_id+ for use on
    # {Rbxl::WriteOnlyCell} instances.
    #
    #   header = book.add_style(
    #     font: { name: "Yu Gothic UI", size: 11, bold: true, color: "333333" },
    #     fill: { color: "EEEEEE" },
    #   )
    #   sheet << [Rbxl::WriteOnlyCell.new("Path", style_id: header)]
    #
    # Calls are not deduplicated by content: two equivalent registrations
    # return distinct +style_id+ values. Callers wanting reuse should hold
    # on to the integer themselves. Sub-tables (+fonts+, +fills+, +borders+,
    # +numFmts+) are deduplicated internally to keep +styles.xml+ compact.
    #
    # @param font [Hash, nil] +name+, +size+, +bold+, +italic+, +underline+, +color+
    # @param fill [Hash, nil] +color+ (RRGGBB hex) for solid fills, or +pattern_type+
    # @param border [Hash, nil] +all+ for four-sided shorthand, or +top+/+right+/+bottom+/+left+
    # @param alignment [Hash, nil] +horizontal+, +vertical+, +wrap_text+, +text_rotation+
    # @param number_format [String, nil] Excel format code
    # @return [Integer] 1-based style id
    def add_style(font: nil, fill: nil, border: nil, alignment: nil, number_format: nil)
      ensure_writable!

      @style_registry.register(
        Style.new(font: font, fill: fill, border: border, alignment: alignment, number_format: number_format)
      )
    end

    # Creates and returns a new worksheet appended to this workbook.
    #
    # @param name [String] visible sheet name
    # @return [Rbxl::WriteOnlyWorksheet]
    # @raise [Rbxl::ClosedWorkbookError] if the workbook has been closed
    # @raise [Rbxl::WorkbookAlreadySavedError] if {#save} has already succeeded
    def add_sheet(name)
      ensure_writable!

      sheet = WriteOnlyWorksheet.new(name: name)
      @worksheets << sheet
      sheet
    end

    # Serializes the workbook to an <tt>.xlsx</tt> file at +path+.
    #
    # On success the workbook is closed automatically; the method returns
    # the path that was written, suitable for chaining.
    #
    # @param path [String, #to_path] destination filesystem path
    # @return [String] the saved path
    # @raise [Rbxl::Error] if no worksheets have been added
    # @raise [Rbxl::ClosedWorkbookError] if the workbook is already closed
    # @raise [Rbxl::WorkbookAlreadySavedError] if {#save} was already called
    def save(path)
      ensure_writable!
      raise Error, "at least one worksheet is required" if worksheets.empty?

      previous_zip64 = Zip.write_zip64_support
      begin
        Zip.write_zip64_support = false

        Zip::OutputStream.open(path) do |zip|
          write_entry(zip, "[Content_Types].xml", content_types_xml)
          write_entry(zip, "_rels/.rels", root_rels_xml)
          write_entry(zip, "xl/workbook.xml", workbook_xml)
          write_entry(zip, "xl/_rels/workbook.xml.rels", workbook_rels_xml)
          write_entry(zip, "xl/styles.xml", styles_xml)

          worksheets.each_with_index do |sheet, index|
            sheet_index = index + 1
            write_entry(zip, "xl/worksheets/sheet#{sheet_index}.xml", sheet.to_xml)

            rels = sheet.external_hyperlink_relationships
            unless rels.empty?
              write_entry(zip, "xl/worksheets/_rels/sheet#{sheet_index}.xml.rels", sheet_rels_xml(rels))
            end
          end
        end
      ensure
        Zip.write_zip64_support = previous_zip64
      end

      @saved = true
      close
      path
    end

    # Marks the workbook as closed. Further mutating operations raise
    # {Rbxl::ClosedWorkbookError}. This is called automatically by a
    # successful {#save}.
    #
    # @return [Boolean] the new closed state (always +true+)
    def close
      @closed = true
    end

    # @return [Boolean] whether the workbook has been closed
    def closed?
      @closed
    end

    private

    def ensure_writable!
      raise WorkbookAlreadySavedError, "write-only workbook can only be saved once by design; call Rbxl.new to build another workbook" if @saved
      raise ClosedWorkbookError, "workbook has been closed" if closed?
    end

    def write_entry(zip, name, content)
      zip.put_next_entry(name)
      zip.write(content)
    end

    def content_types_xml
      worksheet_overrides = worksheets.each_index.map do |index|
        %(<Override PartName="/xl/worksheets/sheet#{index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)
      end.join

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          #{worksheet_overrides}
        </Types>
      XML
    end

    def root_rels_xml
      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      XML
    end

    def workbook_xml
      sheet_nodes = worksheets.each_with_index.map do |sheet, index|
        %(<sheet name="#{escape(sheet.name)}" sheetId="#{index + 1}" r:id="rId#{index + 1}"/>)
      end.join

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>#{sheet_nodes}</sheets>
        </workbook>
      XML
    end

    def workbook_rels_xml
      relationship_nodes = worksheets.each_with_index.map do |_, index|
        %(<Relationship Id="rId#{index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{index + 1}.xml"/>)
      end
      relationship_nodes << %(<Relationship Id="rId#{worksheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>)

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          #{relationship_nodes.join}
        </Relationships>
      XML
    end

    def styles_xml
      return @style_registry.to_xml unless @style_registry.empty?

      <<~XML.chomp
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fonts count="1">
            <font><sz val="11"/><name val="Calibri"/></font>
          </fonts>
          <fills count="1">
            <fill><patternFill patternType="none"/></fill>
          </fills>
          <borders count="1">
            <border><left/><right/><top/><bottom/><diagonal/></border>
          </borders>
          <cellStyleXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          </cellStyleXfs>
          <cellXfs count="1">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
          </cellXfs>
          <cellStyles count="1">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
          </cellStyles>
        </styleSheet>
      XML
    end

    def escape(value)
      CGI.escapeHTML(value.to_s)
    end

    def sheet_rels_xml(rels)
      buf = +'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      buf << "\n"
      buf << '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      rels.each do |r|
        buf << %(<Relationship Id="#{r[:rid]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="#{escape(r[:target])}" TargetMode="External"/>)
      end
      buf << '</Relationships>'
      buf
    end
  end
end
