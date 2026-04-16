#!/usr/bin/env ruby

require "benchmark"
require "tmpdir"
require "zip"
require_relative "../lib/rbxl"

ROWS = Integer(ENV.fetch("RBXL_BENCH_ROWS", "5000"))
COLS = Integer(ENV.fetch("RBXL_BENCH_COLS", "10"))
SPARSE_INTERVAL = Integer(ENV.fetch("RBXL_BENCH_SPARSE_INTERVAL", "3"))

def rss_kb
  status = File.read("/proc/#{$$}/status")
  match = status.match(/^VmRSS:\s+(\d+)\s+kB$/)
  match ? match[1].to_i : 0
rescue Errno::ENOENT
  0
end

def benchmark(label)
  started_rss = rss_kb
  result = nil
  real = Benchmark.realtime { result = yield }
  {
    label: label,
    real: real,
    rss_delta_kb: rss_kb - started_rss,
    result: result
  }
end

def build_dataset(rows:, cols:)
  header = Array.new(cols) { |i| "col_#{i + 1}" }
  body = Array.new(rows) do |row|
    Array.new(cols) do |col|
      case col % 4
      when 0 then row
      when 1 then "row-#{row}-col-#{col}"
      when 2 then (row + col).odd?
      else ((row * 100) + col) / 10.0
      end
    end
  end
  [header, body]
end

def build_sparse_dataset(rows:, cols:, interval:)
  header = Array.new(cols) { |i| "col_#{i + 1}" }
  body = Array.new(rows) do |row|
    Array.new(cols) do |col|
      next nil if ((row + col) % interval).positive?

      case col % 3
      when 0 then "s#{row}-#{col}"
      when 1 then row + col
      else true
      end
    end
  end
  [header, body]
end

def write_with_rbxl(path, header, body)
  book = Rbxl.new(write_only: true)
  sheet = book.add_sheet("Bench")
  sheet.append(header)
  body.each { |row| sheet.append(row) }
  book.save(path)
end

def write_with_shared_strings(path, header, body)
  rows = [header, *body]
  shared_strings = []
  shared_string_index = {}

  sheet_rows = rows.each_with_index.map do |row_values, row_index|
    cells = row_values.each_with_index.map do |value, col_index|
      reference = "#{column_name(col_index + 1)}#{row_index + 1}"
      serialize_shared_string_cell(reference, value, shared_strings, shared_string_index)
    end.join

    %(<row r="#{row_index + 1}">#{cells}</row>)
  end.join

  dimension_ref = rows.empty? ? "A1" : "A1:#{column_name(rows.map(&:length).max)}#{rows.length}"
  worksheet_xml = <<~XML.chomp
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dimension ref="#{dimension_ref}"/>
      <sheetData>#{sheet_rows}</sheetData>
    </worksheet>
  XML

  shared_strings_xml = <<~XML.chomp
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{shared_strings.length}" uniqueCount="#{shared_strings.length}">
      #{shared_strings.map { |text| %(<si><t>#{escape_xml(text)}</t></si>) }.join}
    </sst>
  XML

  Zip::OutputStream.open(path) do |zip|
    write_zip_entry(zip, "[Content_Types].xml", <<~XML.chomp)
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      </Types>
    XML
    write_zip_entry(zip, "_rels/.rels", <<~XML.chomp)
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>
    XML
    write_zip_entry(zip, "xl/workbook.xml", <<~XML.chomp)
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="Bench" sheetId="1" r:id="rId1"/></sheets>
      </workbook>
    XML
    write_zip_entry(zip, "xl/_rels/workbook.xml.rels", <<~XML.chomp)
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
      </Relationships>
    XML
    write_zip_entry(zip, "xl/styles.xml", <<~XML.chomp)
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
      </styleSheet>
    XML
    write_zip_entry(zip, "xl/sharedStrings.xml", shared_strings_xml)
    write_zip_entry(zip, "xl/worksheets/sheet1.xml", worksheet_xml)
  end
end

def serialize_shared_string_cell(reference, value, shared_strings, shared_string_index)
  case value
  when nil
    %(<c r="#{reference}"/>)
  when true
    %(<c r="#{reference}" t="b"><v>1</v></c>)
  when false
    %(<c r="#{reference}" t="b"><v>0</v></c>)
  when Numeric
    %(<c r="#{reference}"><v>#{value}</v></c>)
  else
    index = shared_string_index.fetch(value.to_s) do
      shared_strings << value.to_s
      shared_string_index[value.to_s] = shared_strings.length - 1
    end
    %(<c r="#{reference}" t="s"><v>#{index}</v></c>)
  end
end

def write_zip_entry(zip, name, content)
  zip.put_next_entry(name)
  zip.write(content)
end

def escape_xml(text)
  CGI.escapeHTML(text.to_s)
end

def column_name(index)
  name = +""
  current = index

  while current.positive?
    current -= 1
    name.prepend((65 + (current % 26)).chr)
    current /= 26
  end

  name
end

def read_count(path, **options)
  book = Rbxl.open(path, read_only: true)
  count = 0
  book.sheet("Bench").each_row(**options) do |row|
    count += row.size
  end
  book.close
  count
end

Dir.mktmpdir("rbxl-read-modes-") do |dir|
  dense_header, dense_body = build_dataset(rows: ROWS, cols: COLS)
  sparse_header, sparse_body = build_sparse_dataset(rows: ROWS, cols: COLS, interval: SPARSE_INTERVAL)

  dense_path = File.join(dir, "dense.xlsx")
  sparse_path = File.join(dir, "sparse.xlsx")
  dense_shared_path = File.join(dir, "dense-shared.xlsx")
  sparse_shared_path = File.join(dir, "sparse-shared.xlsx")
  write_with_rbxl(dense_path, dense_header, dense_body)
  write_with_rbxl(sparse_path, sparse_header, sparse_body)
  write_with_shared_strings(dense_shared_path, dense_header, dense_body)
  write_with_shared_strings(sparse_shared_path, sparse_header, sparse_body)

  results = []
  results << benchmark("dense values") { read_count(dense_path, values_only: true) }
  results << benchmark("dense cells") { read_count(dense_path) }
  results << benchmark("dense shared values") { read_count(dense_shared_path, values_only: true) }
  results << benchmark("dense shared cells") { read_count(dense_shared_path) }
  results << benchmark("sparse values") { read_count(sparse_path, values_only: true) }
  results << benchmark("sparse padded") { read_count(sparse_path, values_only: true, pad_cells: true) }
  results << benchmark("sparse merged") { read_count(sparse_path, values_only: true, pad_cells: true, expand_merged: true) }
  results << benchmark("sparse shared values") { read_count(sparse_shared_path, values_only: true) }
  results << benchmark("sparse shared padded") { read_count(sparse_shared_path, values_only: true, pad_cells: true) }

  label_width = results.map { |row| row[:label].length }.max
  puts "rows=#{ROWS} cols=#{COLS} sparse_interval=#{SPARSE_INTERVAL}"
  puts "dense_file_bytes=#{File.size(dense_path)}"
  puts "sparse_file_bytes=#{File.size(sparse_path)}"
  puts "dense_shared_file_bytes=#{File.size(dense_shared_path)}"
  puts "sparse_shared_file_bytes=#{File.size(sparse_shared_path)}"
  puts format("%-#{label_width}s  %10s  %12s  %12s", "benchmark", "real_s", "rss_delta_kb", "cell_count")
  results.each do |row|
    puts format(
      "%-#{label_width}s  %10.4f  %12d  %12d",
      row[:label],
      row[:real],
      row[:rss_delta_kb],
      row[:result]
    )
  end
end
