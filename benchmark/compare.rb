#!/usr/bin/env ruby

require "benchmark"
require "json"
require "tmpdir"
require_relative "../lib/rbxl"

ROWS = Integer(ENV.fetch("RBXL_BENCH_ROWS", "5000"))
COLS = Integer(ENV.fetch("RBXL_BENCH_COLS", "10"))

def rss_kb
  status = File.read("/proc/#{$$}/status")
  match = status.match(/^VmRSS:\s+(\d+)\s+kB$/)
  match ? match[1].to_i : 0
rescue Errno::ENOENT
  0
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

def benchmark(label)
  started_rss = rss_kb
  real = Benchmark.realtime { yield }
  {
    label: label,
    real: real,
    rss_delta_kb: rss_kb - started_rss
  }
end

def load_optional(name)
  require name
  true
rescue LoadError
  false
end

def write_with_rbxl(path, header, body)
  book = Rbxl.new(write_only: true)
  sheet = book.add_sheet("Bench")
  sheet.append(header)
  body.each { |row| sheet.append(row) }
  book.save(path)
end

def read_with_rbxl(path)
  book = Rbxl.open(path, read_only: true)
  count = 0
  book.sheet("Bench").rows.each { |row| count += row.size }
  book.close
  count
end

def read_values_with_rbxl(path)
  book = Rbxl.open(path, read_only: true)
  count = 0
  book.sheet("Bench").rows(values_only: true).each { |row| count += row.size }
  book.close
  count
end

def write_with_caxlsx(path, header, body)
  package = Axlsx::Package.new
  package.workbook.add_worksheet(name: "Bench") do |sheet|
    sheet.add_row(header)
    body.each { |row| sheet.add_row(row) }
  end
  package.serialize(path)
end

def read_with_roo(path)
  workbook = Roo::Excelx.new(path)
  count = 0
  workbook.each_row_streaming(sheet: "Bench") do |row|
    count += row.size
  end
  count
end

def read_with_rubyxl(path)
  workbook = RubyXL::Parser.parse(path)
  worksheet = workbook["Bench"]
  count = 0
  worksheet.each do |row|
    next unless row

    count += row.size
  end
  count
end

def shell_join(parts)
  parts.map do |part|
    if part.match?(/\A[a-zA-Z0-9_\/.\-:]+\z/)
      part
    else
      "'" + part.gsub("'", %q('"'"')) + "'"
    end
  end.join(" ")
end

def openpyxl_available?
  system("python3", "-c", "import openpyxl", out: File::NULL, err: File::NULL)
end

def run_openpyxl_helper
  helper = File.expand_path("openpyxl_compare.py", __dir__)
  command =
    if openpyxl_available?
      ["python3", helper]
    else
      uv = File.expand_path("~/.local/bin/uv")
      return [] unless File.exist?(uv)

      [uv, "run", "--with", "openpyxl", "python3", helper]
    end

  output = `#{shell_join(command)}`
  raise "openpyxl benchmark failed" unless $?.success?

  JSON.parse(output, symbolize_names: true)
end

Dir.mktmpdir("rbxl-compare-") do |dir|
  header, body = build_dataset(rows: ROWS, cols: COLS)
  results = []

  puts "rows=#{ROWS} cols=#{COLS}"

  rbxl_path = File.join(dir, "rbxl.xlsx")
  results << benchmark("rbxl write") { write_with_rbxl(rbxl_path, header, body) }.merge(size: File.size(rbxl_path))
  results << benchmark("rbxl read") { read_with_rbxl(rbxl_path) }
  results << benchmark("rbxl read values") { read_values_with_rbxl(rbxl_path) }

  if load_optional("caxlsx")
    caxlsx_path = File.join(dir, "caxlsx.xlsx")
    results << benchmark("caxlsx write") { write_with_caxlsx(caxlsx_path, header, body) }.merge(size: File.size(caxlsx_path))
  end

  if load_optional("roo")
    results << benchmark("roo read") { read_with_roo(rbxl_path) }
  end

  if load_optional("rubyXL")
    results << benchmark("rubyXL read") { read_with_rubyxl(rbxl_path) }
  end

  begin
    results.concat(run_openpyxl_helper)
  rescue StandardError => e
    warn "openpyxl benchmark skipped: #{e.message}"
  end

  label_width = results.map { |row| row[:label].length }.max
  puts format("%-#{label_width}s  %10s  %12s  %12s", "benchmark", "real_s", "rss_delta_kb", "file_bytes")
  results.each do |row|
    puts format(
      "%-#{label_width}s  %10.4f  %12d  %12s",
      row[:label],
      row[:real],
      row[:rss_delta_kb],
      row[:size] || "-"
    )
  end
end
