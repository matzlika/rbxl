#!/usr/bin/env ruby

require "benchmark"
require "fileutils"
require "tmpdir"
require_relative "../lib/rbxl"

ROWS = Integer(ENV.fetch("RBXL_BENCH_ROWS", "10000"))
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

Dir.mktmpdir("rbxl-bench-") do |dir|
  path = File.join(dir, "bench.xlsx")
  header, body = build_dataset(rows: ROWS, cols: COLS)

  puts "rows=#{ROWS} cols=#{COLS}"

  Benchmark.bm(18) do |x|
    x.report("write_only save") do
      book = Rbxl.new(write_only: true)
      sheet = book.add_sheet("Bench")
      sheet << header
      body.each { |row| sheet << row }
      book.save(path)
    end

    x.report("read_only iterate") do
      book = Rbxl.open(path, read_only: true)
      count = 0
      book.sheet("Bench").each_row do |row|
        count += row.size
      end
      book.close
      raise "unexpected count=#{count}" if count <= 0
    end
  end

  puts "file=#{path}"
  puts "size_bytes=#{File.size(path)}"
  puts "rss_kb=#{rss_kb}"
end
