#!/usr/bin/env ruby

require "optparse"
require_relative "../lib/rbxl"

options = {
  rows: 5,
  sheet: nil
}

OptionParser.new do |opts|
  opts.banner = "Usage: ruby script/parse_xlsx.rb PATH [--sheet NAME] [--rows N]"

  opts.on("--sheet NAME", "Read only the named sheet") do |value|
    options[:sheet] = value
  end

  opts.on("--rows N", Integer, "Preview row count per sheet (default: 5)") do |value|
    options[:rows] = value
  end
end.parse!

path = ARGV.shift
abort("xlsx path is required") unless path

workbook = Rbxl.open(path, read_only: true)
sheet_names = options[:sheet] ? [options[:sheet]] : workbook.sheet_names

puts "file: #{path}"
puts "sheets: #{workbook.sheet_names.join(', ')}"

sheet_names.each do |sheet_name|
  sheet = workbook.sheet(sheet_name)
  puts
  puts "[#{sheet_name}]"
  puts "dimension: #{sheet.calculate_dimension(force: true)}"

  sheet.rows(values_only: true).take(options[:rows]).each_with_index do |row, index|
    printable = row.map { |value| value.nil? ? nil : value.to_s }
    puts "row #{index + 1}: #{printable.inspect}"
  end
end

workbook.close
