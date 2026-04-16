require "cgi"
require "date"
require "nokogiri"
require "stringio"
require "zip"

require_relative "rbxl/cell"
require_relative "rbxl/empty_cell"
require_relative "rbxl/errors"
require_relative "rbxl/read_only_cell"
require_relative "rbxl/read_only_workbook"
require_relative "rbxl/read_only_worksheet"
require_relative "rbxl/row"
require_relative "rbxl/version"
require_relative "rbxl/write_only_cell"
require_relative "rbxl/write_only_workbook"
require_relative "rbxl/write_only_worksheet"

# Minimal streaming XLSX reader/writer inspired by +openpyxl+.
#
# Rbxl exposes two explicit, non-overlapping modes:
#
# * {Rbxl.open} returns a {Rbxl::ReadOnlyWorkbook} for streaming reads
# * {Rbxl.new} returns a {Rbxl::WriteOnlyWorkbook} for one-shot writes
#
# The API is intentionally narrow so that memory usage stays predictable
# for large workbooks. Neither mode materializes the full workbook in
# memory; reads pull rows from the underlying XML one at a time, and writes
# accumulate only the rows added before {Rbxl::WriteOnlyWorkbook#save}.
#
# == Reading
#
#   require "rbxl"
#
#   book = Rbxl.open("report.xlsx", read_only: true)
#   sheet = book.sheet("Report")
#   sheet.each_row(values_only: true) { |values| p values }
#   book.close
#
# == Writing
#
#   require "rbxl"
#
#   book  = Rbxl.new(write_only: true)
#   sheet = book.add_sheet("Report")
#   sheet << ["id", "name", "score"]
#   sheet << [1, "alice", 100]
#   book.save("report.xlsx")
#
# == Native extension
#
# Requiring <tt>"rbxl/native"</tt> after <tt>"rbxl"</tt> swaps the hot
# worksheet XML paths for a libxml2-backed C implementation with the same
# observable behavior. See the README for build requirements.
module Rbxl
  # Maximum number of shared strings accepted from a workbook's
  # +xl/sharedStrings.xml+ entry. Defaults to 10 million, which comfortably
  # covers real-world enterprise workbooks while rejecting files crafted to
  # exhaust memory before any row is read. Set to +nil+ to disable.
  @max_shared_strings = 10_000_000

  # Maximum total byte size of the shared strings table once decoded.
  # Defaults to 512 MiB. Applied both to the ZIP entry's declared
  # uncompressed size (cheap early rejection of zip bombs) and to the
  # running sum while parsing. Set to +nil+ to disable.
  @max_shared_string_bytes = 512 * 1024 * 1024

  class << self
    # @return [Integer, nil] configured shared-strings count cap
    attr_accessor :max_shared_strings

    # @return [Integer, nil] configured shared-strings byte cap
    attr_accessor :max_shared_string_bytes

    # Opens an existing workbook in read-only streaming mode.
    #
    # The +read_only+ keyword is required and must be +true+. It exists to
    # mark the intent explicitly and to leave room for a future read/write
    # mode without changing the default behavior of {.open}.
    #
    # @param path [String, #to_path] filesystem path to an <tt>.xlsx</tt> file
    # @param read_only [Boolean] must be +true+ for the current API
    # @return [Rbxl::ReadOnlyWorkbook]
    # @raise [ArgumentError] if +read_only+ is not +true+
    def open(path, read_only: false)
      raise ArgumentError, "read_only: true is required for this MVP" unless read_only

      ReadOnlyWorkbook.open(path)
    end

    # Creates a new workbook in write-only mode.
    #
    # The +write_only+ keyword is required and must be +true+ to make the
    # save-once, append-only contract obvious at the call site.
    #
    # @param write_only [Boolean] must be +true+ for the current API
    # @return [Rbxl::WriteOnlyWorkbook]
    # @raise [ArgumentError] if +write_only+ is not +true+
    def new(write_only: false)
      raise ArgumentError, "write_only: true is required for this MVP" unless write_only

      WriteOnlyWorkbook.new
    end
  end
end
