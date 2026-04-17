require "cgi"
require "date"
require "nokogiri"
require "set"
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

# Minimal, memory-friendly XLSX reader/writer inspired by +openpyxl+.
#
# Rbxl exposes two explicit, non-overlapping modes:
#
# * {Rbxl.open} returns a {Rbxl::ReadOnlyWorkbook} for row-by-row reads
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
#   book = Rbxl.open("report.xlsx")
#   sheet = book.sheet("Report")
#   sheet.each_row(values_only: true) { |values| p values }
#   book.close
#
# Pass <tt>date_conversion: true</tt> to return Date/Time objects for
# numeric cells that carry a date +numFmt+ style.
#
# == Writing
#
#   require "rbxl"
#
#   book  = Rbxl.new
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

  # Maximum uncompressed byte size accepted from a single worksheet's XML
  # entry while iterating in +streaming: true+ mode. Default is +nil+
  # (unbounded) because legitimate worksheets can be arbitrarily large.
  # Set a positive integer to bound peak inflation and stop high-compression
  # zip-bomb worksheets mid-inflate.
  @max_worksheet_bytes = nil

  class << self
    # @return [Integer, nil] configured shared-strings count cap
    attr_accessor :max_shared_strings

    # @return [Integer, nil] configured shared-strings byte cap
    attr_accessor :max_shared_string_bytes

    # @return [Integer, nil] per-worksheet streaming byte cap
    attr_accessor :max_worksheet_bytes

    # Opens an existing workbook in read-only row-by-row mode.
    #
    # The +read_only+ keyword defaults to +true+ and exists to mark the
    # intent explicitly at the call site. Passing +read_only: false+ raises
    # {NotImplementedError}; a read/write mode is not available.
    #
    # With <tt>streaming: true</tt>, the native backend (when loaded) feeds
    # worksheet XML to the parser in chunks pulled from the ZIP input stream
    # instead of materializing the entire worksheet as one Ruby string. This
    # keeps peak memory roughly independent of worksheet size and lets
    # {Rbxl.max_worksheet_bytes} bound how much is inflated. Streaming mode
    # is the same API and output shape — only the inflation strategy
    # differs — and typically pays back a few percent of throughput on small
    # sheets in exchange for the flat memory profile.
    #
    # With <tt>date_conversion: true</tt>, numeric cells whose style points at
    # a date/time +numFmt+ (built-in ids 14–22, 27–36, 45–47, 50–58, or any
    # custom format code containing a date/time token) are returned as
    # +Date+, +Time+, or +DateTime+ instead of a raw serial +Float+. The flag
    # is off by default to preserve byte-for-byte behavior and skip the
    # styles.xml parse for workbooks that don't need it; enabling it
    # disables the native fast path and routes reads through the Ruby
    # worksheet parser.
    #
    # @param path [String, #to_path] filesystem path to an <tt>.xlsx</tt> file
    # @param read_only [Boolean] retained for call-site clarity; must be +true+
    # @param streaming [Boolean] feed worksheet XML to the native parser in
    #   chunks instead of fully inflating the entry in advance. Ignored when
    #   the native extension is not loaded.
    # @param date_conversion [Boolean] convert numeric cells backed by a
    #   date/time +numFmt+ to +Date+ / +Time+ / +DateTime+
    # @return [Rbxl::ReadOnlyWorkbook]
    # @raise [NotImplementedError] if +read_only+ is not +true+
    def open(path, read_only: true, streaming: false, date_conversion: false)
      raise NotImplementedError, "read/write mode is not supported; pass read_only: true" unless read_only

      ReadOnlyWorkbook.open(path, streaming: streaming, date_conversion: date_conversion)
    end

    # Creates a new workbook in write-only mode.
    #
    # The +write_only+ keyword defaults to +true+ and exists to mark the
    # save-once, append-only contract explicitly. Passing
    # +write_only: false+ raises {NotImplementedError}.
    #
    # @param write_only [Boolean] retained for call-site clarity; must be +true+
    # @return [Rbxl::WriteOnlyWorkbook]
    # @raise [NotImplementedError] if +write_only+ is not +true+
    def new(write_only: true)
      raise NotImplementedError, "read/write mode is not supported; pass write_only: true" unless write_only

      WriteOnlyWorkbook.new
    end
  end
end
