module Rbxl
  # Base class for all errors raised by Rbxl. Rescue this class to catch any
  # library-specific failure without catching unrelated +StandardError+
  # subclasses from the caller's code.
  class Error < StandardError; end

  # Raised by {Rbxl::ReadOnlyWorkbook#sheet} when the requested sheet name
  # is not present in the workbook.
  class SheetNotFoundError < Error; end

  # Raised when an operation is attempted against a workbook whose
  # underlying resources have already been released via +close+.
  class ClosedWorkbookError < Error; end

  # Raised by {Rbxl::WriteOnlyWorkbook#save} when the workbook has already
  # been persisted once. Write-only workbooks are save-once by design.
  class WorkbookAlreadySavedError < Error; end

  # Raised by {Rbxl::ReadOnlyWorksheet#calculate_dimension} when the sheet
  # lacks a stored +<dimension>+ element and the caller has not opted into
  # scanning the worksheet with <tt>force: true</tt>.
  class UnsizedWorksheetError < Error; end

  # Raised when the shared strings table in an opened workbook exceeds the
  # configured count or byte limits (see {Rbxl.max_shared_strings} and
  # {Rbxl.max_shared_string_bytes}). Guards against malicious or malformed
  # +.xlsx+ files that would otherwise exhaust memory before the first row
  # is read.
  class SharedStringsTooLargeError < Error; end

  # Raised when a worksheet's XML payload exceeds {Rbxl.max_worksheet_bytes}
  # while iterating in +streaming: true+ mode. Applies to the uncompressed
  # bytes consumed from the ZIP entry, so high-compression zip-bomb style
  # worksheets are stopped mid-inflate rather than after the fact.
  class WorksheetTooLargeError < Error; end
end
