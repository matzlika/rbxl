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
end
