module Rbxl
  class Error < StandardError; end
  class SheetNotFoundError < Error; end
  class ClosedWorkbookError < Error; end
  class WorkbookAlreadySavedError < Error; end
  class UnsizedWorksheetError < Error; end
end
