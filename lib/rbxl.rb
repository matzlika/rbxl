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

module Rbxl
  class << self
    def open(path, read_only: false)
      raise ArgumentError, "read_only: true is required for this MVP" unless read_only

      ReadOnlyWorkbook.open(path)
    end

    def new(write_only: false)
      raise ArgumentError, "write_only: true is required for this MVP" unless write_only

      WriteOnlyWorkbook.new
    end
  end
end
