module Rbxl
  # @api private
  #
  # Date/time number-format detection and Excel serial conversion, shared by
  # the read-only date conversion path and edit-mode date writes.
  module DateFormat
    # Built-in numFmtId values that Excel resolves to date/time formats.
    # Ids outside this set are dates only when the workbook provides a
    # matching custom +<numFmt>+ entry whose format code contains date
    # tokens. See ECMA-376 part 1 §18.8.30.
    BUILTIN_IDS = Set.new([14, 15, 16, 17, 18, 19, 20, 21, 22,
                           27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
                           45, 46, 47, 50, 51, 52, 53, 54, 55, 56,
                           57, 58]).freeze

    # Built-in short date (+m/d/yyyy+, rendered per the viewer's locale).
    BUILTIN_DATE_ID = 14

    # Built-in date + time (+m/d/yyyy h:mm+, rendered per the viewer's locale).
    BUILTIN_DATETIME_ID = 22

    EPOCH_1900 = Date.new(1899, 12, 31)
    EPOCH_1904 = Date.new(1904, 1, 1)
    private_constant :EPOCH_1900, :EPOCH_1904

    module_function

    # Quoted literals, bracketed directives (e.g. [Red], [$-409]), and
    # backslash-escaped characters never introduce date tokens, so strip
    # them before looking for +y/m/d/h/s+.
    def date_format_code?(code)
      stripped = code.dup
      stripped.gsub!(/\[[^\]]*\]/, "")
      stripped.gsub!(/"[^"]*"/, "")
      stripped.gsub!(/\\./, "")
      stripped.match?(/[ymdhs]/i)
    end

    # Converts a Ruby date/time to an Excel serial number. Time zones are
    # ignored — Excel has no notion of one, so the wall-clock fields of the
    # value are written as-is.
    #
    # The 1900 system reproduces Excel's phantom 1900-02-29 (serial 60), so
    # dates on or after 1900-03-01 get serials matching what Excel shows.
    #
    # @param value [Date, DateTime, Time]
    # @param date_1904 [Boolean] whether the workbook uses the 1904 system
    # @return [Integer, Float] whole serial for midnight, fractional otherwise
    # @raise [ArgumentError] for dates Excel cannot represent
    def to_serial(value, date_1904: false)
      date, seconds =
        case value
        when DateTime
          [value.to_date, (value.hour * 3600) + (value.min * 60) + value.sec + value.sec_fraction]
        when Time
          [value.to_date, (value.hour * 3600) + (value.min * 60) + value.sec + value.subsec]
        when Date
          [value, 0]
        else
          raise ArgumentError, "expected Date, DateTime, or Time, got #{value.class}"
        end

      days =
        if date_1904
          (date - EPOCH_1904).to_i
        else
          d = (date - EPOCH_1900).to_i
          d >= 60 ? d + 1 : d
        end

      min_days = date_1904 ? 0 : 1
      if days < min_days
        raise ArgumentError,
              "#{value.inspect} is before the workbook's date epoch (#{date_1904 ? '1904-01-01' : '1900-01-01'})"
      end

      return days if seconds.zero?

      days + (seconds.to_r / 86_400).to_f
    end
  end
end
