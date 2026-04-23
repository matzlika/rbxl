# Changelog

All notable changes to this project are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `Rbxl.open` (and `Rbxl::ReadOnlyWorkbook.open`) now accept a block. The
  workbook is yielded and closed automatically when the block returns or
  raises, matching the `File.open` / `Zip::File.open` idiom. Previously the
  block was silently ignored.
- `Rbxl::UnsupportedFormatError` raised by `Rbxl.open` when the file is not
  a `.xlsx` container. Legacy `.xls` (BIFF/CFB) inputs are detected by the
  OLE compound-file magic and reported with a conversion hint, instead of
  surfacing an opaque `Zip::Error` from rubyzip five frames deep.
- `Rbxl::ReadOnlyWorkbook#sheet` now accepts an integer index into
  `sheet_names` (including negatives, so `sheet(-1)` returns the last
  sheet), for the common single-sheet case where `book.sheet(0)` reads
  cleaner than `book.sheet(book.sheet_names.first)`.
- `Rbxl::ReadOnlyWorkbook#sheets` iterator over worksheets in workbook
  order. Returns an `Enumerator` when called without a block, so
  `book.sheets.first` and `book.sheets.map(&:name)` compose naturally.
  Worksheet objects are constructed on demand — no eager parse of sibling
  sheets.

## [1.2.0] - 2026-04-23

### Changed

- `WorkbookAlreadySavedError` message now points at the save-once design and
  the next action (open a fresh `Rbxl.new` for another file) so callers who
  trip on the constraint don't have to read the source to understand why.
- Workbook- and worksheet-level parse failures raise `WorkbookFormatError` /
  `WorksheetFormatError` with the workbook path and the XML entry or sheet
  name in the message, replacing generic parser exceptions.

### Added

- Location-aware coverage around malformed workbook and worksheet XML so bad
  inputs surface the specific entry that failed rather than bubbling up an
  unlabelled `Nokogiri::XML::SyntaxError`.
- README sections covering the write-only model (append-only, save-once,
  no in-place edit), a "Reading recipes" walkthrough, and an explicit Out
  of scope entry for read-modify-save workflows.

### Fixed

- Honor Excel's `date1904` workbook setting when `date_conversion: true` is
  enabled, so Mac-originated workbooks map serial dates to the correct Ruby
  `Date` and `Time` values.

## [1.1.0] - 2026-04-21

### Added

- `date_conversion: true` option for `Rbxl.open`: numeric cells whose style
  points at a date/time `numFmt` (built-in ids 14–22, 27–36, 45–47, 50–58,
  or a custom format code containing date tokens) are returned as `Date`
  or `Time`. Off by default — no change in output shape or throughput when
  the flag is absent.

### Changed

- `Rbxl.open` and `Rbxl.new` now default `read_only: true` and
  `write_only: true` respectively, so the call site no longer needs the
  boilerplate. Explicitly passing `false` raises `NotImplementedError`.

### Fixed

- Ruby reader path now iterates self-closing `<row/>` and `<c/>` elements
  instead of silently dropping them, and never yields `nil` for a row.

## [1.0.2] - 2026-04-17

### Added

- `streaming: true` option for `Rbxl.open` feeds worksheet XML to the
  native reader in 64 KiB chunks instead of buffering the full worksheet
  first.
- `Rbxl.max_worksheet_bytes` configuration and `Rbxl::WorksheetTooLargeError`
  so streaming reads can stop oversized worksheet XML entries mid-inflate.

### Changed

- Expand RDoc coverage across the public API.
- Tighten RBS signatures to match the actual runtime types.
- Reword public docs and gem metadata to describe reads as row-by-row and
  writes as append-only, reserving "streaming" for the new opt-in native
  read path.

## [1.0.1] - 2026-04-16

### Added

- Go and Rust benchmark comparisons.

### Fixed

- ZIP64 handling.
- Align `rbxl/native` with Nokogiri's libxml2 to avoid mixed-library
  warnings at runtime.

## [1.0.0] - 2026-04-16

- Initial public release.
