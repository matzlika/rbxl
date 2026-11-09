# Changelog

All notable changes to this project are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `Rbxl::WriteOnlyWorkbook#add_style` registers a workbook-level cell style
  (font, fill, border, alignment, number_format) and returns a `style_id`
  that can be attached to `Rbxl::WriteOnlyCell` instances. Workbooks that
  never call `add_style` emit the same minimal `xl/styles.xml` as before,
  byte-for-byte.
- `Rbxl::WriteOnlyWorksheet#column_width` accepts an integer index or a
  Range plus a width; consecutive same-width columns merge into one
  `<col>` element on save.
- `Rbxl::WriteOnlyWorksheet#row_height` accepts an integer index or a
  Range plus a height. Native fast path supports row heights via an
  optional second argument to `Rbxl::Native.generate_sheet`, so heights
  ride the C path instead of forcing a Ruby fallback.
- `Rbxl::WriteOnlyWorksheet#freeze!` emits a `<sheetView>` `<pane>` with
  `topLeftCell`, `activePane`, and `state="frozen"` derived from `rows:`
  and `cols:` keyword arguments.
- `Rbxl::WriteOnlyWorksheet#merge` records merge ranges (either as
  `"A1:C1"` or as `"A1", "C1"`) and emits them in a `<mergeCells>` block
  after `</sheetData>`.
- `Rbxl::WriteOnlyCell` now accepts a `hyperlink:` keyword. Strings
  starting with `#` are emitted as internal sheet anchors via the
  `<hyperlink location="..."/>` form; everything else becomes an
  external URL with a per-sheet
  `xl/worksheets/_rels/sheetN.xml.rels` entry referenced through
  `r:id`. The `xmlns:r` namespace declaration is added to `<worksheet>`
  only when external links exist.
- Edit mode (`Rbxl.open(path, edit: true)`) grows the pieces a template
  fill-in needs:
  - `EditableCell#value=` accepts `Date`, `Time`, and `DateTime`. Values
    are written as Excel serials honoring the workbook's 1900/1904 date
    system. A cell that already has a date format keeps it; otherwise a
    copy of the cell's current style with the built-in short-date (14) or
    date-time (22) format is derived, so template fonts and borders stay.
  - `EditableCell#style(font:, fill:, border:, alignment:, number_format:)`
    changes part of a cell's current style — font and alignment
    attributes and border sides merge, fill and number format replace.
    `EditableWorkbook#add_style` registers a complete style (same keywords
    as write-only mode) for `EditableCell#style_id=`; `#style_id` reads
    the current index. New entries are only appended to `xl/styles.xml`,
    so existing style indices never change, and the part is rewritten
    only when a style was added.
  - `EditableWorksheet#write_rows(anchor, rows)` writes a 2D array into
    consecutive cells. Given a range such as `"B10:D19"`, data that does
    not fit raises before any cell is touched.
  - `EditableWorkbook#defined_name(name, sheet: nil)` returns the cell a
    single-cell defined name points to; `#defined_names` lists them.
    Raises the new `Rbxl::DefinedNameError` for unknown, ambiguous, or
    range names.
  - `EditableWorkbook#add_sheet(name)` appends an empty worksheet (part,
    workbook relationship, and content-type override).
  - `EditableWorkbook#save(path = nil, recalc: true)`: when cells were
    edited, `<calcPr fullCalcOnLoad="1"/>` is set so Excel and
    LibreOffice recompute stale template formulas on open. If an edit
    replaced a formula, `xl/calcChain.xml` is dropped so Excel rebuilds
    it.

### Changed

- Style colors (`font:` / `fill:` / `border:` `color`) are written as
  8-digit ARGB (`FFRRGGBB`), the form ECMA-376 specifies. A leading `#`
  is accepted; 8-digit input passes through.
- The `rubyzip` dependency now allows 3.x (`>= 2.3, < 4`). The suite
  passes on rubyzip 2.4.1 and 3.7.0.
- Edit mode no longer raises `Rbxl::EditableCellTypeError` for
  `Date`/`Time`/`DateTime` (see above). Editing a cell now also touches
  `xl/workbook.xml` (the recalculation flag) unless `recalc: false`.

### Fixed

- Edit mode: saving to another path, editing more, and saving again no
  longer drops the first round of edits.
- Edit mode: after an in-place `save`, opening a sheet that had not been
  loaded yet no longer reads from the replaced file's stale ZIP
  directory (it raised `WorksheetFormatError`).
- Edit mode: rows and cells inserted into a part that binds SpreadsheetML
  under a prefix (e.g. `x:` from the .NET OpenXML SDK) now get that
  namespace instead of none.
- Edit mode: re-serialized parts keep their original layout; libxml2 no
  longer re-indents them.

## [1.5.0] - 2026-10-31

### Changed

- The `rubyzip` dependency is widened from `~> 2.3` to `>= 2.4, < 4`, so
  rbxl now installs alongside rubyzip 3.x. The floor moves to 2.4, the
  first release accepting the `create:` keyword that 3.x requires. No
  library code changes were needed; the test suite passes against both
  rubyzip 2.4 and 3.x.

## [1.4.0] - 2026-05-01

### Added

- `Rbxl.open(path, edit: true)` opens a new `Rbxl::EditableWorkbook` for
  surgical read-modify-save passes against an existing `.xlsx`. The design
  promise is borrowed from rbpptx: parts you don't touch round-trip
  byte-for-byte (`copy_raw_entry` straight from the source ZIP), and inside
  a worksheet you do edit only the targeted `<c>` element is rewritten —
  sibling cells, row attributes, `<mergeCells>`, `<conditionalFormatting>`,
  `<dataValidations>`, comments, drawings, charts, pivot caches, and any
  unknown OOXML extensions stay in place. The cell's `s` (style index)
  attribute survives an overwrite so template number formats, fonts, and
  fills carry through to the new value. Cells written by this mode become
  inline strings (`t="inlineStr"`); `xl/sharedStrings.xml` is never
  mutated, so the output is deterministic without a second SST pass.
  Touched sheets are parsed as full Nokogiri DOMs, so this is the right
  tool for template fill-ins — not for rewriting the data area of a large
  worksheet (use `Rbxl.new` for that). `Rbxl::EditableCell#value=` accepts
  `nil`, `String`, `Integer`, `Float`, and `true`/`false`; `Date`/`Time`
  raise `Rbxl::EditableCellTypeError` so 1.4.0 doesn't ship a half-baked
  numFmt write. New cells (and their enclosing rows) are inserted in
  column- and row-sorted positions.
- `Rbxl::EditableWorkbook#save` accepts no path to save in place,
  rewriting the original via temp file plus atomic rename so a crash
  mid-save never produces a half-written workbook.

### Changed

- Shared-strings parsing factored into `Rbxl::SharedStringsLoader` so the
  read-only and editable workbooks share a single SST decoder rather than
  carrying parallel copies. Behavior and limits (`Rbxl.max_shared_strings`,
  `Rbxl.max_shared_string_bytes`) are unchanged.

## [1.3.0] - 2026-04-27

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
