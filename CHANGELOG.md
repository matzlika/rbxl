# Changelog

## 1.1.0

- `Rbxl.open` and `Rbxl.new` now default `read_only: true` and `write_only: true` respectively, so the call site no longer needs the boilerplate. Explicitly passing `false` raises `NotImplementedError`.
- Add `date_conversion: true` to `Rbxl.open`: numeric cells whose style points at a date/time `numFmt` (built-in ids 14–22, 27–36, 45–47, 50–58, or a custom format code containing date tokens) are returned as `Date` or `Time`. Off by default — no change in output shape or throughput when the flag is absent.
- Fix Ruby reader path so self-closing `<row/>` and `<c/>` elements are iterated instead of silently dropped, and never yield `nil` for a row.

## 1.0.2

- Add `streaming: true` to `Rbxl.open` to feed worksheet XML to the native reader in 64 KiB chunks instead of buffering the full worksheet first.
- Add `Rbxl.max_worksheet_bytes` and `Rbxl::WorksheetTooLargeError` so streaming reads can stop oversized worksheet XML entries mid-inflate.
- Expand RDoc coverage across the public API.
- Tighten RBS signatures to match the actual runtime types.
- Reword public docs and gem metadata to describe reads as row-by-row and writes as append-only, reserving "streaming" for the new opt-in native read path.

## 1.0.1

- Fix ZIP64 handling.
- Add Go and Rust benchmark comparisons.
- Align `rbxl/native` with Nokogiri's libxml2 to avoid mixed-library warnings at runtime.

## 1.0.0

- Initial 1.0 release.
