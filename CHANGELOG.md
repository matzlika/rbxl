# Changelog

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
