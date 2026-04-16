# rbxl

`openpyxl` inspired Ruby gem for large-ish `.xlsx` files.

Current scope is intentionally small:

- `write_only` workbook generation
- `read_only` row streaming
- `close()` for read-only workbooks
- minimal `openpyxl`-like API

Out of scope for this MVP:

- preserving arbitrary workbook structure on save
- rich style round-tripping
- formulas, images, charts, comments

## Usage

```ruby
require "rbxl"

book = Rbxl.new(write_only: true)
sheet = book.add_sheet("Report")
sheet.append(["id", "name", "score"])
sheet.append([1, "alice", 100])
sheet.append([2, "bob", 95.5])
book.save("report.xlsx")
```

```ruby
require "rbxl"

book = Rbxl.open("report.xlsx", read_only: true)
sheet = book.sheet("Report")

sheet.each_row do |row|
  p row.values
end

p sheet.calculate_dimension

book.close
```

`write_only` workbooks are save-once by design. This matches the optimized
mode tradeoff: low flexibility in exchange for simpler memory behavior.

## Design Notes

- Writer avoids a full workbook object graph and streams rows into sheet XML.
- Reader uses a pull parser for worksheet XML so it can iterate rows without building the full DOM.
- Strings written by the MVP use `inlineStr` to avoid shared string bookkeeping during generation.
- Reader supports both shared strings and inline strings.

## Development

```bash
bundle install
RBENV_VERSION=3.4.5 ruby -Ilib test/rbxl_test.rb
RBENV_VERSION=3.4.5 ruby benchmark/write_read.rb
RBENV_VERSION=3.4.5 ruby benchmark/compare.rb
```

## Benchmarks

The included benchmark is a simple smoke benchmark for the optimized modes.

```bash
RBENV_VERSION=3.4.5 RBXL_BENCH_ROWS=20000 RBXL_BENCH_COLS=12 ruby benchmark/write_read.rb
```

It reports write time, read iteration time, output file size, and current RSS.

For a lightweight comparison against other Ruby Excel gems installed on the
machine:

```bash
RBENV_VERSION=3.4.5 RBXL_BENCH_ROWS=5000 RBXL_BENCH_COLS=10 ruby benchmark/compare.rb
```

The comparison script uses these libraries when available:

- `rbxl` for write/read
- `caxlsx` for write
- `roo` for read streaming
- `rubyXL` for full workbook read
- `openpyxl` as a Python reference point when `openpyxl` or `uv` is available

The `openpyxl` numbers are reference-only. They are useful for directionally
comparing optimized mode behavior, but they are not a pure language-neutral
benchmark.

Current reference result on this machine with `RBXL_BENCH_ROWS=5000` and
`RBXL_BENCH_COLS=10`:

```text
benchmark                 real_s  rss_delta_kb    file_bytes
rbxl write                0.1303         12440        193193
rbxl read                 0.4415          1012             -
rbxl read values          0.3947          1236             -
caxlsx write              0.4943          4152        198420
roo read                  1.1326         17808             -
rubyXL read               1.9982        140736             -
openpyxl write            0.4291             0        194411
openpyxl read             0.3422             0             -
openpyxl read values      0.2626             0             -
```

Interpretation:

- `rbxl` write is currently faster than `caxlsx` and `openpyxl` in this benchmark.
- `rbxl` read is still slower than `openpyxl`, but faster than `roo` and `rubyXL`.
- `rbxl` read with `values_only` remains materially closer to `openpyxl` than the other Ruby readers here.
