require "minitest/autorun"
require "tmpdir"
require "zip"
require_relative "../lib/rbxl"
require_relative "../lib/rbxl/native"

class SecurityTest < Minitest::Test
  NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'

  def parse(xml)
    rows = []
    Rbxl::Native.parse_sheet(xml, []) { |row| rows << row.dup }
    rows
  end

  def with_shared_strings_xlsx(shared_strings_xml)
    Dir.mktmpdir do |dir|
      path = File.join(dir, "attack.xlsx")
      Zip::File.open(path, Zip::File::CREATE) do |zip|
        zip.get_output_stream("xl/sharedStrings.xml") { |f| f.write(shared_strings_xml) }
      end
      yield path
    end
  end

  def with_config(**overrides)
    saved = overrides.keys.each_with_object({}) { |k, h| h[k] = Rbxl.send(k) }
    overrides.each { |k, v| Rbxl.send("#{k}=", v) }
    yield
  ensure
    saved.each { |k, v| Rbxl.send("#{k}=", v) }
  end

  def test_user_defined_entity_is_not_substituted
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <!DOCTYPE worksheet [<!ENTITY pwn "LEAKED">]>
      <worksheet #{NS}>
        <sheetData>
          <row r="1"><c r="A1" t="inlineStr"><is><t>&pwn;</t></is></c></row>
        </sheetData>
      </worksheet>
    XML

    rows = parse(xml)
    flat = rows.flatten.compact.map(&:to_s).join
    refute_includes flat, "LEAKED",
      "user-defined entity must not be expanded (XXE defense)"
  end

  def test_external_system_entity_is_not_loaded
    # Reference a local file; parser must not read it.
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <!DOCTYPE worksheet [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
      <worksheet #{NS}>
        <sheetData>
          <row r="1"><c r="A1" t="inlineStr"><is><t>&xxe;</t></is></c></row>
        </sheetData>
      </worksheet>
    XML

    rows = parse(xml)
    flat = rows.flatten.compact.map(&:to_s).join
    refute_includes flat, "root:", "external entity must not be resolved"
    refute_includes flat, "/bin/", "external entity must not be resolved"
  end

  def test_billion_laughs_does_not_blow_up
    # Classic exponential-expansion payload. With NOENT disabled,
    # entities are never substituted so amplification cannot trigger.
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <!DOCTYPE worksheet [
        <!ENTITY a "aaaaaaaaaa">
        <!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">
        <!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">
        <!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">
        <!ENTITY e "&d;&d;&d;&d;&d;&d;&d;&d;&d;&d;">
      ]>
      <worksheet #{NS}>
        <sheetData>
          <row r="1"><c r="A1" t="inlineStr"><is><t>&e;</t></is></c></row>
        </sheetData>
      </worksheet>
    XML

    started = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    rows = parse(xml)
    elapsed = Process.clock_gettime(Process::CLOCK_MONOTONIC) - started

    assert elapsed < 1.0, "parse should finish quickly, took #{elapsed}s"
    flat = rows.flatten.compact.map(&:to_s).join
    assert flat.bytesize < 1_000, "entities must not be expanded, got #{flat.bytesize} bytes"
  end

  def test_deep_nesting_is_rejected
    # libxml2's default nesting limit (256) should engage without XML_PARSE_HUGE.
    depth = 400
    body = ("<g>" * depth) + ("</g>" * depth)
    xml = <<~XML
      <?xml version="1.0" encoding="UTF-8"?>
      <worksheet #{NS}>
        <sheetData>
          <row r="1"><c r="A1" t="inlineStr"><is><t>#{body}</t></is></c></row>
        </sheetData>
      </worksheet>
    XML

    assert_raises(RuntimeError) { parse(xml) }
  end

  def test_shared_strings_count_cap
    # 1_000 entries; cap at 10 should trip before parsing completes.
    entries = Array.new(1_000) { |i| "<si><t>s#{i}</t></si>" }.join
    xml = %(<?xml version="1.0" encoding="UTF-8"?>
<sst #{NS} count="1000" uniqueCount="1000">#{entries}</sst>)

    with_shared_strings_xlsx(xml) do |path|
      with_config(max_shared_strings: 10) do
        err = assert_raises(Rbxl::SharedStringsTooLargeError) do
          Rbxl.open(path, read_only: true)
        end
        assert_match(/count exceeds limit 10/, err.message)
      end
    end
  end

  def test_shared_strings_bytes_cap
    chunk = "A" * 1024
    entries = Array.new(100) { "<si><t>#{chunk}</t></si>" }.join
    xml = %(<?xml version="1.0" encoding="UTF-8"?>
<sst #{NS} count="100" uniqueCount="100">#{entries}</sst>)

    with_shared_strings_xlsx(xml) do |path|
      with_config(max_shared_string_bytes: 8 * 1024) do
        assert_raises(Rbxl::SharedStringsTooLargeError) do
          Rbxl.open(path, read_only: true)
        end
      end
    end
  end

  def test_shared_strings_zip_bomb_rejected_before_decompression
    # Highly compressible payload (millions of 'A's). Compressed size is tiny
    # but the entry's declared uncompressed size should trigger an early
    # rejection before we decompress anything.
    big = "A" * (5 * 1024 * 1024) # 5 MB
    xml = %(<?xml version="1.0" encoding="UTF-8"?>
<sst #{NS} count="1" uniqueCount="1"><si><t>#{big}</t></si></sst>)

    with_shared_strings_xlsx(xml) do |path|
      # Sanity: the declared uncompressed size is large, the compressed is small.
      Zip::File.open(path) do |zip|
        entry = zip.find_entry("xl/sharedStrings.xml")
        assert entry.size >= 5 * 1024 * 1024, "uncompressed size should be large"
        assert entry.compressed_size < 64 * 1024, "compressed size should be tiny"
      end

      with_config(max_shared_string_bytes: 1 * 1024 * 1024) do
        err = assert_raises(Rbxl::SharedStringsTooLargeError) do
          Rbxl.open(path, read_only: true)
        end
        assert_match(/uncompressed size/, err.message)
      end
    end
  end

  def test_streaming_worksheet_byte_cap_rejects_oversized_sheet
    Dir.mktmpdir do |dir|
      path = File.join(dir, "big.xlsx")
      book = Rbxl.new(write_only: true)
      sheet = book.add_sheet("S")
      1_000.times { |i| sheet.append([i, "row-#{i}", i * 1.5]) }
      book.save(path)

      with_config(max_worksheet_bytes: 1024) do
        wb = Rbxl.open(path, read_only: true, streaming: true)
        err = assert_raises(Rbxl::WorksheetTooLargeError) do
          wb.sheet("S").rows(values_only: true).count
        end
        assert_match(/worksheet bytes exceed limit/, err.message)
        wb.close
      end
    end
  end
end
