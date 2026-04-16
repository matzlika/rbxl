require "mkmf"

# The extension is intentionally built against Nokogiri's vendored libxml2.
# We only borrow Nokogiri's headers at build time and rely on Nokogiri's
# extension to export the libxml2 symbols at runtime. Linking against the
# system libxml2 here would reintroduce mixed-version warnings and can lead
# to process instability.

begin
  require "nokogiri"
rescue LoadError
  warn "rbxl_native: nokogiri is required to build the C extension"
  File.write("Makefile", "all install clean:\n\t@:\n")
  exit 0
end

nokogiri_cppflags = Array(Nokogiri::VERSION_INFO.dig("nokogiri", "cppflags"))
nokogiri_ldflags = Array(Nokogiri::VERSION_INFO.dig("nokogiri", "ldflags"))

$CPPFLAGS = [*nokogiri_cppflags, $CPPFLAGS].reject(&:empty?).join(" ")
$LDFLAGS = [*nokogiri_ldflags, $LDFLAGS].reject(&:empty?).join(" ")

unless have_header("libxml/parser.h")
  warn "rbxl_native: failed to find Nokogiri libxml2 headers"
  File.write("Makefile", "all install clean:\n\t@:\n")
  exit 0
end

# macOS refuses unresolved references in shared objects unless explicitly told
# to leave them for runtime lookup in already-loaded extensions like Nokogiri.
if RUBY_PLATFORM.include?("darwin")
  append_ldflags("-Wl,-undefined,dynamic_lookup")
end

# Hardening flags
$CFLAGS << " -Wall -Wextra -Werror=format-security"
$CFLAGS << " -D_FORTIFY_SOURCE=2" unless $CFLAGS.include?("_FORTIFY_SOURCE")
$CFLAGS << " -fstack-protector-strong" if try_cflags("-fstack-protector-strong")

create_makefile("rbxl_native/rbxl_native")
