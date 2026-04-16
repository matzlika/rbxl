require "mkmf"

# Try to find libxml2 headers and library.
# Priority:
#   1. Nokogiri's bundled libxml2 (avoids version mismatch warnings)
#   2. System pkg-config
#   3. Common system paths
#
# If libxml2 is not available at all, skip compilation gracefully so
# that `gem install rbxl` never fails — the C extension is optional.

found = false

# 1. Try Nokogiri's bundled libxml2
begin
  nokogiri_spec = Gem::Specification.find_by_name("nokogiri")
  nokogiri_include = File.join(nokogiri_spec.full_gem_path, "ext", "nokogiri", "include", "libxml2")
  nokogiri_lib = File.join(nokogiri_spec.full_gem_path, "ext", "nokogiri")

  if File.directory?(nokogiri_include) && find_header("libxml/parser.h", nokogiri_include)
    # Link against Nokogiri's bundled libxml2
    nokogiri_so = Dir.glob(File.join(nokogiri_lib, "**", "nokogiri.{so,bundle}")).first
    if nokogiri_so
      so_dir = File.dirname(nokogiri_so)
      $LDFLAGS << " -L#{so_dir} -Wl,-rpath,#{so_dir}"
    end
    found = have_library("xml2") || true # headers found via Nokogiri, may link at runtime
  end
rescue Gem::MissingSpecError
  # Nokogiri not installed — fall through
end

# 2. System pkg-config
found ||= pkg_config("libxml-2.0")

# 3. Common system paths
found ||= (have_header("libxml/parser.h") && have_library("xml2"))
found ||= (find_header("libxml/parser.h", "/usr/include/libxml2") && have_library("xml2"))

unless found
  warn "rbxl_native: libxml2 not found — skipping C extension build"
  File.write("Makefile", "all install clean:\n\t@:\n")
  exit 0
end

# Hardening flags
$CFLAGS << " -Wall -Wextra -Werror=format-security"
$CFLAGS << " -D_FORTIFY_SOURCE=2" unless $CFLAGS.include?("_FORTIFY_SOURCE")
$CFLAGS << " -fstack-protector-strong" if try_cflags("-fstack-protector-strong")

create_makefile("rbxl_native/rbxl_native")
