require "nokogiri"

# Opt-in loader for the libxml2-backed native extension.
#
# Requiring this file replaces the pure-Ruby worksheet XML parser and
# serializer with a C implementation that uses libxml2's SAX2 API directly.
# The public API exposed by {Rbxl} is unchanged; only the hot paths are
# swapped.
#
# The shared object is located in one of two places:
#
# 1. An installed gem layout (+rbxl_native/rbxl_native.so+ on the load path).
# 2. A development build tree under <tt>ext/rbxl_native/</tt>.
#
# If neither is available a +LoadError+ is raised with guidance on how to
# build the extension.
begin
  require "rbxl_native/rbxl_native"
rescue LoadError
  ext_path = File.expand_path("../../ext/rbxl_native", __dir__)
  so = Dir.glob(File.join(ext_path, "**", "rbxl_native.{so,bundle,dll}")).first
  if so
    require so
  else
    raise LoadError,
      "rbxl_native C extension not found. " \
      "Ensure libxml2 development headers are installed and run: " \
      "cd ext/rbxl_native && ruby extconf.rb && make"
  end
end
