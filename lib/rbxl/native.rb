begin
  require "rbxl_native/rbxl_native"
rescue LoadError
  # Try loading from ext/ build directory (development)
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
