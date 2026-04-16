require_relative "lib/rbxl/version"

Gem::Specification.new do |spec|
  spec.name = "rbxl"
  spec.version = Rbxl::VERSION
  spec.authors = ["Codex"]
  spec.summary = "Streaming xlsx reader/writer inspired by openpyxl"
  spec.description = "A small Ruby gem for read-only and write-only xlsx workflows."
  spec.homepage = "https://example.invalid/rbxl"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 3.1"

  spec.files = Dir["lib/**/*.rb"] + Dir["ext/**/*.{rb,c,h}"] + %w[README.md]
  spec.require_paths = ["lib"]
  spec.extensions = ["ext/rbxl_native/extconf.rb"]

  spec.add_dependency "rubyzip", "~> 2.3"
  spec.add_dependency "nokogiri", ">= 1.19", "< 2.0"
end
