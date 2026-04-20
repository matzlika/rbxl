require_relative "lib/rbxl/version"

Gem::Specification.new do |spec|
  spec.name = "rbxl"
  spec.version = Rbxl::VERSION
  spec.authors = ["Taro KOBAYASHI"]
  spec.email = ["taro@matzlika.co.jp"]
  spec.summary = "Fast, low-memory XLSX processing for Ruby."
  spec.description = "rbxl is a fast, low-memory Ruby gem for row-by-row XLSX reads and append-only XLSX writes, with an optional native extension for higher-throughput XML parsing."
  spec.homepage = "https://github.com/matzlika/rbxl"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 3.1"
  spec.metadata = {
    "homepage_uri" => spec.homepage,
    "source_code_uri" => "https://github.com/matzlika/rbxl",
    "bug_tracker_uri" => "https://github.com/matzlika/rbxl/issues",
    "changelog_uri" => "https://github.com/matzlika/rbxl/releases"
  }

  spec.files = Dir["lib/**/*.rb"] + Dir["ext/**/*.{rb,c,h}"] + Dir["sig/**/*.rbs"] + %w[CHANGELOG.md LICENSE.txt README.md Rakefile]
  spec.require_paths = ["lib"]
  spec.extensions = ["ext/rbxl_native/extconf.rb"]

  spec.add_dependency "rubyzip", "~> 2.3"
  spec.add_dependency "nokogiri", ">= 1.19", "< 2.0"
end
