# frozen_string_literal: true

require "bundler/gem_helper"
require "rake/testtask"
require "rdoc/task"

Bundler::GemHelper.install_tasks

Rake::TestTask.new(:test) do |t|
  t.libs << "test"
  t.libs << "lib"
  t.test_files = FileList["test/**/*_test.rb"]
  t.warning = false
end

RDoc::Task.new(:rdoc) do |rdoc|
  rdoc.main = "README.md"
  rdoc.rdoc_files.include("README.md", "lib/**/*.rb")
end

desc "Build the rbxl_native C extension in place"
task :compile do
  ext_dir = File.expand_path("ext/rbxl_native", __dir__)
  Dir.chdir(ext_dir) do
    ruby "extconf.rb"
    sh "make"
  end
end

task test: :compile
