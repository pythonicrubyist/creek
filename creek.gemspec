# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'creek/version'

Gem::Specification.new do |spec|
  spec.name          = "creek"
  spec.version       = Creek::VERSION
  spec.authors       = ["pythonicrubyist"]
  spec.email         = ["pythonicrubyist@gmail.com"]
  spec.description   = %q{A Ruby gem that streams and parses large Excel(xlsx and xlsm) files fast and efficiently.}
  spec.summary       = %q{A Ruby gem for parsing large Excel(xlsx and xlsm) files.}
  spec.homepage      = "https://github.com/pythonicrubyist/creek"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.required_ruby_version = '>= 2.0.0'

  spec.add_development_dependency "bundler"
  spec.add_development_dependency "rake"
  spec.add_development_dependency 'rspec', '~> 3.6.0'
  spec.add_development_dependency 'pry-byebug'

  spec.add_dependency 'nokogiri', '>= 1.10.0'
  spec.add_dependency 'rubyzip', '>= 1.0.0'
end
