lib = File.expand_path('lib', __dir__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'creek/version'

Gem::Specification.new do |spec|
  spec.name          = 'creek'
  spec.version       = Creek::VERSION
  spec.authors       = ['pythonicrubyist']
  spec.email         = ['pythonicrubyist@gmail.com']
  spec.description   = 'A Ruby gem that streams and parses large Excel(xlsx and xlsm) files fast and efficiently.'
  spec.summary       = 'A Ruby gem for parsing large Excel(xlsx and xlsm) files.'
  spec.homepage      = 'https://github.com/pythonicrubyist/creek'
  spec.license       = 'MIT'

  spec.files         = `git ls-files lib`.split($INPUT_RECORD_SEPARATOR) + %w(LICENSE.txt README.md)
  spec.require_paths = ['lib']

  spec.required_ruby_version = '>= 2.0.0'

  spec.add_development_dependency 'bundler'
  spec.add_development_dependency 'pry-byebug'
  spec.add_development_dependency 'rake'
  spec.add_development_dependency 'rspec', '~> 3.6.0'
  spec.add_development_dependency 'rubocop'

  spec.add_dependency 'nokogiri', '>= 1.10.0'
  spec.add_dependency 'rubyzip', '>= 1.0.0'
end
