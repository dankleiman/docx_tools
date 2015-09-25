lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'docx_tools/version'

Gem::Specification.new do |spec|
  spec.name          = 'docx_tools'
  spec.version       = DocxTools::VERSION
  spec.authors       = ['Kevin Deisz']
  spec.email         = ['info@trialnetworks.com']
  spec.homepage      = 'https://github.com/drugdev/docx_tools'

  spec.summary       = 'Tools for manipulating docx files'
  spec.description   = 'An API for managing merge fields within docx files'
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test)/}) }
  spec.require_paths = ['lib']

  spec.add_runtime_dependency 'nokogiri', '~> 1.6', '>= 1.6.6'
  spec.add_runtime_dependency 'rubyzip', '~> 1.1', '>= 1.1.7'

  spec.add_development_dependency 'minitest', '~> 5.8'
  spec.add_development_dependency 'simplecov', '~> 0.10'
  spec.add_development_dependency 'yard', '~> 0.8'
end
