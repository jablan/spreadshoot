# -*- encoding: utf-8 -*-
require File.expand_path('../lib/spreadshoot/version', __FILE__)

Gem::Specification.new do |gem|
  gem.authors       = ["Mladen Jablanovic"]
  gem.email         = ["jablan@radioni.ca"]
  gem.description   = %q{Create XLSX files from scratch using Ruby}
  gem.summary       = %q{Ruby DSL for creating Excel xlsx spreadsheets}
  gem.homepage      = "https://github.com/jablan/spreadshoot"

  gem.add_dependency "builder"

  gem.files         = `git ls-files`.split($\)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.name          = "spreadshoot"
  gem.require_paths = ["lib"]
  gem.version       = Spreadshoot::VERSION
end
