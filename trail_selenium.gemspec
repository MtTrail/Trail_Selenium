# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'trail_selenium/version'

Gem::Specification.new do |spec|
  spec.name          = "trail_selenium"
  spec.version       = TrailSelenium::VERSION
  spec.authors       = ["Mt.Trail"]
  spec.email         = ["trail@trail4you.com"]

  spec.summary       = %q{Handling selenium from ruby script}
  spec.description   = %q{Handling selenium from ruby script}
  spec.homepage      = "http://www.trail4you.com/TechNote/Ruby/selenium.html"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.12"
  spec.add_development_dependency "rake", "~> 10.0"
end
