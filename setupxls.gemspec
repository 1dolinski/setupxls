$LOAD_PATH.unshift 'lib'
require 'setupxls/version'

Gem::Specification.new do |s|
  s.name              = "setupxls"
  s.version           = SetupXls::Version
  s.date              = Time.now.strftime('%Y-%m-%d')
  s.summary           = "Access an open Excel workbook worksheet."
  s.homepage          = "http://github.com/ebbflowgo/setupxls"
  s.email             = "ebbflowgo@gmail.com"
  s.authors           = ["ebb flowgo"]

  s.files             = %w( README.md Rakefile LICENSE )
  s.files            += Dir.glob("lib/**/*")
  s.files            += Dir.glob("spec/**/*")
  s.files            += Dir.glob("test/**/*")

  s.description = <<description
    Setupxls allows you to connect ruby to an excel worksheet.
description
end
