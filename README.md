= Creek -- Stream parser for large Excel(xlsx and xlsm) files.
Inspired by [Dullard](https://github.com/thirtyseven/dullard), Creek is a fast, smple and efficient way of parsing XLSX and XLSXM files using SAX parsing.
* Can simply parse a large Excel file by looping through the rows enumerator:
   require 'creek'
   creek = Creek::Book.new "specs/fixtures/sample.xlsx"
   creek.sheets[0].rows.each do |row|
      puts row.inspect
      # => {"A1"=>"Content 1", "B1"=>nil, "C1"=>"Content 2", "D1"=>nil, "E1"=>"Content 3"}
   end