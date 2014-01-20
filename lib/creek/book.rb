require 'zip/filesystem'
require 'nokogiri'

module Creek

  class Creek::Book

    attr_reader :files,
                :sheets,
                :shared_strings

    def initialize path
      raise 'Not a valid file format.' unless (['.xlsx', '.xlsm'].include? File.extname(path).downcase)
      @files = Zip::File.open path
      @shared_strings = Creek::SharedStrings.new(self)
    end

    def sheets
      doc = @files.file.open "xl/workbook.xml"
      xml = Nokogiri::XML::Document.parse doc
      @sheets = xml.css('sheet').each_with_index.map  do |sheet, i|
        Creek::Sheet.new(self, sheet.attr("name"), sheet.attr("sheetid"),  sheet.attr("state"), sheet.attr("visible"), sheet.attr("r:id"), i+1)
      end
    end

    def close
      @files.close
    end
  end
end