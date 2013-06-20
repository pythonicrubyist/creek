require 'zip/zipfilesystem'
require 'nokogiri'

module Creek

  class Creek::SharedStrings

    attr_reader :workbook, :dictionary

    def initialize book
      @book = book
      parse_shared_shared_strings
    end

    def parse_shared_shared_strings
      @dictionary = Hash.new
      doc = @book.files.file.open "xl/sharedStrings.xml"
      xml = Nokogiri::XML::Document.parse doc
      @sheets = xml.css('t').each_with_index.map  do |str, i|
        @dictionary[i] = str.content
      end
    end
  end
end