require 'zip/filesystem'
require 'nokogiri'

module Creek

  class Creek::SharedStrings

    attr_reader :book, :dictionary

    def initialize book
      @book = book
      parse_shared_shared_strings
    end

    def parse_shared_shared_strings
      path = "xl/sharedStrings.xml"
      @dictionary = Hash.new
      if @book.files.file.exist?(path)
        doc = @book.files.file.open path
        xml = Nokogiri::XML::Document.parse doc
        xml.css('t').each_with_index.map  do |str, i|
          @dictionary[i] = str.content
        end
      end
    end
  end
end