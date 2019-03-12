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
      if @book.files.file.exist?(path)
        doc = @book.files.file.open path
        xml = Nokogiri::XML::Document.parse doc
        parse_shared_string_from_document(xml)
      end
    end

    def parse_shared_string_from_document(xml)
      @dictionary = self.class.parse_shared_string_from_document(xml)
    end

    def self.parse_shared_string_from_document(xml)
      dictionary = Hash.new

      xml.css('si').each_with_index do |si, idx|
        text_nodes = si.css('>t, r t')
        if text_nodes.count == 1 # plain text node
          dictionary[idx] = Creek::Styles::Converter.unescape_string(text_nodes.first.content)
        else # rich text nodes with text fragments
          dictionary[idx] = text_nodes.map { |n| Creek::Styles::Converter.unescape_string(n.content) }.join('')
        end
      end

      dictionary
    end

  end
end
