# frozen_string_literal: true

module Creek
  class Styles
    attr_accessor :book
    def initialize(book)
      @book = book
    end

    def path
      "xl/styles.xml"
    end

    def styles_xml
      @styles_xml ||= begin
        if @book.files.file.exist?(path)
          doc = @book.files.file.open path
          Nokogiri::XML::Document.parse doc
        end
      end
    end

    def style_types
      @style_types ||= begin
        Creek::Styles::StyleTypes.new(styles_xml).call
      end
    end
  end
end
