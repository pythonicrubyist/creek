module Creek
  class Styles
    attr_accessor :book
    def initialize(book)
      @book = book
    end

    def path
      "xl/styles.xml"
    end

    #
    # Document representing styles file
    #
    # @return [Creek::Document] styles file
    #
    def document
      return unless @book.files.file.exist?(path)
      Document.new(book, path)
    end

    def style_types
      @style_types ||= begin
        Creek::Styles::StyleTypes.new(document).call
      end
    end
  end
end
