module Creek
  module Utils
    def expand_to_rels_path(filepath)
      filepath.sub(/(\/[^\/]+$)/, '/_rels\1.rels')
    end

    def file_exist?(path)
      return false unless path

      @book.nil? ? File.exist?(path) : @book.files.file.exist?(path)
    end

    def parse_xml(path)
      return unless file_exist?(path)

      file = @book.nil? ? File.open(path) : @book.files.file.open(path)
      Nokogiri::XML::Document.parse(file)
    end
  end
end
