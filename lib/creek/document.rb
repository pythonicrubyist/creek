module Creek
  # Decorates xml document
  class Document
    include Creek::Utils
    attr_reader :file_path, :book

    #
    # Initializer
    #
    # Note: Logically book parameter is an unnecessary dependency. It kept
    #       here for the sake of reducing code refactoring volume.
    #
    # @param [Creek::Book] book Workbook
    # @param [String] file_path Internal xlsx-file path
    #
    def initialize(book, file_path)
      @book = book
      @file_path = file_path
    end

    #
    # Loads the XML file located by file_path
    #
    # @return [Nokogiri::XML::Document] parsed xml file
    #
    def xml
      @xml ||= parse_xml(file_path)
    end

    #
    # Resources information related to xml-file
    #
    # @return [Nokogiri::XML::NodeSet] Resources elements
    #
    def relationships
      @relationships ||= begin
        xml = parse_xml(rels_filepath)
        xml ? xml.css('Relationships') : []
      end
    end

    #
    # Path to XML relationship file
    #
    # @return [String] file path
    #
    def rels_filepath
      @rels_filepath ||= expand_to_rels_path(file_path)
    end

    #
    # Root xml namespace prefix
    #
    # @return [String, nil] namespace prefix
    #
    def namespace_prefix
      @namespace_prefix ||= xml.root.namespace.prefix
    end

    #
    # Returns xml-tree elements by selector_names path
    # Automatically adds root namespace prefix
    #
    # @param [Array<String>] selector_names Required elements path
    #
    # @return [Nokogiri::XML::NodeSet] Data elements
    #
    def css(selector_names)
      selector = css_selector(selector_names)
      xml.css(selector)
    end

    #
    # Builds CSS selector
    #
    # @param [Array] selector_names Selectors list
    #
    # @return [String] xml tree path
    #
    def css_selector(selector_names)
      selectors = selector_names if namespace_prefix.nil?
      selectors ||= selector_names
        .map { |selector| "#{namespace_prefix}|#{selector}" }
      selectors.join(' ')
    end
  end
end
