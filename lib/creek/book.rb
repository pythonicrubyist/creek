require 'zip/filesystem'
require 'nokogiri'

module Creek

  class Creek::Book

    attr_reader :files,
                :sheets,
                :shared_strings

    def initialize path, options = {}
      check_file_extension = options.fetch(:check_file_extension, true)
      if check_file_extension
        extension = File.extname(options[:original_filename] || path).downcase
        raise 'Not a valid file format.' unless (['.xlsx', '.xlsm'].include? extension)
      end
      @files = Zip::File.open path
      @shared_strings = SharedStrings.new(self)
    end

    def sheets
      doc = @files.file.open "xl/workbook.xml"
      xml = Nokogiri::XML::Document.parse doc
      @sheets = xml.css('sheet').each_with_index.map  do |sheet, i|
        Sheet.new(self, sheet.attr("name"), sheet.attr("sheetid"),  sheet.attr("state"), sheet.attr("visible"), sheet.attr("r:id"), i+1)
      end
    end

    def style_types
      @style_types ||= Creek::Styles.new(self).style_types
    end

    def close
      @files.close
    end
  end
end
