require 'zip/filesystem'
require 'nokogiri'

module Creek

  class Creek::Book

    attr_reader :files,
                :sheets,
                :shared_strings

<<<<<<< HEAD
    def initialize path
      raise 'Not a valid file format.' unless (['.xlsx', '.xlsm'].include? File.extname(path).downcase)
      @files = Zip::File.open path
=======
    def initialize path, options = {}
      check_file_extension = options.fetch(:check_file_extension, true)
      if check_file_extension
        extension = File.extname(options[:original_filename] || path).downcase
        raise 'Not a valid file format.' unless (['.xlsx', '.xlsm'].include? extension)
      end
      @files = Zip::ZipFile.open path
>>>>>>> 7161c30c5c78609e689a4e5e36bd3f2d5884684c
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