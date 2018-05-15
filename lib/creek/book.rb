require 'zip/filesystem'
require 'nokogiri'
require 'date'
require 'http'

module Creek

  class Creek::Book

    attr_reader :files,
                :sheets,
                :shared_strings

    DATE_1900 = Date.new(1899, 12, 30).freeze
    DATE_1904 = Date.new(1904, 1, 1).freeze

    def initialize path, options = {}
      check_file_extension = options.fetch(:check_file_extension, true)
      if check_file_extension
        extension = File.extname(options[:original_filename] || path).downcase
        raise 'Not a valid file format.' unless (['.xlsx', '.xlsm'].include? extension)
      end
      if options[:remote]
        zipfile = Tempfile.new("file")
        zipfile.binmode
        zipfile.write(HTTP.get(path).to_s)
        zipfile.close
        path = zipfile.path
      end
      @files = Zip::File.open(path)
      @shared_strings = SharedStrings.new(self)
    end

    #
    # List of xlsx worksheets
    #
    # @return [Array<Creek::Sheet>] worksheets array
    #
    def sheets
      @sheets = document.css(['sheet']).map do |sheet|
        sheetfile = document.relationships
          .css("Relationship[@Id=#{sheet.attr('r:id')}]")
          .first.attributes['Target'].value
        Sheet.new(self, sheet.attr("name"), sheet.attr("sheetid"),  sheet.attr("state"), sheet.attr("visible"), sheet.attr("r:id"), sheetfile)
      end
    end

    def style_types
      @style_types ||= Creek::Styles.new(self).style_types
    end

    def close
      @files.close
    end

    def base_date
      @base_date ||= begin
        # Default to 1900 (minus one day due to excel quirk) but use 1904 if
        # it's set in the Workbook's workbookPr
        # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx

        workbook_pr_1904 = document.css(['workbookPr[date1904]'])
          .find { |w_pr| w_pr['date1904'] =~ /true|1/i }
        workbook_pr_1904.nil? ? DATE_1900 : DATE_1904
      end
    end

    #
    # Document representing workbook
    #
    # @return [Creek::Document] workbook
    #
    def document
      @workbook ||= Document.new(self, 'xl/workbook.xml')
    end
  end
end
