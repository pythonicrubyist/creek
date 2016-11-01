require 'pathname'

module Creek
  class Drawing

    COLUMNS = ('A'..'AZ').to_a

    def initialize(book, sheetfile)
      @book = book
      @sheetfile = sheetfile
      @images_pathnames = Hash.new { |hash, key| hash[key] = [] }

      load_drawings_and_rels
      load_images_pathnames_by_cells if has_images?
    end

    def has_images?
      @has_images ||= !@drawings.nil? && @drawings.size > 0
    end

    def images_at(cell_name)
      coordinate = calc_coordinate(cell_name)
      pathnames_at_coordinate = @images_pathnames[coordinate]
      return if pathnames_at_coordinate.empty?

      pathnames_at_coordinate.map do |image_pathname|
         if image_pathname.exist?
           image_pathname
         else
           excel_image_path = "xl/media/#{image_pathname.to_path.split(tmpdir).last}"
           IO.copy_stream(@book.files.file.open(excel_image_path), image_pathname.to_path)
           image_pathname
         end
      end
    end

    private

    def calc_coordinate(cell_name)
      col = COLUMNS.index(cell_name.slice /[A-Z]+/)
      row = (cell_name.slice /\d+/).to_i - 1 # rows in drawings start with 0
      [row, col]
    end

    # for saving extracted images from excel
    def tmpdir
      @tmpdir ||= ::Dir.mktmpdir('creek__drawing')
    end

    def load_drawings_and_rels
      drawing_filepath = extract_drawing_filepath
      return if drawing_filepath.nil?

      drawing_filepath.sub!('..', 'xl')
      @drawings = load_drawings(drawing_filepath)
      @drawings_rels = load_drawings_rels(drawing_filepath)
    end

    def load_drawings(drawing_filepath)
      parse_xml(drawing_filepath).css('xdr|twoCellAnchor')
    end

    def load_drawings_rels(drawing_filepath)
      drawing_rels_filepath = expand_to_rels_path(drawing_filepath)
      parse_xml(drawing_rels_filepath).css('Relationships')
    end

    def extract_drawing_filepath
      # read drawing relationship ID from the sheet
      sheet_filepath = "xl/#{@sheetfile}"
      drawing = parse_xml(sheet_filepath).css('drawing').first
      return if drawing.nil?

      drawing_rid = drawing.attributes['id'].value

      # read sheet rels to find drawing file location
      sheet_rels_filepath = expand_to_rels_path(sheet_filepath)
      parse_xml(sheet_rels_filepath).css("Relationship[@Id='#{drawing_rid}']").first.attributes['Target'].value
    end

    def expand_to_rels_path(filepath)
      filepath.sub(/(\/[^\/]+$)/, '/_rels\1.rels')
    end

    def parse_xml(xml_path)
      doc = @book.files.file.open(xml_path)
      Nokogiri::XML::Document.parse(doc)
    end

    def load_images_pathnames_by_cells
      image_selector = 'xdr:pic/xdr:blipFill/a:blip'.freeze
      row_from_selector = 'xdr:from/xdr:row'.freeze
      row_to_selector = 'xdr:to/xdr:row'.freeze
      col_from_selector = 'xdr:from/xdr:col'.freeze
      col_to_selector = 'xdr:to/xdr:col'.freeze

      @drawings.xpath('//xdr:twoCellAnchor').each do |drawing|
        embed = drawing.xpath(image_selector).first.attributes['embed']
        next if embed.nil?

        rid = embed.value
        path = Pathname.new("#{tmpdir}#{extract_drawing_path(rid).slice(/[^\/]*$/)}")

        row_from = drawing.xpath(row_from_selector).text.to_i
        col_from = drawing.xpath(col_from_selector).text.to_i
        row_to = drawing.xpath(row_to_selector).text.to_i
        col_to = drawing.xpath(col_to_selector).text.to_i

        (col_from..col_to).each do |col|
          (row_from..row_to).each do |row|
            @images_pathnames[[row, col]].push(path)
          end
        end
      end
    end

    def extract_drawing_path(rid)
      @drawings_rels.css("Relationship[@Id=#{rid}]").first.attributes['Target'].value
    end
  end
end