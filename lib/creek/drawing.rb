require 'pathname'

module Creek
  class Creek::Drawing
    include Creek::Utils

    COLUMNS = ('A'..'AZ').to_a

    def initialize(book, drawing_filepath)
      @book = book
      @drawing_filepath = drawing_filepath
      @drawings = []
      @drawings_rels = []
      @images_pathnames = Hash.new { |hash, key| hash[key] = [] }

      if file_exist?(@drawing_filepath)
        load_drawings_and_rels
        load_images_pathnames_by_cells if has_images?
      end
    end

    ##
    # Returns false if there are no images in the drawing file or the drawing file does not exist, true otherwise.
    def has_images?
      @has_images ||= !@drawings.empty?
    end

    ##
    # Extracts images from excel to tmpdir for a cell, if the images are not already extracted (multiple calls or same image file in multiple cells).
    # Returns array of images as Pathname objects or nil.
    def images_at(cell_name)
      coordinate = calc_coordinate(cell_name)
      pathnames_at_coordinate = @images_pathnames[coordinate]
      return if pathnames_at_coordinate.empty?

      pathnames_at_coordinate.map do |image_pathname|
        if image_pathname.exist?
          image_pathname
        else
          excel_image_path = "xl/media#{image_pathname.to_path.split(tmpdir).last}"
          IO.copy_stream(@book.files.file.open(excel_image_path), image_pathname.to_path)
          image_pathname
         end
      end
    end

    private

    ##
    # Transforms cell name to [row, col], e.g. A1 => [0, 0], B3 => [1, 2]
    # Rows and cols start with 0.
    def calc_coordinate(cell_name)
      col = COLUMNS.index(cell_name.slice /[A-Z]+/)
      row = (cell_name.slice /\d+/).to_i - 1 # rows in drawings start with 0
      [row, col]
    end

    ##
    # Creates/loads temporary directory for extracting images from excel
    def tmpdir
      @tmpdir ||= ::Dir.mktmpdir('creek__drawing')
    end

    ##
    # Parses drawing and drawing's relationships xmls.
    # Drawing xml contains relationships ID's and coordinates (row, col).
    # Drawing relationships xml contains images' locations.
    def load_drawings_and_rels
      @drawings = parse_xml(@drawing_filepath).css('xdr|twoCellAnchor')
      drawing_rels_filepath = expand_to_rels_path(@drawing_filepath)
      @drawings_rels = parse_xml(drawing_rels_filepath).css('Relationships')
    end

    ##
    # Iterates through the drawings and saves images' paths as Pathname objects to a hash with [row, col] keys.
    # As multiple images can be located in a single cell, hash values are array of Pathname objects.
    # One image can be spread across multiple cells (defined with from-row/to-row/from-col/to-col attributes) - same Pathname object is associated to each row-col combination for the range.
    def load_images_pathnames_by_cells
      image_selector = 'xdr:pic/xdr:blipFill/a:blip'.freeze
      row_from_selector = 'xdr:from/xdr:row'.freeze
      row_to_selector = 'xdr:to/xdr:row'.freeze
      col_from_selector = 'xdr:from/xdr:col'.freeze
      col_to_selector = 'xdr:to/xdr:col'.freeze

      @drawings.xpath('//xdr:twoCellAnchor').each do |drawing|
        # embed = drawing.xpath(image_selector).first.attributes['embed']
        temp = drawing.xpath(image_selector).first
        embed = temp.attributes['embed'] if temp
        next if embed.nil?

        rid = embed.value
        path = Pathname.new("#{tmpdir}/#{extract_drawing_path(rid).slice(/[^\/]*$/)}")

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
