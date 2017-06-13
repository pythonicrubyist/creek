require 'zip/filesystem'
require 'nokogiri'

module Creek
  class Creek::Sheet
    include Creek::Utils

    attr_reader :book,
                :name,
                :sheetid,
                :state,
                :visible,
                :rid,
                :index


    def initialize book, name, sheetid, state, visible, rid, sheetfile
      @book = book
      @name = name
      @sheetid = sheetid
      @visible = visible
      @rid = rid
      @state = state
      @sheetfile = sheetfile
      @images_present = false
    end

    ##
    # Preloads images info (coordinates and paths) from related drawing.xml and drawing rels.
    # Must be called before #rows method if you want to have images included.
    # Returns self so you can chain the calls (sheet.with_images.rows).
    def with_images
      @drawingfile = extract_drawing_filepath
      if @drawingfile
        @drawing = Creek::Drawing.new(@book, @drawingfile.sub('..', 'xl'))
        @images_present = @drawing.has_images?
      end
      self
    end

    ##
    # Extracts images for a cell to a temporary folder.
    # Returns array of Pathnames for the cell.
    # Returns nil if images asre not found for the cell or images were not preloaded with #with_images.
    def images_at(cell)
      @drawing.images_at(cell) if @images_present
    end

    ##
    # Provides an Enumerator that returns a hash representing each row.
    # The key of the hash is the Cell id and the value is the value of the cell.
    def rows
      rows_generator
    end

    ##
    # Provides an Enumerator that returns a hash representing each row.
    # The hash contains meta data of the row and a 'cells' embended hash which contains the cell contents.
    def rows_with_meta_data
      rows_generator true
    end

    private

    ##
    # Returns a hash per row that includes the cell ids and values.
    # Empty cells will be also included in the hash with a nil value.
    def rows_generator include_meta_data=false
      path = if @sheetfile.start_with? "/xl/" or @sheetfile.start_with? "xl/" then @sheetfile else "xl/#{@sheetfile}" end
      if @book.files.file.exist?(path)
        # SAX parsing, Each element in the stream comes through as two events:
        # one to open the element and one to close it.
        opener = Nokogiri::XML::Reader::TYPE_ELEMENT
        closer = Nokogiri::XML::Reader::TYPE_END_ELEMENT
        Enumerator.new do |y|
          row, cells, cell = nil, {}, nil
          cell_type  = nil
          cell_style_idx = nil
          @book.files.file.open(path) do |xml|
            Nokogiri::XML::Reader.from_io(xml).each do |node|
              if (node.name.eql? 'row') and (node.node_type.eql? opener)
                row = node.attributes
                row['cells'] = Hash.new
                cells = Hash.new
                y << (include_meta_data ? row : cells) if node.self_closing?
              elsif (node.name.eql? 'row') and (node.node_type.eql? closer)
                processed_cells = fill_in_empty_cells(cells, row['r'], cell)

                if @images_present
                  processed_cells.each do |cell_name, cell_value|
                    next unless cell_value.nil?
                    processed_cells[cell_name] = images_at(cell_name)
                  end
                end

                row['cells'] = processed_cells
                y << (include_meta_data ? row : processed_cells)
              elsif (node.name.eql? 'c') and (node.node_type.eql? opener)
                cell_type      = node.attributes['t']
                cell_style_idx = node.attributes['s']
                cell           = node.attributes['r']
              elsif (node.name.eql? 'v') and (node.node_type.eql? opener)
                unless cell.nil?
                  cells[cell] = convert(node.inner_xml, cell_type, cell_style_idx)
                end
              elsif (node.name.eql? 't') and (node.node_type.eql? opener)
                unless cell.nil?
                  cells[cell] = convert(node.inner_xml, cell_type, cell_style_idx)
                end
              end
            end
          end
        end
      end
    end

    def convert(value, type, style_idx)
      style = @book.style_types[style_idx.to_i]
      Creek::Styles::Converter.call(value, type, style, converter_options)
    end

    def converter_options
      @converter_options ||= {
        shared_strings: @book.shared_strings.dictionary,
        base_date: @book.base_date
      }
    end

    ##
    # The unzipped XML file does not contain any node for empty cells.
    # Empty cells are being padded in using this function
    def fill_in_empty_cells(cells, row_number, last_col)
      new_cells = Hash.new

      unless cells.empty?
        last_col = last_col.gsub(row_number, '')

        ("A"..last_col).to_a.each do |column|
          id = "#{column}#{row_number}"
          new_cells[id] = cells[id]
        end
      end

      new_cells
    end

    ##
    # Find drawing filepath for the current sheet.
    # Sheet xml contains drawing relationship ID.
    # Sheet relationships xml contains drawing file's location.
    def extract_drawing_filepath
      # Read drawing relationship ID from the sheet.
      sheet_filepath = "xl/#{@sheetfile}"
      drawing = parse_xml(sheet_filepath).css('drawing').first
      return if drawing.nil?

      drawing_rid = drawing.attributes['id'].value

      # Read sheet rels to find drawing file's location.
      sheet_rels_filepath = expand_to_rels_path(sheet_filepath)
      parse_xml(sheet_rels_filepath).css("Relationship[@Id='#{drawing_rid}']").first.attributes['Target'].value
    end
  end
end
