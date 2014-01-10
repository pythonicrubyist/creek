require 'zip/filesystem'
require 'nokogiri'

module Creek
  class Creek::Sheet

    attr_reader :book,
                :name,
                :sheetid,
                :state,
                :visible,
                :rid,
                :index


    def initialize book, name, sheetid, state, visible, rid, index
      @book = book
      @name = name
      @sheetid = sheetid
      @visible = visible
      @rid = rid
      @state = state
      @index = index

      # An XLS file has only 256 columns, however, an XLSX or XLSM file can contain up to 16384 columns.
      # This function creates a hash with all valid XLSX column names and associated indices.
      @@excel_col_names = Hash.new
      (0...16384).each do |i|
        @@excel_col_names[col_name(i)] = i
      end
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
    # Returns valid Excel column name for a given column index.
    # For example, returns "A" for 0, "B" for 1 and "AQ" for 42.
    def col_name i
      quot = i/26
      (quot>0 ? col_name(quot-1) : "") + (i%26+65).chr
    end

    ##
    # Returns a hash per row that includes the cell ids and values.
    # Empty cells will be also included in the hash with a nil value.
    def rows_generator include_meta_data=false
      path = "xl/worksheets/sheet#{@index}.xml"
      if @book.files.file.exist?(path)
        # SAX parsing, Each element in the stream comes through as two events:
        # one to open the element and one to close it.
        opener = Nokogiri::XML::Reader::TYPE_ELEMENT
        closer = Nokogiri::XML::Reader::TYPE_END_ELEMENT
        Enumerator.new do |y|
          shared, row, cells, cell = false, nil, {}, nil
          @book.files.file.open(path) do |xml|
            Nokogiri::XML::Reader.from_io(xml).each do |node|
              if (node.name.eql? 'row') and (node.node_type.eql? opener)
                row = node.attributes
                row['cells'] = Hash.new
                cells = Hash.new
                y << (include_meta_data ? row : cells) if node.self_closing?
              elsif (node.name.eql? 'row') and (node.node_type.eql? closer)
                processed_cells = fill_in_empty_cells(cells, row['r'], cell)
                row['cells'] = processed_cells
                y << (include_meta_data ? row : processed_cells)
              elsif (node.name.eql? 'c') and (node.node_type.eql? opener)
                  shared = node.attribute('t').eql? 's'
                  cell = node.attribute('r')
              elsif node.value?
                if shared
                  cells[cell] = @book.shared_strings.dictionary[node.value.to_i] if @book.shared_strings.dictionary.has_key? node.value.to_i
                else
                  cells[cell] = node.value
                end
              end
            end
          end
        end
      end
    end
    ##
    # The unzipped XML file does not contain any node for empty cells.
    # Empty cells are being padded in using this function
    def fill_in_empty_cells cells, row_number, last_col
      new_cells = Hash.new
      unless cells.empty?
        keys = cells.keys.sort
        last_col = last_col.gsub(row_number, '')
        last_col_index = @@excel_col_names[last_col]
        [*(0..last_col_index)].each do |i|
          col = col_name i
          id = "#{col}#{row_number}"
          unless cells.has_key? id
              new_cells[id] = nil
          else
            new_cells[id] = cells[id]
          end
        end
      end
      new_cells
    end
  end
end
