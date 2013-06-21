require 'zip/zipfilesystem'
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

    # Returns valid Excel column name for a given column index. 
    # For example, returns "A" for 0, "B" for 1 and "AQ" for 42.
    def col_name i
      quot = i/26
      (quot>0 ? col_name(quot-1) : "") + (i%26+65).chr
    end


    # This will return a hash per row that includes the column names and cell values.
    # Empty cells will be also included in the hash with a nil value.
    def rows
      path = "xl/worksheets/sheet#{@index}.xml"
      if @book.files.file.exist?(path)
        # SAX parsing, Each element in the stream comes through as two events:
        # one to open the element and one to close it.
        opener = Nokogiri::XML::Reader::TYPE_ELEMENT
        closer = Nokogiri::XML::Reader::TYPE_END_ELEMENT
        Enumerator.new do |y|
          shared, row, cell = false, nil, nil
          @book.files.file.open(path) do |xml|
            Nokogiri::XML::Reader.from_io(xml).each do |node|
              if (node.name.eql? 'row') and (node.node_type.eql? opener)
                row = {:row => node.attribute('r'), :cells => {}}
              elsif (node.name.eql? 'row') and (node.node_type.eql? closer)
                y << fill_in_empty_cells(row)
              elsif (node.name.eql? 'c') and (node.node_type.eql? opener)
                  shared = node.attribute('t').eql? 's'
                  cell = node.attribute('r')
              elsif node.value?
                if shared
                  row[:cells][cell] = @book.shared_strings.dictionary[node.value.to_i] if @book.shared_strings.dictionary.has_key? node.value.to_i
                else
                  row[:cells][cell] = node.value
                end
              end
            end
          end
        end
      end
    end


    # The unzipped XML file does not contain any node for empty cells.
    # Empty cells are being padded in using this function
    def fill_in_empty_cells row
      cells = Hash.new
      unless row[:cells].empty?
        keys = row[:cells].keys.sort
        last_col =  keys.last.gsub(row[:row], '')
        last_col_index = @@excel_col_names[last_col]
        [*(0..last_col_index)].each do |i|
          col = col_name i
          id = "#{col}#{row[:row]}"
          unless row[:cells].has_key? id
              cells[id] = nil
          else
            cells[id] = row[:cells][id] 
          end
        end
      end
      cells
    end
  end
end