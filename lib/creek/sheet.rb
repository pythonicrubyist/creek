require 'zip/zipfilesystem'
require 'nokogiri'
require 'sax-machine'

module Creek


  class SAXValue
    include SAXMachine
    value :content
  end


  class SAXCell
    include SAXMachine
    attribute :r, :as => :row
    attribute :t, :as => :type
    attribute :s, :as => :sheet
    elements :v, :as => :vals, :class => SAXValue
  end


  class SAXRow
    include SAXMachine
    attribute :collapsed
    attribute :customFormat
    attribute :customHeight
    attribute :hidden
    attribute :ht
    attribute :outlineLevel
    attribute :r, :as => :row_number
    elements :c, :as => :cells, :class => SAXCell
  end


  class SAXSheetData
    include SAXMachine
    elements :row, :as => :rows, :class => SAXRow
  end


  class SAXSheet
    include SAXMachine
    element :worksheet
    elements :sheetData, :as => :sheetDatas, :class => SAXSheetData
  end


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
        doc = @book.files.file.open path
        stream = SAXSheet.parse(doc, :lazy => true)
        Enumerator.new do |y|
          stream.sheetDatas.first.rows.each do |row|
            record = {:row => row.row_number, :cells => {}}
            row.cells.each do |cell|
              shared = (cell.type.eql? 's')
              unless cell.vals.empty?
                content = cell.vals.first.content
                record[:cells][cell.row] = (shared ? @book.shared_strings.dictionary[content.to_i] : content)
              end
            end
            y << fill_in_empty_cells(record)
          end
        end
      end
    end 

    private 
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
