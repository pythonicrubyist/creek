require 'creek'

describe 'Creek trying to parsing an invalid file.' do
  it 'open an XLSX file successfully.' do
    lambda { Creek::Book.new 'specs/fixtures/invalid.xls' }.should raise_error 'Not a valid file format.'
  end
end

describe 'Creek parsing a sample XLSX file' do
  before(:all) do
    @creek = Creek::Book.new 'spec/fixtures/sample.xlsx'
    @expected_rows = [{'A1'=>'Content 1', 'B1'=>nil, 'C1'=>'Content 2', 'D1'=>nil, 'E1'=>'Content 3'},
                    {'A2'=>nil, 'B2'=>'Content 4', 'C2'=>nil, 'D2'=>'Content 5', 'E2'=>nil, 'F2'=>'Content 6'},
                    {},
                    {'A4'=>'Content 7', 'B4'=>'Content 8', 'C4'=>'Content 9', 'D4'=>'Content 10', 'E4'=>'Content 11', 'F4'=>'Content 12'},
                    {'A5'=>nil, 'B5'=>nil, 'C5'=>nil, 'D5'=>nil, 'E5'=>nil, 'F5'=>nil, 'G5'=>nil, 'H5'=>nil, 'I5'=>nil, 'J5'=>nil, 'K5'=>nil, 'L5'=>nil, 'M5'=>nil, 'N5'=>nil, 'O5'=>nil, 'P5'=>nil, 'Q5'=>nil, 'R5'=>nil, 'S5'=>nil, 'T5'=>nil, 'U5'=>nil, 'V5'=>nil, 'W5'=>nil, 'X5'=>nil, 'Y5'=>nil, 'Z5'=>nil, 'AA5'=>nil, 'AB5'=>nil, 'AC5'=>nil, 'AD5'=>nil, 'AE5'=>nil, 'AF5'=>nil, 'AG5'=>nil, 'AH5'=>nil, 'AI5'=>nil, 'AJ5'=>nil, 'AK5'=>nil, 'AL5'=>nil, 'AM5'=>nil, 'AN5'=>nil, 'AO5'=>nil, 'AP5'=>nil, 'AQ5'=>nil, 'AR5'=>nil, 'AS5'=>nil, 'AT5'=>nil, 'AU5'=>nil, 'AV5'=>nil, 'AW5'=>nil, 'AX5'=>nil, 'AY5'=>nil, 'AZ5'=>'Content 13'},
                    {'A6'=>'1', 'B6'=>'2', 'C6'=>'3'}, {'A7'=>'Content 15', 'B7'=>'Content 16', 'C7'=>'Content 18', 'D7'=>'Content 19'},
                    {'A8'=>nil, 'B8'=>'Content 20', 'C8'=>nil, 'D8'=>nil, 'E8'=>nil, 'F8'=>'Content 21'}]
  end

  after(:all) do
    @creek.close
  end

  it 'open an XLSX file successfully.' do
    @creek.should_not be_nil
  end

  it 'find sheets successfully.' do
    @creek.sheets.count.should == 1
    sheet = @creek.sheets.first
    sheet.state.should eql 'visible'
    sheet.name.should eql 'Sheet1'
    sheet.rid.should eql 'rId2'
  end

  it 'Parse rows with empty cells successfully.' do
    rows = Array.new
    row_count = 0
    @creek.sheets[0].rows.each do |row|
      rows << row     
      row_count += 1
    end

    rows.should == @expected_rows

    @creek.sheets[0].rows.first.count.should == 5
    row_count.should == 8
  end

  it 'Parse rows with empty cells and meta data successfully.' do
    rows = Array.new
    row_count = 0
    @creek.sheets[0].rows_with_meta_data.each do |row|
      rows << row     
      row_count += 1
    end
    rows.map{|r| r['cells']}.should == @expected_rows
  end
end
