require './spec/spec_helper'

describe 'Creek trying to parsing an invalid file.' do
  it 'Fail to open a legacy xls file.' do
    expect { Creek::Book.new 'spec/fixtures/invalid.xls' }
      .to raise_error 'Not a valid file format.'
  end

  it 'Ignore file extensions on request.' do
    path = 'spec/fixtures/sample-as-zip.zip'
    expect { Creek::Book.new path, check_file_extension: false }
      .not_to raise_error
  end

  it 'Check file extension when requested.' do
    expect { Creek::Book.new 'spec/fixtures/invalid.xls', check_file_extension: true }
      .to raise_error 'Not a valid file format.'
  end

  it 'Check file extension of original_filename if passed.' do
    path = 'spec/fixtures/temp_string_io_file_path_with_no_extension'
    expect { Creek::Book.new path, :original_filename => 'invalid.xls' }
      .to raise_error 'Not a valid file format.'
    expect { Creek::Book.new path, :original_filename => 'valid.xlsx' }
      .not_to raise_error
  end
end

describe 'Creek parsing dates on a sample XLSX file' do
  before(:all) do
    @creek = Creek::Book.new 'spec/fixtures/sample_dates.xlsx'

    @expected_datetime_rows = [
      {'A3' => 'Date',              'B3' => Date.parse('2018-01-01')},
      {'A4' => 'Datetime 00:00:00', 'B4' => Time.parse('2018-01-01 00:00:00')},
      {'A5' => 'Datetime',          'B5' => Time.parse('2018-01-01 23:59:59')}]
  end

  after(:all) do
    @creek.close
  end

  it 'parses dates successfully' do
    rows = Array.new
    row_count = 0
    @creek.sheets[0].rows.each do |row|
      rows << row
      row_count += 1
    end

    (2..5).each do |number|
      expect(rows[number]).to eq(@expected_datetime_rows[number-2])
    end
  end
end

describe 'Creek parsing a file with large numbrts.' do
  before(:all) do
    @creek = Creek::Book.new 'spec/fixtures/large_numbers.xlsx'
    @expected_simple_rows = [{"A"=>"7.83294732E8", "B"=>"783294732", "C"=>783294732.0}]
  end

  after(:all) do
    @creek.close
  end

  it 'Parse simple rows successfully.' do
    rows = Array.new
    row_count = 0
    @creek.sheets[0].simple_rows.each do |row|
      rows << row
      row_count += 1
    end
    expect(rows[0]).to eq(@expected_simple_rows[0])
  end
end

describe 'Creek parsing a sample XLSX file' do
  before(:all) do
    @creek = Creek::Book.new 'spec/fixtures/sample.xlsx'
    @expected_rows = [{'A1'=>'Content 1', 'B1'=>nil, 'C1'=>'Content 2', 'D1'=>nil, 'E1'=>'Content 3'},
                    {'A2'=>nil, 'B2'=>'Content 4', 'C2'=>nil, 'D2'=>'Content 5', 'E2'=>nil, 'F2'=>'Content 6'},
                    {},
                    {'A4'=>'Content 7', 'B4'=>'Content 8', 'C4'=>'Content 9', 'D4'=>'Content 10', 'E4'=>'Content 11', 'F4'=>'Content 12'},
                    {'A5'=>nil, 'B5'=>nil, 'C5'=>nil, 'D5'=>nil, 'E5'=>nil, 'F5'=>nil, 'G5'=>nil, 'H5'=>nil, 'I5'=>nil, 'J5'=>nil, 'K5'=>nil, 'L5'=>nil, 'M5'=>nil, 'N5'=>nil, 'O5'=>nil, 'P5'=>nil, 'Q5'=>nil, 'R5'=>nil, 'S5'=>nil, 'T5'=>nil, 'U5'=>nil, 'V5'=>nil, 'W5'=>nil, 'X5'=>nil, 'Y5'=>nil, 'Z5'=>'Z Content', 'AA5'=>nil, 'AB5'=>nil, 'AC5'=>nil, 'AD5'=>nil, 'AE5'=>nil, 'AF5'=>nil, 'AG5'=>nil, 'AH5'=>nil, 'AI5'=>nil, 'AJ5'=>nil, 'AK5'=>nil, 'AL5'=>nil, 'AM5'=>nil, 'AN5'=>nil, 'AO5'=>nil, 'AP5'=>nil, 'AQ5'=>nil, 'AR5'=>nil, 'AS5'=>nil, 'AT5'=>nil, 'AU5'=>nil, 'AV5'=>nil, 'AW5'=>nil, 'AX5'=>nil, 'AY5'=>nil, 'AZ5'=>'Content 13'},
                    {'A6'=>'1', 'B6'=>'2', 'C6'=>'3'}, {'A7'=>'Content 15', 'B7'=>'Content 16', 'C7'=>'Content 18', 'D7'=>'Content 19'},
                    {'A8'=>nil, 'B8'=>'Content 20', 'C8'=>nil, 'D8'=>nil, 'E8'=>nil, 'F8'=>'Content 21'},
                    {'A10' => 0.15, 'B10' => 0.15}]

    @expected_simple_rows = [{"A"=>"Content 1", "B"=>nil, "C"=>"Content 2", "D"=>nil, "E"=>"Content 3"},
                    {"A"=>nil, "B"=>"Content 4", "C"=>nil, "D"=>"Content 5", "E"=>nil, "F"=>"Content 6"},
                    {},
                    {"A"=>"Content 7", "B"=>"Content 8", "C"=>"Content 9", "D"=>"Content 10", "E"=>"Content 11", "F"=>"Content 12"},
                    {"A"=>nil, "B"=>nil, "C"=>nil, "D"=>nil, "E"=>nil, "F"=>nil, "G"=>nil, "H"=>nil, "I"=>nil, "J"=>nil, "K"=>nil, "L"=>nil, "M"=>nil, "N"=>nil, "O"=>nil, "P"=>nil, "Q"=>nil, "R"=>nil, "S"=>nil, "T"=>nil, "U"=>nil, "V"=>nil, "W"=>nil, "X"=>nil, "Y"=>nil, "Z"=>"Z Content", "AA"=>nil, "AB"=>nil, "AC"=>nil, "AD"=>nil, "AE"=>nil, "AF"=>nil, "AG"=>nil, "AH"=>nil, "AI"=>nil, "AJ"=>nil, "AK"=>nil, "AL"=>nil, "AM"=>nil, "AN"=>nil, "AO"=>nil, "AP"=>nil, "AQ"=>nil, "AR"=>nil, "AS"=>nil, "AT"=>nil, "AU"=>nil, "AV"=>nil, "AW"=>nil, "AX"=>nil, "AY"=>nil, "AZ"=>"Content 13"},
                    {"A"=>"1", "B"=>"2", "C"=>"3"},
                    {"A"=>"Content 15", "B"=>"Content 16", "C"=>"Content 18", "D"=>"Content 19"},
                    {"A"=>nil, "B"=>"Content 20", "C"=>nil, "D"=>nil, "E"=>nil, "F"=>"Content 21"},
                    {"A"=>0.15, "B"=>0.15}]
  end

  after(:all) do
    @creek.close
  end

  it 'open an XLSX file successfully.' do
    expect(@creek).not_to be_nil
  end

  it 'opens small remote files successfully', remote: true do
    url = 'https://file-examples.com/wp-content/uploads/2017/02/file_example_XLSX_10.xlsx'
    @creek = Creek::Book.new(url, remote: true)

    expect(@creek.sheets[0]).to be_a Creek::Sheet
  end

  it 'opens large remote files successfully', remote: true do
    url = 'http://www.house.leg.state.mn.us/comm/docs/BanaianZooExample.xlsx'
    @creek = Creek::Book.new(url, remote: true)

    expect(@creek.sheets[0]).to be_a Creek::Sheet
  end

  it 'find sheets successfully.' do
    expect(@creek.sheets.count).to eq(1)
    sheet = @creek.sheets.first
    expect(sheet.state).to eql nil
    expect(sheet.name).to eql 'Sheet1'
    expect(sheet.rid).to eql 'rId1'
  end

  it 'Parse simple rows successfully.' do
    rows = Array.new
    row_count = 0
    @creek.sheets[0].simple_rows.each do |row|
      rows << row
      row_count += 1
    end
    (0..8).each do |number|
      expect(rows[number]).to eq(@expected_simple_rows[number])
    end
    expect(row_count).to eq(9)
  end


  it 'Parse rows with empty cells successfully.' do
    rows = Array.new
    row_count = 0
    @creek.sheets[0].rows.each do |row|
      rows << row
      row_count += 1
    end

    (0..8).each do |number|
      expect(rows[number]).to eq(@expected_rows[number])
    end
    expect(row_count).to eq(9)
  end

  it 'Parse rows with empty cells and meta data successfully.' do
    rows = Array.new
    row_count = 0
    @creek.sheets[0].rows_with_meta_data.each do |row|
      rows << row
      row_count += 1
    end
    expect(rows.map{|r| r['cells']}).to eq(@expected_rows)
  end
end
