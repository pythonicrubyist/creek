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

describe 'Creek parsing a sample XLSX file' do
  %w[file io stringio].each do |mode|
    context "With a #{mode}" do
      before(:all) do
        path = 'spec/fixtures/sample.xlsx'
        @creek = case mode
                 when 'file' then Creek::Book.new path
                 when 'io'
                   @io = File.open(path, 'r')
                   Creek::Book.new @io
                 when 'stringio'
                   Creek::Book.new StringIO.new(File.read(path))
                 end
        @expected_rows = [{'A1'=>'Content 1', 'B1'=>nil, 'C1'=>'Content 2', 'D1'=>nil, 'E1'=>'Content 3'},
                        {'A2'=>nil, 'B2'=>'Content 4', 'C2'=>nil, 'D2'=>'Content 5', 'E2'=>nil, 'F2'=>'Content 6'},
                        {},
                        {'A4'=>'Content 7', 'B4'=>'Content 8', 'C4'=>'Content 9', 'D4'=>'Content 10', 'E4'=>'Content 11', 'F4'=>'Content 12'},
                        {'A5'=>nil, 'B5'=>nil, 'C5'=>nil, 'D5'=>nil, 'E5'=>nil, 'F5'=>nil, 'G5'=>nil, 'H5'=>nil, 'I5'=>nil, 'J5'=>nil, 'K5'=>nil, 'L5'=>nil, 'M5'=>nil, 'N5'=>nil, 'O5'=>nil, 'P5'=>nil, 'Q5'=>nil, 'R5'=>nil, 'S5'=>nil, 'T5'=>nil, 'U5'=>nil, 'V5'=>nil, 'W5'=>nil, 'X5'=>nil, 'Y5'=>nil, 'Z5'=>'Z Content', 'AA5'=>nil, 'AB5'=>nil, 'AC5'=>nil, 'AD5'=>nil, 'AE5'=>nil, 'AF5'=>nil, 'AG5'=>nil, 'AH5'=>nil, 'AI5'=>nil, 'AJ5'=>nil, 'AK5'=>nil, 'AL5'=>nil, 'AM5'=>nil, 'AN5'=>nil, 'AO5'=>nil, 'AP5'=>nil, 'AQ5'=>nil, 'AR5'=>nil, 'AS5'=>nil, 'AT5'=>nil, 'AU5'=>nil, 'AV5'=>nil, 'AW5'=>nil, 'AX5'=>nil, 'AY5'=>nil, 'AZ5'=>'Content 13'},
                        {'A6'=>'1', 'B6'=>'2', 'C6'=>'3'}, {'A7'=>'Content 15', 'B7'=>'Content 16', 'C7'=>'Content 18', 'D7'=>'Content 19'},
                        {'A8'=>nil, 'B8'=>'Content 20', 'C8'=>nil, 'D8'=>nil, 'E8'=>nil, 'F8'=>'Content 21'},
                        {'A10' => 0.15, 'B10' => 0.15}]
      end

      after(:all) do
        @creek.close
        @io && @io.close
      end

      it 'open an XLSX file successfully.' do
        expect(@creek).not_to be_nil
      end

      it 'find sheets successfully.' do
        expect(@creek.sheets.count).to eq(1)
        sheet = @creek.sheets.first
        expect(sheet.state).to eql nil
        expect(sheet.name).to eql 'Sheet1'
        expect(sheet.rid).to eql 'rId1'
      end

      it 'Parse rows with empty cells successfully.' do
        rows = Array.new
        row_count = 0
        @creek.sheets[0].rows.each do |row|
          rows << row
          row_count += 1
        end

        expect(rows[0]).to eq(@expected_rows[0])
        expect(rows[1]).to eq(@expected_rows[1])
        expect(rows[2]).to eq(@expected_rows[2])
        expect(rows[3]).to eq(@expected_rows[3])
        expect(rows[4]).to eq(@expected_rows[4])
        expect(rows[5]).to eq(@expected_rows[5])
        expect(rows[6]).to eq(@expected_rows[6])
        expect(rows[7]).to eq(@expected_rows[7])
        expect(rows[8]).to eq(@expected_rows[8])
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
  end
end
