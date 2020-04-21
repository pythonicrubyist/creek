# frozen_string_literal: true

require './spec/spec_helper'

describe 'sheet' do
  let(:book_with_images) { Creek::Book.new('spec/fixtures/sample-with-images.xlsx') }
  let(:sheetfile) { 'worksheets/sheet1.xml' }
  let(:sheet_with_images) { Creek::Sheet.new(book_with_images, 'Sheet 1', 1, '', '', '1', sheetfile) }

  def load_cell(rows, cell_name)
    cell = rows.find { |row| row[cell_name] }
    cell[cell_name] if cell
  end

  context 'escaped ampersand' do
    let(:book_escaped) { Creek::Book.new('spec/fixtures/escaped.xlsx') }
    it 'does NOT escape ampersand' do
      expect(book_escaped.sheets[0].rows.to_enum.map(&:values)).to eq([%w[abc def], %w[ghi j&k]])
    end

    let(:book_escaped2) { Creek::Book.new('spec/fixtures/escaped2.xlsx') }
    it 'does escape ampersand' do
      expect(book_escaped2.sheets[0].rows.to_enum.map(&:values)).to eq([%w[abc def], %w[ghi j&k]])
    end
  end

  describe '#rows' do
    context 'with excel with images' do
      context 'with images preloading' do
        let(:rows) { sheet_with_images.with_images.rows.map { |r| r } }

        it 'parses single image in a cell' do
          expect(load_cell(rows, 'A2').size).to eq(1)
        end

        it 'returns nil for cells without images' do
          expect(load_cell(rows, 'A3')).to eq(nil)
          expect(load_cell(rows, 'A7')).to eq(nil)
          expect(load_cell(rows, 'A9')).to eq(nil)
        end

        it 'returns nil for merged cell within empty row' do
          expect(load_cell(rows, 'A5')).to eq(nil)
        end

        it 'returns nil for image in a cell with empty row' do
          expect(load_cell(rows, 'A8')).to eq(nil)
        end

        it 'returns images for merged cells' do
          expect(load_cell(rows, 'A4').size).to eq(1)
          expect(load_cell(rows, 'A6').size).to eq(1)
        end

        it 'returns multiple images' do
          expect(load_cell(rows, 'A10').size).to eq(2)
        end
      end

      it 'ignores images' do
        rows = sheet_with_images.rows.map { |r| r }
        expect(load_cell(rows, 'A2')).to eq(nil)
        expect(load_cell(rows, 'A3')).to eq(nil)
        expect(load_cell(rows, 'A4')).to eq(nil)
      end
    end

    context 'with excel without images' do
      let(:book_no_images) { Creek::Book.new('spec/fixtures/sample.xlsx') }
      let(:sheet_no_images) { Creek::Sheet.new(book_no_images, 'Sheet 1', 1, '', '', '1', sheetfile) }

      it 'does not break on with_images' do
        rows = sheet_no_images.with_images.rows.map { |r| r }
        expect(load_cell(rows, 'A10')).to eq(0.15)
      end
    end
  end

  describe '#images_at' do
    it 'returns images for merged cell' do
      image = sheet_with_images.with_images.images_at('A5')[0]
      expect(image.class).to eq(Pathname)
    end

    it 'returns images for empty row' do
      image = sheet_with_images.with_images.images_at('A8')[0]
      expect(image.class).to eq(Pathname)
    end

    it 'returns nil for empty cell' do
      image = sheet_with_images.with_images.images_at('B3')
      expect(image).to eq(nil)
    end

    it 'returns nil for empty cell without preloading images' do
      image = sheet_with_images.images_at('B3')
      expect(image).to eq(nil)
    end
  end

  describe '#simple_rows' do
    let(:book_with_headers) { Creek::Book.new('spec/fixtures/sample-with-headers.xlsx') }
    let(:sheet) { Creek::Sheet.new(book_with_headers, 'Sheet 1', 1, '', '', '1', sheetfile) }

    subject { sheet.simple_rows.to_a[1] }

    it 'returns values by letters' do
      expect(subject['A']).to eq 'value1'
      expect(subject['B']).to eq 'value2'
    end

    context 'when enable with_headers property' do
      before { sheet.with_headers = true }

      subject { sheet.simple_rows.to_a[1] }

      it 'returns values by headers name' do
        expect(subject['HeaderA']).to eq 'value1'
        expect(subject['HeaderB']).to eq 'value2'
        expect(subject['HeaderC']).to eq 'value3'
      end
    end
  end
end
