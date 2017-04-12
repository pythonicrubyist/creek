require './spec/spec_helper'

describe 'sheet' do
  let(:book_with_images) { Creek::Book.new('spec/fixtures/sample-with-images.xlsx') }
  let(:book_no_images) { Creek::Book.new('spec/fixtures/sample.xlsx') }
  let(:sheetfile) { 'worksheets/sheet1.xml' }
  let(:sheet_with_images) { Creek::Sheet.new(book_with_images, 'Sheet 1', 1, '', '', '1', sheetfile) }
  let(:sheet_no_images) { Creek::Sheet.new(book_no_images, 'Sheet 1', 1, '', '', '1', sheetfile) }

  def load_cell(rows, cell_name)
    cell = rows.find { |row| !row[cell_name].nil? }
    cell[cell_name] if cell
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
end
