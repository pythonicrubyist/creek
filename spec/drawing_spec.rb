require './spec/spec_helper'

describe 'drawing' do
  let(:book) { Creek::Book.new('spec/fixtures/sample-with-images.xlsx') }
  let(:book_no_images) { Creek::Book.new('spec/fixtures/sample.xlsx') }
  let(:book_with_one_cell_anchored_images) { Creek::Book.new('spec/fixtures/sample-with-one-cell-anchored-images.xlsx') }
  let(:drawingfile) { 'xl/drawings/drawing1.xml' }
  let(:drawing) { Creek::Drawing.new(book, drawingfile) }
  let(:drawing_without_images) { Creek::Drawing.new(book_no_images, drawingfile) }
  let(:drawing_with_one_cell_anchored_images) { Creek::Drawing.new(book_with_one_cell_anchored_images, drawingfile) }

  describe '#has_images?' do
    it 'has' do
      expect(drawing.has_images?).to eq(true)
    end

    it 'does not have' do
      expect(drawing_without_images.has_images?).to eq(false)
    end
  end

  describe '#images_at' do
    it 'returns images pathnames at cell' do
      image = drawing.images_at('A2')[0]
      expect(image.class).to eq(Pathname)
      expect(image.exist?).to eq(true)
      expect(image.to_path).to match(/.+creek__drawing.+\.jpeg$/)
    end

    context 'when no images in cell' do
      it 'returns nil' do
        images = drawing.images_at('B2')
        expect(images).to eq(nil)
      end
    end

    context 'when more images in one cell' do
      it 'returns all images at cell' do
        images = drawing.images_at('A10')
        expect(images.size).to eq(2)
        expect(images.all?(&:exist?)).to eq(true)
      end
    end

    context 'when same image across multiple cells' do
      it 'returns same image for each cell' do
        image1 = drawing.images_at('A4')[0]
        image2 = drawing.images_at('A5')[0]
        expect(image1.class).to eq(Pathname)
        expect(image1).to eq(image2)
      end
    end

    context 'when one cell anchored images in cell' do
      it 'returns image for anchored cell' do
        image = drawing_with_one_cell_anchored_images.images_at('A2')[0]
        expect(image.class).to eq(Pathname)
        expect(image.exist?).to eq(true)
      end

      it 'returns nil for non-anchored cell' do
        image = drawing_with_one_cell_anchored_images.images_at('A3')[0]
        # Image can be seen present on cell A4 in `sample-with-one-cell-anchored-images.xlsx`
        image_at_non_anchored_cell = drawing_with_one_cell_anchored_images.images_at('A4')
        expect(image.class).to eq(Pathname)
        expect(image_at_non_anchored_cell).to eq(nil)
      end
    end
  end
end
