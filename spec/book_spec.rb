require './spec/spec_helper'

describe Creek::Book do
  let(:book) { Creek::Book.new('spec/fixtures/sample.xlsx') }
  let(:book_mac) { Creek::Book.new('spec/fixtures/mac.xlsx') }

  describe '#sheets' do
    let(:sheet_names) { ['Sheet X', 'Sheet Y'] }

    it { expect(book_mac.sheets).to be_an(Array) }
    it { expect(book_mac.sheets.map(&:name)).to match_array(sheet_names) }
  end

  describe '#style_types' do
    it { expect(book_mac.style_types).to be_an(Array) }
  end

  describe '#base_date' do
    it { expect(book_mac.base_date).to eq(described_class::DATE_1904) }
    it { expect(book.base_date).to eq(described_class::DATE_1900) }
  end

  describe '#document' do
    it { expect(book_mac.document).to be_a(Creek::Document) }
  end
end
