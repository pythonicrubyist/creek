require './spec/spec_helper'

describe Creek::Document do
  let(:path) { 'xl/workbook.xml' }
  let(:book) { Creek::Book.new('spec/fixtures/sample-with-images.xlsx') }
  let(:document) { described_class.new(book, path) }

  describe '#xml' do
    it { expect(document.xml).to be_a(Nokogiri::XML::Document) }
  end

  describe '#relationships' do
    it { expect(document.relationships).to be_an(Nokogiri::XML::NodeSet) }
    it { expect(document.relationships.map(&:name)).to eq(['Relationships']) }
  end

  describe '#rels_filepath' do
    it { expect(document.rels_filepath).to be_an(String) }
    it { expect(document.rels_filepath).to eq('xl/_rels/workbook.xml.rels') }
  end

  describe '#namespace_prefix' do
    it { expect(document.namespace_prefix).to be_a(NilClass) }

    context 'when root namespace has prefix' do
      let(:path) { 'xl/drawings/drawing1.xml' }

      it { expect(document.namespace_prefix).to be_a(String) }
      it { expect(document.namespace_prefix).to eq('xdr') }
    end
  end

  describe '#css' do
    it { expect(document.css(['unknown'])).to be_an(Nokogiri::XML::NodeSet) }
    it { expect(document.css(['unknown']).map(&:name)).to eq([]) }

    it 'returns requested data' do
      expect(document.css(['sheet']).map { |el| el['name'] }).to eq(['Sheet1'])
    end
  end

  describe '#css_selector' do
    let(:css_path) { %w[a b c d] }

    it { expect(document.css_selector(css_path)).to be_an(String) }
    it { expect(document.css_selector(css_path)).to eq('a b c d') }

    context 'when root namespace has prefix' do
      let(:path) { 'xl/drawings/drawing1.xml' }
      let(:expected_selector) { 'xdr|a xdr|b xdr|c xdr|d' }

      it { expect(document.css_selector(css_path)).to eq(expected_selector) }
    end
  end
end
