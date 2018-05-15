require './spec/spec_helper'

describe Creek::Styles::StyleTypes do
  describe :call do
    it "return array of styletypes with mapping to ruby types" do
      doc = Creek::Document.new(nil, 'spec/fixtures/styles/first.xml')
      res = Creek::Styles::StyleTypes.new(doc).call
      expect(res.size).to eq(8)
      expect(res[3]).to eq(:date_time)
      expect(res).to eq([:unsupported, :unsupported, :unsupported, :date_time, :unsupported, :unsupported, :unsupported, :unsupported])
    end
  end
end
