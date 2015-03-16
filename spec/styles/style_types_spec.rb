require './spec/spec_helper'

describe Creek::Styles::StyleTypes do

  describe :call do
    it "return array of styletypes with mapping to ruby types" do
      xml_file = File.open('spec/fixtures/styles/first.xml')
      doc      = Nokogiri::XML(xml_file)
      res      = Creek::Styles::StyleTypes.new(doc).call
      res.size.should == 8
      res[3].should == :date_time
      res.should == [:unsupported, :unsupported, :unsupported, :date_time, :unsupported, :unsupported, :unsupported, :unsupported]
    end
  end
end
