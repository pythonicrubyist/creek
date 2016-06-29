require './spec/spec_helper'

describe 'shared strings' do

  it 'parses rich text strings correctly' do
    shared_strings_xml_file = File.open('spec/fixtures/sst.xml')
    doc = Nokogiri::XML(shared_strings_xml_file)
    dictionary = Creek::SharedStrings.parse_shared_string_from_document(doc)

    dictionary.keys.size.should == 5
    dictionary[0].should == 'Cell A1'
    dictionary[1].should == 'Cell B1'
    dictionary[2].should == 'My Cell'
    dictionary[3].should == 'Cell A2'
    dictionary[4].should == 'Cell B2'
  end

  it 'decodes _x000D_ to carriage return character (\r)' do
    shared_strings_xml_file = File.open('spec/fixtures/sst-x000D.xml')
    doc = Nokogiri::XML(shared_strings_xml_file)
    dictionary = Creek::SharedStrings.parse_shared_string_from_document(doc)

    dictionary.keys.size.should == 17
    dictionary.each do |index, text|
      text.should_not include '_x000D_'
    end

    dictionary[13].should include "\u000D"
    dictionary[14].should include "\u000D"
    dictionary[15].should include "\u000D"
  end

end