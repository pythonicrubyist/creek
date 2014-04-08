require 'creek'

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

end