require './spec/spec_helper'

describe 'shared strings' do

  it 'parses rich text strings correctly' do
    shared_strings_xml_file = File.open('spec/fixtures/sst.xml')
    doc = Nokogiri::XML(shared_strings_xml_file)
    dictionary = Creek::SharedStrings.parse_shared_string_from_document(doc)

    expect(dictionary.keys.size).to eq(7)
    expect(dictionary[0]).to eq('Cell A1')
    expect(dictionary[1]).to eq('Cell B1')
    expect(dictionary[2]).to eq('My Cell')
    expect(dictionary[3]).to eq('Cell A2')
    expect(dictionary[4]).to eq('Cell B2')
    expect(dictionary[5]).to eq("Cell with\rescaped\rcharacters")
    expect(dictionary[6]).to eq('吉田兼好')
  end

end