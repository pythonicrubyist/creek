require './spec/spec_helper'

describe Creek::Styles::Converter do
  describe :call do
    def convert(value, type, style)
      Creek::Styles::Converter.call(value, type, style)
    end

    describe :date do
      it 'works' do
        expect(convert('41275', 'n', :date)).to eq(Date.new(2013, 0o1, 0o1))
      end
    end

    describe :date_time do
      it 'works' do
        expect(convert('41275', 'n', :date_time)).to eq(Time.new(2013, 0o1, 0o1))
      end
    end
  end
end
