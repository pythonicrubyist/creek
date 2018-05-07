require 'set'

module Creek
  class Styles
    class Converter
      include Creek::Styles::Constants
      ##
      # The heart of typecasting. The ruby type is determined either explicitly
      # from the cell xml or implicitly from the cell style, and this
      # method expects that work to have been done already. This, then,
      # takes the type we determined it to be and casts the cell value
      # to that type.
      #
      # types:
      # - s: shared string (see #shared_string)
      # - n: number (cast to a float)
      # - b: boolean
      # - str: string
      # - inlineStr: string
      # - ruby symbol: for when type has been determined by style
      #
      # options:
      # - shared_strings: needed for 's' (shared string) type
      # - base_date: from what date to begin, see method #base_date

      DATE_TYPES = [:date, :time, :date_time].to_set
      def self.call(value, type, style, options = {})
        return nil if value.nil? || value.empty?

        # Sometimes the type is dictated by the style alone
        if type.nil? || (type == 'n' && DATE_TYPES.include?(style))
          type = style
        end

        case type

        ##
        # There are few built-in types
        ##

        when 's' # shared string
          options[:shared_strings][value.to_i]
        when 'n' # number
          value.to_f
        when 'b'
          value.to_i == 1
        when 'str'
          value
        when 'inlineStr'
          value

        ##
        # Type can also be determined by a style,
        # detected earlier and cast here by its standardized symbol
        ##

        when :string, :unsupported
          value
        when :fixnum
          value.to_i
        when :float, :percentage
          value.to_f
        when :date
          convert_date(value, options)
        when :time, :date_time
          convert_datetime(value, options)
        when :bignum
          convert_bignum(value)

        ## Nothing matched
        else
          value
        end
      end

      def self.convert_date(value, options)
        date = base_date(options) + value.to_i
        yyyy, mm, dd = date.strftime('%Y-%m-%d').split('-')

        ::Date.new(yyyy.to_i, mm.to_i, dd.to_i)
      end

      def self.convert_datetime(value, options)
        date = base_date(options) + value.to_f.round(6)

        round_datetime(date.strftime('%Y-%m-%d %H:%M:%S.%N'))
      end

      def self.convert_bignum(value)
        if defined?(BigDecimal)
          BigDecimal.new(value)
        else
          value.to_f
        end
      end

      private

        def self.base_date(options)
          options.fetch(:base_date, Date.new(1899, 12, 30))
        end

        def self.round_datetime(datetime_string)
          /(?<yyyy>\d+)-(?<mm>\d+)-(?<dd>\d+) (?<hh>\d+):(?<mi>\d+):(?<ss>\d+.\d+)/ =~ datetime_string

          ::Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0)
        end
    end
  end
end
