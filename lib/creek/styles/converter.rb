require 'set'

module Creek
  class Styles
    class Converter
      include Creek::Styles::Constants

      # Excel non-printable character escape sequence
      HEX_ESCAPE_REGEXP = /_x[0-9A-Fa-f]{4}_/

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
          unescape_string(value)
        when 'inlineStr'
          unescape_string(value)

        ##
        # Type can also be determined by a style,
        # detected earlier and cast here by its standardized symbol
        ##

        when :string
          value
        when :unsupported
          convert_unknown(value)
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
          convert_unknown(value)
        end
      end
      
      def self.convert_unknown(value)
        begin
          if value.nil? or value.empty?
            return value
          elsif value.to_i.to_s == value.to_s
            return value.to_i
          elsif value.to_f.to_s == value.to_s
            return value.to_f
          else
            return value
          end
        rescue
          return value
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
          BigDecimal(value)
        else
          value.to_f
        end
      end

      def self.unescape_string(value)
        # excel encodes some non-printable characters using a hex code in the format _xHHHH_
        # e.g. Carriage Return (\r) is encoded as _x000D_
        value.gsub(HEX_ESCAPE_REGEXP) { |match| match[2, 4].to_i(16).chr(Encoding::UTF_8) }
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
