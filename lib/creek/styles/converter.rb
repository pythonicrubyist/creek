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
        when :float
          value.to_f
        when :percentage
          value.to_f / 100
        when :date, :time, :date_time
          convert_date(value, options)
        when :bignum
          convert_bignum(value)

        ## Nothing matched
        else
          value
        end
      end

      # the trickiest. note that  all these formats can vary on
      # whether they actually contain a date, time, or datetime.
      def self.convert_date(value, options)
        value                        = value.to_f
        days_since_date_system_start = value.to_i
        fraction_of_24               = value - days_since_date_system_start

        # http://stackoverflow.com/questions/10559767/how-to-convert-ms-excel-date-from-float-to-date-format-in-ruby
        date = options.fetch(:base_date, Date.new(1899, 12, 30)) + days_since_date_system_start

        if fraction_of_24 > 0 # there is a time associated
          seconds = (fraction_of_24 * 86400).round
          return Time.utc(date.year, date.month, date.day) + seconds
        else
          return date
        end
      end

      def self.convert_bignum(value)
        if defined?(BigDecimal)
          BigDecimal.new(value)
        else
          value.to_f
        end
      end
    end
  end
end
