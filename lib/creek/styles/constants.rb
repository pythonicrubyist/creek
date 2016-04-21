module Creek
  class Styles
    module Constants
      # Map of non-custom numFmtId to casting symbol
      NumFmtMap = {
        0  => :string,         # General
        1  => :fixnum,         # 0
        2  => :float,          # 0.00
        3  => :fixnum,         # #,##0
        4  => :float,          # #,##0.00
        5  => :unsupported,    # $#,##0_);($#,##0)
        6  => :unsupported,    # $#,##0_);[Red]($#,##0)
        7  => :unsupported,    # $#,##0.00_);($#,##0.00)
        8  => :unsupported,    # $#,##0.00_);[Red]($#,##0.00)
        9  => :percentage,     # 0%
        10 => :percentage,     # 0.00%
        11 => :bignum,         # 0.00E+00
        12 => :unsupported,    # # ?/?
        13 => :unsupported,    # # ??/??
        14 => :date,           # mm-dd-yy
        15 => :date,           # d-mmm-yy
        16 => :date,           # d-mmm
        17 => :date,           # mmm-yy
        18 => :time,           # h:mm AM/PM
        19 => :time,           # h:mm:ss AM/PM
        20 => :time,           # h:mm
        21 => :time,           # h:mm:ss
        22 => :date_time,      # m/d/yy h:mm
        37 => :unsupported,    # #,##0 ;(#,##0)
        38 => :unsupported,    # #,##0 ;[Red](#,##0)
        39 => :unsupported,    # #,##0.00;(#,##0.00)
        40 => :unsupported,    # #,##0.00;[Red](#,##0.00)
        45 => :time,           # mm:ss
        46 => :time,           # [h]:mm:ss
        47 => :time,           # mmss.0
        48 => :bignum,         # ##0.0E+0
        49 => :unsupported     # @
      }
    end
  end
end
