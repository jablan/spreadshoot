require 'builder'

class Spreadshoot
  attr_reader :xml, :worksheets

  # <workbook 
  # xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
  # xmlns:r="http:// schemas.openxmlformats.org /officeDocument/2006/relationships">
  # <sheets>
  # <sheet name="Sheet1" sheetId="1" r:id="rId1" />
  # </sheets>
  # </workbook>
  def initialize name, options = {}, &block
    @worksheets = []
    yield(self)
    @builder = Builder::XmlMarkup.new
    @xml = @builder.workbook(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main", :"xmlns:r" => "http:// schemas.openxmlformats.org /officeDocument/2006/relationships") do |wb|
      wb.sheets do |sheets|
        @worksheets.each_with_index do |ws, i|
          sheets.sheet(:name => ws.title, :sheetId => i+1, :"r:id" => "rId#{i+1}")
        end
      end
    end
  end

  # <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  # <si>
  # <t>Region</t>
  # </si>
  # <si>
  # <t>Sales Person</t>
  # </si>
  # <si>
  # <t>Sales Quota</t>
  # </si>
  # ...12 more items ...
  # </sst>
  def ss_index string
    @ss ||= {}
    unless i = @ss[string]
      i = @ss.length
      @ss[string] = i
    end
    i
  end

  def worksheet title, options = {}, &block
    ws = Worksheet.new(self, title, options, &block)
    @worksheets << ws
  end


  def save filename

  end

  def dump
    puts @xml

    @worksheets.each do |ws|
      puts '==='
      puts ws.xml
    end

    puts '==='
    builder = Builder::XmlMarkup.new
    xml = builder.sst(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do |xsst|
      xsst.si do |xsi|
        @ss.keys.each do |str|
          xsi.t(str)
        end
      end
    end
    puts xml
  end

  class Worksheet
    # <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >
    # 	<sheetData>
    # 		<row>
    # 			<c>
    # 				<v>1234</v>
    # 			</c>
    # 		</row>
    # 	</sheetData>
    # </worksheet>

    attr_reader :title, :xml, :sheet_data, :spreadsheet
    def initialize spreadsheet, title, options = {}, &block
      @spreadsheet = spreadsheet
      @title = title

      @builder = Builder::XmlMarkup.new
      @xml = @builder.worksheet(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do |ws|
        ws.sheetData do |sd|
          @sheet_data = sd
          yield self
        end
      end
    end

    def row options = {}
      @sheet_data.row do |xr|

        row = Row.new(self, xr, options)
        yield(row)
      end
    end
  end

  class Row
    def initialize worksheet, xr, options = {}
      @worksheet = worksheet
      @spreadsheet = worksheet.spreadsheet
      @xr = xr
    end

    def cell value, options = {}
      case value
      when String
        i = @spreadsheet.ss_index(value)
        @xr.c(:t => 's'){ |xc| xc.v(i) }
      else
        @xr.c{ |xc| xc.v(value) }
      end
    end
  end

  class Cell

  end
end

