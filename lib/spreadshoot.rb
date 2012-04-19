require 'builder'
require 'fileutils'

class Spreadshoot
  attr_reader :xml, :worksheets

  def initialize name, options = {}, &block
    @worksheets = []
    @ss = {}
    @rels = {}
    yield(self)
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
    @rels[:ss] ||= []
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

  # <workbook 
  # xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
  # xmlns:r="http:// schemas.openxmlformats.org /officeDocument/2006/relationships">
  # <sheets>
  # <sheet name="Sheet1" sheetId="1" r:id="rId1" />
  # </sheets>
  # </workbook>
  def workbook
    Builder::XmlMarkup.new.workbook(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main", :"xmlns:r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships") do |wb|
      wb.sheets do |sheets|
        @worksheets.each_with_index do |ws, i|
          sheets.sheet(:name => ws.title, :sheetId => i+1, :"r:id" => "rId#{i+1}")
        end
      end
    end
  end

  # <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  # <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  # <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
  # <Override PartName="/customXml/itemProps2.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
  # <Override PartName="/customXml/itemProps3.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
  # <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  # <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  # <Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
  # <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  # <Default Extension="xml" ContentType="application/xml"/>
  # <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  # <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  # <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  # <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  # <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  # <Override PartName="/xl/comments2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  # <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
  # <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  # <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  # <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  # <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  # <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  # <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  # </Types>
  def content_types
    Builder::XmlMarkup.new.Types(:xmlns => "http://schemas.openxmlformats.org/package/2006/content-types") do |xt|
      xt.Default(:Extension => "bin", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings")
      xt.Default(:Extension => "rels", :ContentType => "application/vnd.openxmlformats-package.relationships+xml")
      xt.Default(:Extension => "xml", :ContentType => "application/xml")
      xt.Default(:Extension => "vml", :ContentType => "application/vnd.openxmlformats-officedocument.vmlDrawing")
      xt.Override :PartName => "/xl/workbook.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
      xt.Override :PartName => "/xl/sharedStrings.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
      @worksheets.count.times do |i|
        xt.Override :PartName => "/xl/worksheets/sheet#{i+1}.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      end
    end
  end

  def shared_strings
    Builder::XmlMarkup.new.sst(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do |xsst|
      @ss.keys.each do |str|
        xsst.si do |xsi|
          xsi.t(str)
        end
      end
    end
  end

  # <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  # 	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  # </Relationships>
  def rels
    Builder::XmlMarkup.new.Relationships(:xmlns => "http://schemas.openxmlformats.org/package/2006/relationships") do |rs|
      rs.Relationship :Id => "rId1", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", :Target => "xl/workbook.xml"
    end
  end

  # <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  # <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  # <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  # <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  # <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  # <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  # </Relationships>
  def xl_rels
    Builder::XmlMarkup.new.Relationships(:xmlns => "http://schemas.openxmlformats.org/package/2006/relationships") do |rs|
      count = @worksheets.count
      count.times do |i|
        rs.Relationship :Id => "rId#{i+1}", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", :Target => "worksheets/sheet#{i+1}.xml"
      end
      rs.Relationship :Id => "rId#{count+1}", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", :Target => "sharedStrings.xml"
    end
  end

  def save filename
    dir = '/tmp/spreadshoot/'
    FileUtils.mkdir_p(dir)
    FileUtils.mkdir_p(File.join(dir, '_rels'))
    FileUtils.mkdir_p(File.join(dir, 'xl', 'worksheets'))
    FileUtils.mkdir_p(File.join(dir, 'xl', '_rels'))
    File.open(File.join(dir, '[Content_Types].xml'), 'w') do |f|
      f.write content_types
    end
    File.open(File.join(dir, '_rels', '.rels'), 'w') do |f|
      f.write rels
    end
    File.open(File.join(dir, 'xl', 'workbook.xml'), 'w') do |f|
      f.write workbook
    end
    File.open(File.join(dir, 'xl', 'sharedStrings.xml'), 'w') do |f|
      f.write shared_strings
    end
    File.open(File.join(dir, 'xl', '_rels', 'workbook.xml.rels'), 'w') do |f|
      f.write xl_rels
    end
    @worksheets.each_with_index do |ws, i|
      File.open(File.join(dir, 'xl', 'worksheets', "sheet#{i+1}.xml"), 'w') do |f|
        f.write(ws)
      end
    end

    filename = File.absolute_path(filename)
    FileUtils.chdir(dir)
    # zip the result
    puts `zip -r #{filename} ./`
#    FileUtils.rmdir_rf(dir)
  end

  def dump
    puts @xml

    @worksheets.each do |ws|
      puts '==='
      puts ws
    end

    puts '==='
    puts shared_strings
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

    def to_s
      @xml
    end

    def row options = {}
      @sheet_data.row do |xr|

        row = Row.new(self, xr, options)
        yield(row) if block_given?
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
