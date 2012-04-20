require 'builder'
require 'fileutils'

class Spreadshoot

  def initialize options = {}, &block
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
    FileUtils.rm_rf(dir)
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
    File.delete(filename) if File.exists?(filename)
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

    attr_reader :title, :xml, :spreadsheet, :row_index, :col_index, :cells
    def initialize spreadsheet, title, options = {}, &block
      @cells = {}
      @spreadsheet = spreadsheet
      @title = title
      @options = options
      @row_index = 0
      @col_index = 0
      # default table, if none defined
      @current_table = Table.new(self, options)

      yield self
    end

    def to_s
      @xml ||= Builder::XmlMarkup.new.worksheet(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do |ws|
        ws.sheetData do |sd|
          @cells.keys.sort.each do |row|
            sd.row(:r => row+1) do |xr|

              @cells[row].keys.sort.each do |col|
                cell = @cells[row][col]
                cell.output(xr)
              end
            end
          end
        end
      end
    end

    def row options = {}, &block
      @current_table.row options, &block
      @row_index += 1
    end

    def table options = {}
      @current_table = table = Table.new(self, options)
      yield table
      @row_index += table.row_index
      @col_index += table.col_index
      @current_table = Table.new(self, @options) # preparing one in case row directly called next
      table
    end
  end


  class Table
    attr_reader :worksheet, :direction
    attr_accessor :col_index, :row_index

    def initialize worksheet, options = {}
      @worksheet = worksheet
      @options = options
      @direction = options[:direction] || :vertical
      @row_index = 0
      @col_index = 0
      @row_topleft = options[:row_topleft] || @worksheet.row_index
      @col_topleft = options[:col_topleft] || @worksheet.col_index
    end

    def row options = {}
      row = Row.new(self, options)
      yield(row) if block_given?

      if @direction == :vertical
        @row_index += 1
        @col_index = 0
      else
        @col_index += 1
        @row_index = 0
      end
      row
    end

    def current_row
      @row_topleft + @row_index
    end

    def current_col
      @col_topleft + @col_index
    end

    # alphanumeric representation of coordinates
    def coords
      "#{Cell.alpha_index(current_col)}#{current_row+1}"
    end

  end

  class Row
    def initialize table, options = {}
      @table = table
    end

    def cell value = nil, options = {}
      cell = Cell.new(@table, value, options)
      @table.worksheet.cells[@table.current_row] ||= {}
      @table.worksheet.cells[@table.current_row][@table.current_col] = cell
      if @table.direction == :vertical
        @table.col_index += 1
      else
        @table.row_index += 1
      end
      cell
    end
  end

  class Cell
    # maps numeric column indices to letter based:
    # 1 -> 'A', 2 -> 'B', 27 -> 'AA' and so on
    def self.alpha_index i
      @alpha_indices ||= ('A'..'ZZ').to_a
      @alpha_indices[i]
    end

    def initialize table, value, options = {}
      @table = table
      @value = value
      @options = options
      @coords = @table.coords
    end

    # outputs the cell into the resulting xml
    def output xn_parent
      r = {:r => @coords}
      case @value
      when String
        i = @table.worksheet.spreadsheet.ss_index(@value)
        xn_parent.c(r.merge(:t => 's')){ |xc| xc.v(i) }
      when Hash # no @value, formula in options
        @options = @value
        xn_parent.c(r) do |xc|
          xc.f(@options[:formula])
        end
      when nil
        xn_parent.c(r)
      else
        xn_parent.c(r){ |xc| xc.v(@value) }
      end
    end

    def to_s
      @coords
    end

  end

end
