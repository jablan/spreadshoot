require 'set'
require 'date'
require 'builder'
require 'fileutils'

class Spreadshoot

  # Create a new sheet, with given default formatting options
  def initialize options = {}, &block
    @worksheets = []
    @ss = {}
    @borders = {}
    @fonts = {}
    @styles = {}
    @components = Set.new
    yield(self)
  end

  # Gets the shared index of a given string
  def ss_index string
    @components << :ss
    unless i = @ss[string]
      i = @ss.length
      @ss[string] = i
    end
    i
  end

  # gets the shared index of a border set (one or more of :top, :bottom, :left, :right)
  def border borders
    borders = [borders] unless borders.is_a?(Array)
    borders.sort!
    unless i = @borders[borders]
      i = @borders.length
      @borders[borders] = i
    end
    i
  end

  def font options
    unless i = @fonts[options]
      i = @fonts.length
      @fonts[options] = i
    end
    i
  end

  # gets the shared index of a cell style (given by options hash)
  def style options = {}
    font = {}
    style = options.each_with_object({}) do |(option, value), acc|
      case option
      when :border
        acc[:border] = self.border(value)
      when :align
        acc[:align] = value
      when :bold, :italic, :font
        font[option] = value
      when :format
        acc[:format] = case value
                       when :date
                         14
                       when :percent
                         10
                       else
                         value
                       end
      end
    end
    style[:font] = self.font(font) unless font.empty?
    return nil if style.empty?

    unless i = @styles[style]
      i = @styles.length
      @styles[style] = i
    end
    i
  end

  # Create a new worksheet within a spreadsheet
  def worksheet title, options = {}, &block
    ws = Worksheet.new(self, title, options, &block)
    @worksheets << ws
  end

  # Saves the spreadsheet to an XLSX file with a given name
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
    @worksheets.each_with_index do |ws, i|
      File.open(File.join(dir, 'xl', 'worksheets', "sheet#{i+1}.xml"), 'w') do |f|
        f.write(ws)
      end
    end
    File.open(File.join(dir, 'xl', 'sharedStrings.xml'), 'w') do |f|
      f.write shared_strings
    end if @components.member?(:ss)
    File.open(File.join(dir, 'xl', 'styles.xml'), 'w') do |f|
      f.write styles
    end
    File.open(File.join(dir, 'xl', '_rels', 'workbook.xml.rels'), 'w') do |f|
      f.write xl_rels
    end

    filename = File.absolute_path(filename)
    FileUtils.chdir(dir)
    File.delete(filename) if File.exists?(filename)
    # zip the result
    puts `zip -r #{filename} ./`
#    FileUtils.rmdir_rf(dir)
  end

  # Dumps main XMLs to the stdout (for debugging purposes)
  def dump
    puts @xml

    @worksheets.each do |ws|
      puts '==='
      puts ws
    end

    puts '==='
    puts shared_strings

    puts '==='
    puts styles
  end

  private

  # Outputs final workbook XML
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
      xt.Override :PartName => "/xl/styles.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
      @worksheets.count.times do |i|
        xt.Override :PartName => "/xl/worksheets/sheet#{i+1}.xml", :ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
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
      rs.Relationship(:Id => "rId#{count+2}", :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", :Target => "styles.xml")
    end
  end

  def styles
    Builder::XmlMarkup.new.styleSheet(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do |xs|
      xs.fonts do |xf|
        xf.font
        @fonts.keys.each do |font|
          xf.font do |x|
            x.b(:val => 1) if font[:bold]
            x.i(:val => 1) if font[:italic]
            x.name(:val => font[:font]) if font[:font]
          end
        end
      end

      xs.fills do |xf|
        xf.fill do |x|
          x.patternFill(:patternType => 'none')
        end
      end

      xs.borders do |xbs|
        xbs.border do |xb|
          [:left, :right, :top, :bottom, :diagonal].each do |kind|
            xb.tag!(kind)
          end
        end
        @borders.keys.each do |border_set|
          xbs.border do |xb|
            [:left, :right, :top, :bottom, :diagonal].each do |kind|
              if border_set.member?(kind)
                xb.tag!(kind, :style => 'thin') do |x|
                  x.color(:rgb => 'FF000000')
                end
              else
                xb.tag!(kind)
              end
            end
          end
        end
      end

      xs.cellStyleXfs do |xcsx|
        xcsx.xf(:fillId => 0, :borderId => 0, :numFmtId => 0)
      end

      xs.cellXfs do |xcx|
        xcx.xf(:fillId => 0, :numFmtId => 0, :borderId => 0, :xfId => 0)
        @styles.keys.each do |style|
          align = style[:align]
          border = style[:border]
          font = style[:font]
          options = {:fillId => 0, :xfId => 0} # default
          options[:applyAlignment] = align ? 1 : 0
          options[:applyBorder] = border ? 1 : 0
          options[:borderId] = border ? border + 1 : 0
          options[:fontId] = font + 1 if font
          options[:applyFont] = font ? 1 : 0
          options[:numFmtId] = style[:format] ? style[:format] : 0
          options[:applyNumberFormat] = style[:format] ? 1 : 0
          xcx.xf(options) do |xf|
            if align = style[:align]
              xf.alignment(:horizontal => align)
            end
          end
        end
      end

#      xs.cellStyles do |xcs|
#        xcs.cellStyle(:name => "Normal", :xfId => "0", :builtinId => "0")
#      end
    end
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
      @column_widths = {}
      # default table, if none defined
      @current_table = Table.new(self, options)

      yield self
    end

    # outputs the worksheet as OOXML
    def to_s
      @xml ||= Builder::XmlMarkup.new.worksheet(:xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do |ws|
        unless @column_widths.empty?
          ws.cols do |xcols|
            @column_widths.keys.sort.each do |i|
              width = @column_widths[i]
              params = {:min => i+1, :max => i+1, :bestFit => 1}
              params.merge!({:customWidth => 1, :width => width}) if width
              xcols.col(params)
            end
          end
        end
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
      row = @current_table.row options, &block
      @row_index += 1
      row
    end

    def table options = {}
      @current_table = table = Table.new(self, options)
      yield table
      @row_index += table.row_max
      @col_index = 0
      @current_table = Table.new(self, @options) # preparing one in case row directly called next
      table
    end

    def set_col_width col, width
      @column_widths[col] = width
    end
  end # Worksheet


  # Allows you to group cells to a logical table within a worksheet. Makes putting several tables
  # to the same worksheet easier.
  class Table
    attr_reader :worksheet, :direction, :col_max, :row_max, :col_index, :row_index

    def initialize worksheet, options = {}
      @worksheet = worksheet
      @options = options
      @direction = options[:direction] || :vertical
      @row_index = 0
      @col_index = 0
      @row_max = 0
      @col_max = 0
      @row_topleft = options[:row_topleft] || @worksheet.row_index
      @col_topleft = options[:col_topleft] || @worksheet.col_index
    end

    def col_index= val
      @col_max = val if val > @col_max
      @col_index = val
    end

    def row_index= val
      @row_max = val if val > @row_max
      @row_index = val
    end

    def row options = {}
      row = Row.new(self, options)
      yield(row) if block_given?

      if @direction == :vertical
        self.row_index += 1
        self.col_index = 0
      else
        self.col_index += 1
        self.row_index = 0
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

  # A row of a table. The table could be horizontal oriented, or vertical oriented.
  class Row
    def initialize table, options = {}
      @table = table
      @options = options
    end

    def cell value = nil, options = {}
      cell = Cell.new(@table, value, @options.merge(options))
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

  # A cell within a row.
  class Cell
    # maps numeric column indices to letter based:
    # 0 -> 'A', 1 -> 'B', 26 -> 'AA' and so on
    def self.alpha_index i
      @alpha_indices ||= ('A'..'ZZ').to_a
      @alpha_indices[i]
    end

    def initialize table, value, options = {}
      @table = table
      @value = value
      @options = options
      @coords = @table.coords
      @table.worksheet.set_col_width(@table.current_col, @options[:width]) if @options.has_key?(:width)
      @options[:format] ||= :date if @value.is_a?(Date) || @value.is_a?(Time)
    end

    def current_col
      @table.current_col
    end

    def current_row
      @table.current_row
    end

    # outputs the cell into the resulting xml
    def output xn_parent
      r = {:r => @coords}
      if style = @table.worksheet.spreadsheet.style(@options)
        r.merge!(:s => style + 1)
      end
      case @value
      when String
        i = @table.worksheet.spreadsheet.ss_index(@value)
        xn_parent.c(r.merge(:t => 's')){ |xc| xc.v(i) }
      when Hash # no @value, formula in options
        @options = @value
        xn_parent.c(r) do |xc|
          xc.f(@options[:formula])
        end
      when Date, Time
        xn_parent.c(r){|xc| xc.v((@value - Date.new(1899,12,30)).to_i)}
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
