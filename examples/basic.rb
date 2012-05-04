require 'spreadshoot'
require 'date'

spreadsheet = Spreadshoot.new do |s|
  s.worksheet('Simple') do |w|
    w.row do |r|
      r.cell Date.today
      r.cell 'foo'
      @foo = r.cell 2
    end
    w.row do |r|
      r.cell Time.now
      r.cell 'bar', :font => 'Times New Roman'
      @bar = r.cell 3
    end
    w.row # empty one
    w.row(:border => :top, :bold => true) do |r|
      r.cell('total', :align => :center)
      r.cell # empty cell
      r.cell :formula => "#{@foo} + #{@bar}"
    end
  end

  s.worksheet('Tables') do |w|
    w.table(:direction => :horizontal) do |t|
      t.row do |r|
        r.cell 'foo'
        @foo = r.cell 2
      end
      t.row do |r|
        r.cell 'bar'
        @bar = r.cell 3
      end
      t.row # empty one
      t.row(:line => :above, :bold => true) do |r|
        r.cell 'total'
        r.cell # empty cell
        @total1 = r.cell :formula => "#{@foo} + #{@bar}"
      end
    end
    w.row # empty row
    w.table do |t| # another table
      t.row do |r|
        r.cell 'foo'
        @foo2 = r.cell 6
      end
      t.row do |r|
        r.cell 'bar'
        @bar2 = r.cell 7
      end
      t.row # empty one
      t.row(:line => :above, :bold => true) do |r|
        r.cell 'total'
        r.cell # empty cell
        @total2 = r.cell :formula => "#{@foo2} + #{@bar2}"
      end
    end
    w.row do |r|
      r.cell 'Grand total:'
      r.cell
      r.cell :formula => "#{@total1} + #{@total2}"
    end
  end

  s.worksheet('Inverted Tables') do |w|
    w.table(:direction => :horizontal) do |t|
      t.row do |r|
        r.cell 'foo'
        @foo = r.cell 2
      end
      t.row do |r|
        r.cell 'bar'
        @bar = r.cell 3
      end
      t.row # empty one
      t.row(:line => :above, :bold => true) do |r|
        r.cell 'total'
        r.cell # empty cell
        r.cell :formula => "#{@foo} + #{@bar}"
      end
    end
  end
end

spreadsheet.dump
spreadsheet.save(ARGV[0])
