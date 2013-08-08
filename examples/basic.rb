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

  s.worksheet('Relative positioned tables') do |w|
    t1 = w.table do |t|
      3.times do
        t.row do |r|
          4.times do
            r.cell(1)
          end
        end
      end
    end
    w.table(:next_to => t1) do |t|
      5.times do
        t.row do |r|
          2.times do
            r.cell(2)
          end
        end
      end
    end
    t3 = w.table do |t|
      4.times do
        t.row do |r|
          5.times do
            r.cell(3)
          end
        end
      end
    end
    w.table(:next_to => t3) do |t|
      2.times do
        t.row do |r|
          2.times do
            r.cell(4)
          end
        end
      end
    end
  end

  s.worksheet('Cell formats') do |w|
    w.table do |t|
      t.row do |r|
        r.cell 'No formatting'
        r.cell 0.345
      end
      t.row do |r|
        r.cell 'Percent'
        r.cell 0.345, :format => :percent
      end
      t.row do |r|
        r.cell 'Rounded percent'
        r.cell 0.345, :format => :percent_rounded
      end
      t.row do |r|
        r.cell 'Currency'
        r.cell 0.345, :format => :currency
      end
      t.row 'Shorthand for', 'multiple cells'
    end
  end
end

spreadsheet.dump
spreadsheet.save(ARGV[0])
