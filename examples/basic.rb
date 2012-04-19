require 'spreadshoot'

spreadsheet = Spreadshoot.new do |s|
  s.worksheet('Foobar') do |w|
    w.row do |r|
      r.cell 'foo'
      @foo = r.cell 2
    end
    w.row do |r|
      r.cell 'bar'
      @bar = r.cell 3
    end
    w.row # empty one
    w.row(:line => :above, :bold => true) do |r|
      r.cell 'total'
      r.cell # empty cell
      r.cell :formula => "#{@foo} + #{@bar}"
    end
  end
end

spreadsheet.dump
spreadsheet.save(ARGV[0])
