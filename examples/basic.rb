require 'spreadshoot'

spreadsheet = Spreadshoot.new('title') do |s|
  s.worksheet('Foobar') do |w|
    w.row do |r|
      r.cell('foo')
      r.cell(2, :name => :foo)
    end
    w.row do |r|
      r.cell('bar')
      r.cell(3, :name => :bar)
    end
    w.row # empty one
    w.row(:line => :above, :bold => true) do |r|
      r.cell('total')
      r.cell(:formula => 'foo + bar')
    end
  end
end

spreadsheet.dump
spreadsheet.save(ARGV[0])
