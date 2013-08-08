# Spreadshoot

Create Excel (OOXML) spreadsheets with ease, using a simple Ruby DSL.

## Installation

Add this line to your application's Gemfile:

    gem 'spreadshoot'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install spreadshoot

## Usage

The simplest example follows:

    # require us
    require 'spreadshoot'
    require 'date'

    spreadsheet = Spreadshoot.new do |s|
      s.worksheet('Simple') do |w| # multiple worksheet support
        w.row do |r|
          r.cell Date.today
          r.cell 'foo'
          @foo = r.cell 2 # we will remember this cell so we can use it later in the formula
        end
        w.row do |r|
          r.cell Time.now
          r.cell 'bar', :font => 'Times New Roman' # some formatting is supported, more to come
          @bar = r.cell 3
        end
        w.row # empty row
        w.row(:border => :top, :bold => true) do |r|
          r.cell('total', :align => :center)
          r.cell # empty cell
          r.cell :formula => "#{@foo} + #{@bar}" # yes, we support formulas
        end
      end
    end

    # save your resulting spreadsheet
    spreadsheet.save 'MySpreadsheet.xlsx'

See examples directory for some use cases.

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Added some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
