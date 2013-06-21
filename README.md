## Creek -- Stream parser for large Excel(xlsx and xlsm) files.

Inspired by [Dullard](https://github.com/thirtyseven/dullard), Creek is a fast, simple and efficient way of parsing xlsx and xlsm files using SAX parsing.


## Installation

Creek can be used from the command line or as part of a Ruby web framework. To install the gem using terminal, run the following command:

    gem install creek

To use it in Rails, add this line to your Gemfile:

    gem "creek"


## Basic Usage
Creek can simply parse an Excel file by looping through the rows enumerator:
```ruby
    require 'creek'
    creek = Creek::Book.new "specs/fixtures/sample.xlsx"
    creek.sheets[0].rows.each do |row|
      puts row.inspect
      # => {"A1"=>"Content 1", "B1"=>nil, "C1"=>"Content 2", "D1"=>nil, "E1"=>"Content 3"}
    end
```


## Contributing

Contributions are welcomed. You can fork a repository, add your code changes to the forked branch, ensure all existing unit tests pass, create new unit tests cover your new changes and finally create a pull request.

After forking and then cloning the repository locally, install Bundler and then use it
to install the development gem dependecies:

    gem install bundler
    bundle install

Once this is complete, you should be able to run the test suite:

    rake


## Bug Reporting/Feature Request

Please use the https://github.com/pythonicrubyist/creek/issues page to report bugs or suggest new enhancements.


## License

Creek has been published under [MIT License](https://github.com/pythonicrubyist/creek/blob/master/LICENSE.txt)