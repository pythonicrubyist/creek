[![version](https://badge.fury.io/rb/creek.svg)](https://badge.fury.io/rb/creek)
[![downloads](https://ruby-gem-downloads-badge.herokuapp.com/creek?type=total&total_label=downloads)](https://ruby-gem-downloads-badge.herokuapp.com/creek?type=total&total_label=downloads)

# Creek - Stream parser for large Excel (xlsx and xlsm) files.

Creek is a Ruby gem that provides a fast, simple and efficient method of parsing large Excel (xlsx and xlsm) files.


## Installation

Creek can be used from the command line or as part of a Ruby web framework. To install the gem using terminal, run the following command:

```
gem install creek
```

To use it in Rails, add this line to your Gemfile:

```ruby
gem 'creek'
```

## Basic Usage
Creek can simply parse an Excel file by looping through the rows enumerator:

```ruby
require 'creek'
creek = Creek::Book.new 'spec/fixtures/sample.xlsx'
sheet = creek.sheets[0]

sheet.rows.each do |row|
  puts row # => {"A1"=>"Content 1", "B1"=>nil, "C1"=>nil, "D1"=>"Content 3"}
end

sheet.simple_rows.each do |row|
  puts row # => {"A"=>"Content 1", "B"=>nil, "C"=>nil, "D"=>"Content 3"}
end

sheet.rows_with_meta_data.each do |row|
  puts row # => {"collapsed"=>"false", "customFormat"=>"false", "customHeight"=>"true", "hidden"=>"false", "ht"=>"12.1", "outlineLevel"=>"0", "r"=>"1", "cells"=>{"A1"=>"Content 1", "B1"=>nil, "C1"=>nil, "D1"=>"Content 3"}}
end

sheet.simple_rows_with_meta_data.each do |row|
  puts row # => {"collapsed"=>"false", "customFormat"=>"false", "customHeight"=>"true", "hidden"=>"false", "ht"=>"12.1", "outlineLevel"=>"0", "r"=>"1", "cells"=>{"A"=>"Content 1", "B"=>nil, "C"=>nil, "D"=>"Content 3"}}
end

sheet.state   # => 'visible'
sheet.name    # => 'Sheet1'
sheet.rid     # => 'rId2'
```

## Filename considerations
By default, Creek will ensure that the file extension is either *.xlsx or *.xlsm, but this check can be circumvented as needed:

```ruby
path = 'sample-as-zip.zip'
Creek::Book.new path, :check_file_extension => false
```

By default, the Rails [file_field_tag](http://api.rubyonrails.org/classes/ActionView/Helpers/FormTagHelper.html#method-i-file_field_tag) uploads to a temporary location and stores the original filename with the StringIO object. (See [this section](http://guides.rubyonrails.org/form_helpers.html#uploading-files) of the Rails Guides for more information.)

Creek can parse this directly without the need for file upload gems such as Carrierwave or Paperclip by passing the original filename as an option:

```ruby
# Import endpoint in Rails controller
def import
  file = params[:file]
  Creek::Book.new file.path, check_file_extension: false
end
```

## Parsing images
Creek does not parse images by default. If you want to parse the images,
use `with_images` method before iterating over rows to preload images information. If you don't call this method, Creek will not return images anywhere.

Cells with images will be an array of Pathname objects.
If an image is spread across multiple cells, same Pathname object will be returned for each cell.

```ruby
sheet.with_images.rows.each do |row|
  puts row # => {"A1"=>[#<Pathname:/var/folders/ck/l64nmm3d4k75pvxr03ndk1tm0000gn/T/creek__drawing20161101-53599-274q0vimage1.jpeg>], "B2"=>"Fluffy"}
end
```

Images for a specific cell can be obtained with images_at method:

```ruby
puts sheet.images_at('A1') # => [#<Pathname:/var/folders/ck/l64nmm3d4k75pvxr03ndk1tm0000gn/T/creek__drawing20161101-53599-274q0vimage1.jpeg>]

# no images in a cell
puts sheet.images_at('C1') # => nil
```

Creek will most likely return nil for a cell with images if there is no other text cell in that row - you can use *images_at* method for retrieving images in that cell.

## Remote files

```ruby
remote_url = 'http://dev-builds.libreoffice.org/tmp/test.xlsx'
Creek::Book.new remote_url, remote: true
```

## Mapping cells with header names
By default, Creek will map cell names with letter and number(A1, B3 and etc). To be able to get cell values by header column name use ***with_headers*** (can be used only with ***#simple_rows*** method!!!) during creation *(Note: header column is first string of sheet)*

```ruby
creek = Creek::Book.new file.path, with_headers: true
```


## Contributing

Contributions are welcomed. You can fork a repository, add your code changes to the forked branch, ensure all existing unit tests pass, create new unit tests which cover your new changes and finally create a pull request.

After forking and then cloning the repository locally, install the Bundler and then use it
to install the development gem dependencies:

```
gem install bundler
bundle install
```

Once this is complete, you should be able to run the test suite:

```
rake
```

There are some remote tests that are excluded by default. To run those, run

```
bundle exec rspec --tag remote
```

## Bug Reporting

Please use the [Issues](https://github.com/pythonicrubyist/creek/issues) page to report bugs or suggest new enhancements.


## License

Creek has been published under [MIT License](https://github.com/pythonicrubyist/creek/blob/master/LICENSE.txt)
