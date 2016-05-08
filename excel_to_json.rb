require 'rubygems'
require 'rubyXL'
require 'colorize'
require 'json'
require 'fileutils'

$file_name = Time.new.strftime("%H:%M:%S")

current_user = `whoami`

Dir.chdir("/Users/#{current_user}/Desktop")

puts 'What is the path to the excel file (relative to Desktop)'.colorize :blue
excel_path = gets.chomp

$workbook = RubyXL::Parser.parse(excel_path)

puts 'Which tab?'.colorize :blue
tab = gets.chomp

puts 'Which column?'.colorize :blue
col = gets.chomp

puts 'Starting at which row?'.colorize :blue
row_start = gets.chomp

puts 'Ending at which row?'.colorize :blue
row_end = gets.chomp

$active_sheet = $workbook["#{tab}"]
$years = (1997..2014).to_a.map { |n| n.to_s }

def copy_excel_cells(col, row_start, row_end)
  start = row_start.to_i - 1
  last = row_end.to_i - 1
  current = start

  row_collection = []

  while current <= last
    row_data = {}

    row_data[:value] = $active_sheet[current][col].value
    row_data[:year] = $years.shift

    row_collection.push(row_data)

    current += 1
  end

  return row_collection
end

def convert_to_json(row_collection)
  row_collection.map { |obj| Hash[obj.each_pair.to_a] }.to_json
end

def save_to_file(json)
  File.write("output/#{$file_name}", json)
end

def convert_column_to_index(col)
  letters = ('a'..'z').to_a
  hash = {}

  26.times do |n|
    key = letters[n]
    hash[key] = n
  end

  return hash[ col.downcase ]
end

collection = copy_excel_cells(convert_column_to_index(col), row_start, row_end)
json = convert_to_json(collection)
save_to_file(json)
