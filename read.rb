require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet.open 'read_book.xls'

sheet1 = book.worksheet 0

puts "\n\nReading from Excel:"

puts "\n\nSimple loop:"
sheet1.each do |row|
  print row
end

puts "\n\nFirst (0th) row omission:"
sheet1.each 1 do |row|
  print row
end
print "\n"

puts "\nRow access:"
row = sheet1.row(2)
print row

puts "\n\nValue access:"
print row[0]



puts "\n\nHello world"