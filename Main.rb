require_relative 'xlsxClass'
require_relative 'xlsClass'
# require "readline"

xlsx = Xlsx.new('testFile1.xlsx')
# xlsx = Xlsx.new('testFile2.xlsx')
xlsx.nilRowKiller

xls = Xls.new('test.xls')
# xls = Xls.new('test2.xls')
xls.nilRowKiller

input = ""

while input != "exit"
    input = gets.chomp

    case input
    when "x table"
        p xlsx.table
    when "s table"
        p xls.table
    
    when "x row element"
         p xlsx.table[0][1]
         p xlsx.table[0][2]
    when "s row element"
         p xls.table[0][1]
         p xls.table[0][2]

    when "x row"
        p xlsx.row(0)
        p xlsx.row(2)
        p xlsx.row(0)[2]
        p xlsx.row(2)[2]
    when "s row"
        p xls.row(0)
        p xls.row(2)
        p xls.row(0)[2]
        p xls.row(2)[2]

    when "x each"
        xlsx.each do |cell|
            p cell
        end
    when "s each"
        xls.each do |cell|
            p cell
        end

    when "x row syntax"
        p xlsx.column_table["prva"]
        p xlsx.column_table["prva"][0]
    when "s row syntax"
        p xls.column_table["prva"]
        p xls.column_table["prva"][0]

    when "x copy table"
        xlsx.copyTable('testFile2.xlsx',1)
    when "x remove table"
        xlsx.removeTable('testFile2.xlsx', 1)

    when "s copy table"
        xls.copyTable("test2.xls",1)
    when "s remove table"
        xls.removeTable("test2.xls",1)

    else
        p "wrong input command"
    end
end

p "goodbye dear user"