require 'roo'
require 'spreadsheet'
require 'rubyXL'
require 'rubyXL/convenience_methods'

require 'rubyXL/convenience_methods/cell'
require 'rubyXL/convenience_methods/color'
require 'rubyXL/convenience_methods/font'
require 'rubyXL/convenience_methods/workbook'
require 'rubyXL/convenience_methods/worksheet'

class XlsxFile
    attr_accessor :path, :file, :table, :table2, :column_table, :row, :tableBorders

    def initialize(path)#1
        @path = path
        @file = Roo::Excelx.new("#@path", {:expand_merged_ranges => true})#5
        @table =  nil
        @table2 = nil
        @column_table = nil
        @row = nil
        @tableBorders =  Array.new(3)#prvi red, prva kolona, poslednji red, poslednja kolona

        self.load_column_table
        self.load_table(nil, @file)
    end

    def load_table(bool, file)#2
        lastRow = nil
        lastColumn = nil
        firstRow = nil
        firstColumn = nil
        file.each_with_pagename do |name, sh|
            if sh.first_row != nil then        
                @sheet = sh
                
                lastRow = sh.last_row
                lastColumn = sh.last_column
                firstRow = sh.first_row
                firstColumn = sh.first_column

                @tableTmp =  Array.new(sh.last_row - sh.first_row + 1){Array.new(sh.last_column - sh.first_column + 1)}
                @row = Array.new(sh.last_row - sh.first_row + 1)

                rowCnt = 0
                colCnt = 0
                flag = 0
                row_to_remove = -1

                sh.first_row.upto(sh.last_row) do |row|

                    sh.first_column.upto(sh.last_column) do |column|
                        @tableTmp[rowCnt][colCnt] = sh.cell(row, column)

                        if (sh.formula(row, column).to_s.include? "SUBTOTAL") || (sh.formula(row, column).to_s.include? "SUM")#TOTAL #8
                            flag = 1
                        end

                        colCnt += 1
                    end

                    if flag == 1 then
                        row_to_remove = rowCnt
                        @tableTmp.delete_at(row_to_remove)
                        #lastRow -= 1 #ovo mozda da se vrati, ako nam je nebitan red sa total i subtotal
                    end

                    rowCnt += 1
                    colCnt = 0

                end             
            end
        end
        if bool == 1 
             @table2 = @tableTmp 
        else 
            @table = @tableTmp
            @tableBorders[0] = firstRow
            @tableBorders[1] = firstColumn
            @tableBorders[2] = lastRow
            @tableBorders[3] = lastColumn

        end       
    end

    def load_column_table #6
        @file.each_with_pagename do |name, sh|
            if sh.first_column != nil then
                @column_table = Hash[]
                @row = Array.new(sh.last_row - sh.first_row + 1)

                flag = 0
                col_name = ""

                sh.first_column.upto(sh.last_column) do |column|
                    col_to_add = Column.new

                    sh.first_row.upto(sh.last_row) do |row|

                        if row == sh.first_row then
                            col_name = sh.cell(row, column)
                            column_table[col_name] = nil
                        else
                            col_to_add << sh.cell(row, column)
                        end

                        if (sh.formula(row, column).to_s.include? "SUBTOTAL") || (sh.formula(row, column).to_s.include? "SUM")#8
                            flag = 1
                        end

                    end

                    column_table[col_name] = col_to_add
                end

                if flag == 1 then
                    column_table.each_value do |array|
                        array.pop
                    end
                end
            end
        end
    end    

    def row(nr)#3
        @row = table[nr]
    end

    def each(&block)#4
        @table.each(&block)
    end

    def copyTable(secondTablePath) #9
        book2 = Roo::Excelx.new(secondTablePath, {:expand_merged_ranges => true})
        load_table(1,book2)

        unless table[0] == table2[0]
            puts "tabele nisu iste"
            return
        end

        # p table
        # p table2

        workbook = RubyXL::Parser.parse(@path)#ne zaboravi write
        worksheet = workbook[1]
        width = table2[0].length
        hight =  table2.length

        for i in 1..hight-1 do#red
            for j in 0..width-1 do#kolona
                worksheet.add_cell(tableBorders[2] + i -1, tableBorders[1] + j -1, table2[i][j])
            end            
        end

        workbook.write(@path)        
    end

    def removeTable(secondTablePath)#10
        book2 = Roo::Excelx.new(secondTablePath, {:expand_merged_ranges => true})
        load_table(1,book2)

        unless table[0] == table2[0]
            puts "tabele nisu iste"
            return
        end

        p table
        p table2

        secondTableRows = []#redovi gde se nalazi tabela 2 u tabeli 1

        i = tableBorders[0]-1
        table.each do |row|
            i+=1
            if row == table2[1]
                secondTableRows << i
                #puts "nasli ga u #{i}"
            end
        end


        workbook = RubyXL::Parser.parse(@path)#ne zaboravi write
        worksheet = workbook[1]

        for j in 0..secondTableRows.length() do#kolona
            worksheet.delete_row(secondTableRows[0]-1)
        end            

        workbook.write(@path)  
    end


end

class Column < Array#6

    def sum
        sum = 0

        self.each do |el|
            if el != nil then
                sum += el.to_i
            end
        end

        sum
    end

end


xl = XlsxFile.new('testFile1.xlsx')


p xl.table
p xl.column_table

# p xl.row(1)

# xl.each do |cell|
#     p cell
# end

# p xl.column_table["prva"]
# p xl.column_table["prva"][3]

# xl.copyTable('testFile2.xlsx')
# xl.removeTable('testFile2.xlsx')

