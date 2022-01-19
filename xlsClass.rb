
require 'roo-xls'
require 'spreadsheet'

class Xls
  attr_accessor :path, :file, :table, :table2, :column_table, :row, :tableBorders

  def initialize(path)#1
      @path = path
      @file = Roo::Spreadsheet.open("#@path")
      @table =  nil
      @table2 = nil
      @column_table = nil
      @row = nil
      @tableBorders =  Array.new(3)#prvi red, prva kolona, poslednji red, poslednja kolona

     self.load_column_table
     self.load_table(2, @file)
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

                      # if (sh.formula(row, column).to_s.include? "SUBTOTAL") || (sh.formula(row, column).to_s.include? "SUM")#TOTAL #8
                      #     flag = 1
                      # end

                      colCnt += 1
                  end

                  # if flag == 1 then
                  #     row_to_remove = rowCnt
                  #     @tableTmp.delete_at(row_to_remove)
                  #     #lastRow -= 1 #ovo mozda da se vrati, ako nam je nebitan red sa total i subtotal
                  # end

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

  def load_column_table
    @file.each_with_pagename do |name, sh|
        if sh.first_column != nil then
            @column_table = Hash[]
            @row = Array.new(sh.last_row - sh.first_row + 1)

            rowCnt = 0
            col_name = ""

            sh.first_column.upto(sh.last_column) do |column|
                col_to_add = Column.new

                sh.first_row.upto(sh.last_row) do |row|

                    if rowCnt == 0 then
                        col_name = sh.cell(row, column)
                        column_table[col_name] = nil
                    else
                        col_to_add << sh.cell(row, column)
                    end

                    rowCnt += 1
                end

                column_table[col_name] = col_to_add
                rowCnt = 0
            end

            column_table.each_value do |array|
                array.pop
            end
        end
    end
  end

  def row(nr)
      @row = table[nr]
  end

  def each(&block)
      @table.each(&block)
  end

  def copyTable(secondTablePath, sheet_num) 
    book2 = Roo::Spreadsheet.open(secondTablePath)
    load_table(1,book2)

    unless table[0] == table2[0]
        puts "tabele nisu iste"
        return
    end

    # p table
    # p table2

    workbook = Spreadsheet.open @path
    worksheet = workbook.worksheets[sheet_num]
    width = table2[0].length
    height =  table2.length

    worksheet.row(tableBorders[2]+height+10).insert 1

    for i in 1..height-1 do#red
        for j in 0..width-1 do#kolona
            #worksheet.add_cell(tableBorders[2] + i -1, tableBorders[1] + j -1, table2[i][j])                       
			 worksheet.rows[tableBorders[2] + i -2][tableBorders[1] + j -1] = table2[i][j]
        end            
    end
	  workbook.write(@path)
  end

  def removeTable(secondTablePath, sheet_num) 
    book2 = Roo::Spreadsheet.open(secondTablePath)
    load_table(1,book2)

    unless table[0] == table2[0]
        puts "tabele nisu iste"
        return
    end

    # p table
    # p table2

    secondTableRows = []

    i = tableBorders[0]-1
    j = 1
        table.each do |row|
            i+=1
            if row == table2[j]
                secondTableRows << i
                j +=1
            end        
        end

    workbook = Spreadsheet.open @path
    worksheet = workbook.worksheets[sheet_num]

    for j in 0..secondTableRows.length()-1 do#kolona
        worksheet.row(secondTableRows[j]-1).replace [""]
    end    

    workbook.write(@path)

  end
 
  def nilRowKiller
    nilCounter = 0
    rowCnt = -1

    table.each do |row|#trazimo red koji ima nil u sebi
        rowCnt += 1;
        #puts row
        if row.include? nil
          #  puts "nasli nil u redu #{rowCnt}"
            row.each do |cell|#kada nadjemo nil prodjemo i proverimo da li je ceo red nil
                if cell == nil 
                     nilCounter += 1    
                end              
            end
        end

        #puts "counter #{nilCounter} length #{table[0].length}"

        if nilCounter ==  table[0].length#ceo red je nil, znaci da je ukrasni red i da treba da se skloni
            table.delete_at(rowCnt)

            column_table.each_value do |array| #nakon sto obrisemo iz matrice, obrisemo i iz column_table
                array.delete_at(rowCnt-1)
            end

        end            
        nilCounter = 0
    end
    
end

end

class Column < Array

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

#  x = Xls.new('test.xls')


#x.copyTable
#x.removeTable

# p x.table
# p x.column_table
# puts"--------------"
# x.nilRowKiller
# p x.table
# p x.column_table


# p x.table[0][1]

# p x.row(0)[0]

# x.each do |cell|
#     p cell
# end

# p x.column_table["druga"]
# p x.column_table["header1"][0]

# p x.header1
# p x.header1[0]
# p x.header1.sum