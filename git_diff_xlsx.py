# Converts a Microsoft Excel 2007+ file into plain text
# for comparison using git diff
#
# Instructions for setup:
# 1. Place this file in a folder
# 2. Add the following line to the global .gitconfig:
#     [diff "zip"]
#   	    binary = True
#	    textconv = python c:/path/to/git_diff_xlsx.py
# 3. Add the following line to the repository's .gitattributes	
#    *.xlsx diff=zip
# 4. Now, typing [git diff] at the prompt will produce text versions
# of Excel .xlsx files
#
# Copyright William Usher 2013
# Contact: w.usher@ucl.ac.uk
#

import xlrd as xl
import sys

def parse(file,outfile):    
    
    book = xl.open_workbook(file)

    num_sheets = book.nsheets

    print book.sheet_names()

#   print "File last edited by " + book.user_name + "\n"
    outfile.write("File last edited by " + book.user_name + "\n")
    
    def get_cells(sheet, rowx, colx):
        return sheet.cell_value(rowx, colx)
    
    # loop over worksheets

    for index in range(0,num_sheets):
        # find non empty cells
        sheet = book.sheet_by_index(index)
        outfile.write("=================================\n")
        outfile.write("Sheet: " + sheet.name + "[ " + str(sheet.nrows) + " , " + str(sheet.ncols) + " ]\n")
        outfile.write("=================================\n") 
        for row in range(0,sheet.nrows):
            for col in range(0,sheet.ncols):
                content = get_cells(sheet, row, col)
                if content <> "":
                    outfile.write("    " + str(xl.cellname(row,col)) + ": " + str(content) + "\n")
        print "\n"
        
# output cell address and contents of cell
def main():
    args = sys.argv[1:]
    if len(args) != 1:
        print 'usage: python git_diff_xlsx.py infile.xlsx'
        sys.exit(-1)
    outfile = sys.stdout
    parse(args[0],outfile)

if __name__ == '__main__':
    main()
