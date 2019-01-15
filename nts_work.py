#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Jason.Zhang
#
# Created:     30/11/2018
# Copyright:   (c) Jason.Zhang 2018
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def main():
    pass

if __name__ == '__main__':
    main()

import openpyxl
from openpyxl import load_workbook
from collections import Counter
#change the file name to work on target file
wb = load_workbook("C:\\Users\\Jason.Zhang\\Dropbox\\NTS-\\2019\\Q1\\First Review\\"
                    +"table_01_48PG2-12&2-13&3-5.xlsx")
sheet_list = wb.sheetnames
#sheet_list[0]: the first sheet you want to check
#sheet_list[1]: the second sheet you want to check, etc
sheet1 = wb.get_sheet_by_name(sheet_list[0])
max_row1 = sheet1.max_row
max_col1 = sheet1.max_column

for i in range(1, max_row1):
    bold_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row = i, column =j)
        bold_list.append(str(cell.font.bold))
    print("Row " + str(i) + ", height: " + str(sheet1.row_dimensions[i].height)
    + " font bold information : " + Counter(bold_list).__str__())
print('\n')
for i in range(1, max_row1):
    font_name_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row = i, column =j)
        font_name_list.append(cell.font.name)
    print("Row " + str(i) + ", font name information: "
    + Counter(font_name_list).__str__())
print('\n')
for i in range(1, max_row1):
    font_size_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row = i, column =j)
        font_size_list.append(int(cell.font.size))
    print("Row " + str(i) + ", font size information: "
    + Counter(font_size_list).__str__())








