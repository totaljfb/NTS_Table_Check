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
                    +"table_02_39.xlsx")

sheet_list = wb.sheetnames

#sheet_list[0]: the first sheet you want to check
#sheet_list[1]: the second sheet you want to check, etc
sheet_name = sheet_list[0]
sheet1 = wb.get_sheet_by_name(sheet_name)
print("Sheet name: " + sheet_name)

#sometimes the max_row returns a large number, temperary solution is
#to manually set the row count here, can be improved in the future
max_row1 = sheet1.max_row
max_col1 = sheet1.max_column
#example for manually setting the row count
#max_row1 = 58

for i in range(1, max_row1):
    bold_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row = i, column = j)
        #if the cell is not empty, then get the property and append to list
        if cell.value:
            bold_list.append(str(cell.font.bold))
    #if the list is not empty, then print it
    if bold_list:
        print("Row " + str(i) + ", height: "
        + str(sheet1.row_dimensions[i].height)
        + " font bold information : " + Counter(bold_list).__str__())
print('\n')
print("Sheet name: " + sheet_name)
for i in range(1, max_row1):
    font_name_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row = i, column =j)
        #if the cell is not empty, then get the property and append to list
        if cell.value:
            font_name_list.append(cell.font.name)
    #if the list is not empty, then print it
    if font_name_list:
        print("Row " + str(i) + ", font name information: "
        + Counter(font_name_list).__str__())
print('\n')
print("Sheet name: " + sheet_name)
for i in range(1, max_row1):
    font_size_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row = i, column =j)
        #if the cell is not empty, then get the property and append to list
        if cell.value:
            font_size_list.append(int(cell.font.size))
    #if the list is not empty, then print it
    if font_size_list:
        print("Row " + str(i) + ", font size information: "
        + Counter(font_size_list).__str__())








