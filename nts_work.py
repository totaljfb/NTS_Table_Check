# -------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Jason.Zhang
#
# Created:     30/11/2018
# Copyright:   (c) Jason.Zhang 2018
# Licence:     <your licence>
# -------------------------------------------------------------------------------
from openpyxl import load_workbook
from collections import Counter
from openpyxl.utils import get_column_letter


def main():
    pass


if __name__ == '__main__':
    main()


# change the file name to work on target file
wb = load_workbook("C:\\Users\\Jason.Zhang\\Dropbox\\NTS-\\2019\\Q1\\First Review\\"
                   + "table_01_64.xlsx", data_only=True)

sheet_list = wb.sheetnames

# sheet_list[0]: the first sheet you want to check
# sheet_list[1]: the second sheet you want to check, etc
sheet_name1 = sheet_list[0]
sheet_name2 = sheet_list[1]
sheet1 = wb.get_sheet_by_name(sheet_name1)
sheet2 = wb.get_sheet_by_name(sheet_name2)
print("Sheet name: " + sheet_name1)

# sometimes the max_row returns a large number, temporary solution is
# to manually set the row count here, can be improved in the future
max_row1 = sheet1.max_row
max_col1 = sheet1.max_column
max_row2 = sheet2.max_row
max_col2 = sheet2.max_column
# example for manually setting the row count
# max_row1 = 58

# check sheet font bolds
for i in range(1, max_row1):
    bold_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row=i, column=j)
        # if the cell is not empty, then get the property and append to list
        if cell.value:
            bold_list.append(str(cell.font.bold))
    # if the list is not empty, then print it
    if bold_list:
        print("Row " + str(i) + ", height: " + str(sheet1.row_dimensions[i].height)
              + " font bold information : " + Counter(bold_list).__str__())

# check sheet font names
print('\n')
print("Sheet name: " + sheet_name1)
for i in range(1, max_row1):
    font_name_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row=i, column=j)
        # if the cell is not empty, then get the property and append to list
        if cell.value:
            font_name_list.append(cell.font.name)
    # if the list is not empty, then print it
    if font_name_list:
        print("Row " + str(i) + ", font name information: "
        + Counter(font_name_list).__str__())

# check sheet font sizes
print('\n')
print("Sheet name: " + sheet_name1)
for i in range(1, max_row1):
    font_size_list = []
    for j in range(1, max_col1):
        cell = sheet1.cell(row=i, column=j)
        # if the cell is not empty, then get the property and append to list
        if cell.value:
            font_size_list.append(int(cell.font.size))
    # if the list is not empty, then print it
    if font_size_list:
        print("Row " + str(i) + ", font size information: " + Counter(font_size_list).__str__())

# check sheet highlight cells
print('\n')
print("Sheet name: " + sheet_name2)
incorrect_highlight_list = []
for i in range(1, max_row2 + 1):
    # don't check the last column because it should always be highlighted(supposed to be new column)
    for j in range(1, max_col2):
        cell = sheet2.cell(row=i, column=j)
        # FFFFFF00--yellow highlighted, 00000000--no highlight
        if (cell.fill.start_color.rgb == "FFFFFF00" and str(cell.value) == "True") or \
                (cell.fill.start_color.rgb == "00000000" and str(cell.value) == "False"):
            incorrect_highlight_cell = "[" + str(i) + "][" + get_column_letter(j) + "]"
            incorrect_highlight_list.append(incorrect_highlight_cell)
# if the list is not empty, then print it
if incorrect_highlight_list:
    print("Cells which may be incorrectly highlighted in working sheet: ")
    for item in incorrect_highlight_list:
        print(item)







