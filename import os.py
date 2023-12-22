import os
import openpyxl

values = list(range(1,11))

workbook = openpyxl.Workbook()

sheet = workbook.active

for i in values:
    cell = sheet.cell(row=1,column=i)
    cell.value = f'Column{i}'

workbook.save("C:/Users/m243553/Desktop/test 1.xlsx")

print("Excel file created successfuly")