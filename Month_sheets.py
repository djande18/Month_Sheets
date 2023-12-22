import tkinter as tk
from tkinter import ttk
import os
import shutil
import openpyxl
from datetime import datetime

input_month = input("What month is it?\n")

current_year = datetime.now().year

days = {
    "January": 31,
    "February": 28,
    "March": 31,
    "April":30,
    "May":31,
    "June":30,
    "July":31,
    "August": 31,
    "September":30,
    "October": 31,
    "November": 30,
    "December": 31
}

month_names = {
    "January": "Jan",
    "February": "Feb",
    "March": "March",
    "April": "April",
    "May": "May",
    "June":"June",
    "July":"July",
    "August": "Aug",
    "September": "Sept",
    "October": "Oct",
    "November": "Nov",
    "December": "Dec"
}

# make a copy of the original file

original_file = "Leadership Report.xlsx"
current_directory = os.getcwd()

new_file_name = f"{input_month} {current_year} "+ original_file
destination_file = os.path.join(current_directory,new_file_name)

file_directory = os.path.join(current_directory,original_file)
shutil.copy(original_file,destination_file)


workbook = openpyxl.load_workbook(destination_file)
sheet = workbook.active

for i in range(1,days[input_month]+1):
    sheet_name = f'{month_names[input_month]} {i}'
    new_sheet = workbook.copy_worksheet(workbook['Template'])
    new_sheet.title = sheet_name
    print(f"Creating sheet '{sheet_name}'")

    
    # date_cell = sheet.cell(row=1,column=1)

    # date_cell.value = f"Date: {sheet_name}"
workbook.save(destination_file)