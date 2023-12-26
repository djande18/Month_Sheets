import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import os
import shutil
import openpyxl
from datetime import datetime


#Building the UI

#This will handle the combobox
def on_combobox_change():
    selected_value.set("Selected: " + combobox.get())

#Runs the function below when a button is clicked
def on_submit():
    selected_month = combobox.get()
    destination_path = path_field.get()

    if not destination_path:
        messagebox.showerror("Error","Please enter a valid destination path.")

    else:
        generate_sheet(selected_month,destination_path)

        completion_label.config(text = "Spreadsheet Generated!")

#Original Script

def generate_sheet(input_month,destination_path):

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

    new_file_name = f"{input_month} {current_year} "+ original_file
    destination_file = os.path.join(destination_path,new_file_name)

    file_directory = os.path.join(destination_path,original_file)
    shutil.copy(original_file,destination_file)


    workbook = openpyxl.load_workbook(destination_file)
    sheet = workbook.active

    for i in range(1,days[input_month]+1):
        sheet_name = f'{month_names[input_month]} {i}'
        new_sheet = workbook.copy_worksheet(workbook['Template'])
        new_sheet.title = sheet_name
        print(f"Creating sheet '{sheet_name}'")

        
        date_cell = new_sheet.cell(row=1,column=1)
        date_cell.value = f"Date: {sheet_name}"

    workbook.save(destination_file)

#Create the UI

window = tk.Tk()
window.title("Monthly Spreadheet Generator")
window.geometry("550x175")
# label = tk.Label(window,text="Monthly Spreadheet Generator")
# label.pack()

#Dropdown options
options = ["January","February","March","April","May","June","July","August","September","October","November","December"]

#This is where the dropdown value will be stored
selected_value = tk.StringVar()

#Generate dropdown
combobox = ttk.Combobox(window, values = options,width=10)
combobox.set(options[0])
combobox.bind("<<ComboboxSelected>>", on_combobox_change)
combobox.grid(row=1,column=2,pady=20,sticky="w")

combo_label = tk.Label(window,text="Month: ")
combo_label.grid(row=1,column=1,pady=10,padx=20,sticky="e")

path_field = tk.Entry(window,width=50)
button = tk.Button(window,text="Generate",command = on_submit)

field_label = tk.Label(window,text="File Path:")
field_label.grid(row=2,column=1,pady=10,padx=20,sticky="e")

path_field.grid(row=2,column=2,padx=0,pady=20,sticky="w")
button.grid(row=2,column=3,pady=10,padx=30)

completion_label = tk.Label(window,text="")
completion_label.grid(row=3,column=2,sticky="w")

window.mainloop()
