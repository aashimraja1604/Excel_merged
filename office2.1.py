import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

#Create a GUI window
root = tk.Tk()
root.withdraw() #Hide the main window

#Asking the user to select the main file
input_file_paths = filedialog.askopenfilenames(title="Select input Excel files", filetypes=[("Excel files", "*.xlsx")])

#Creating a list to store individual sheet Data Frames
sheet_data_list = []

#Iterate through selected files
for input_file_path in input_file_paths:
    xls = pd.ExcelFile(input_file_path)
    for sheet_name in xls.sheet_names:
        sheet_data = xls.parse(sheet_name)
        sheet_data_list.append(sheet_data)

#Concatenate all sheet DataFrames into a single DataFrame
all_data = pd.concat(sheet_data_list, ignore_index=True)

#Ask the user to specify the output Excel file path
output_file_path = filedialog.asksaveasfilename(title="Save merged data to Excel file", defaultextension=".xlsx")

#Save the merged data to an Excel file
all_data.to_excel(output_file_path, index=False)
print("Merged data saved to", output_file_path)

# Function to open the output directory
def open_output_directory():
    output_directory = os.path.dirname(output_file_path)
    os.system(f"start {output_directory}")

# Create a button to open the output directory
open_dir_button = tk.Button(root, text="Open Output Directory", command=open_output_directory)
open_dir_button.pack()

# Start the GUI event loop
root.mainloop()