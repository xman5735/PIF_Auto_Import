import tkinter as tk
from tkinter import messagebox, filedialog
from tkcalendar import DateEntry
import os
import shutil
from wand.image import Image

# Create the GUI
root = tk.Tk()
root.title("PIT Auto Import")

# Initialize global variables
tif_file_path = ""
pdf_file_path = ""

# Function to handle the file selection
def select_file():
    global tif_file_path
    tif_file_path = filedialog.askopenfilename(defaultextension=".tif", filetypes=[("TIF Files", "*.tif")])
    file_label.config(text=f"Selected File: {os.path.basename(tif_file_path)}")

# Function to convert TIF to PDF
def convert_to_pdf():
    global pdf_file_path
    pdf_file_path = os.path.splitext(tif_file_path)[0] + ".pdf"
    with Image(filename=tif_file_path) as img:
        img.save(filename=pdf_file_path)
    file_label.config(text=f"Converted File: {os.path.basename(pdf_file_path)}")

# Function to handle the submission
def submit():
    convert_to_pdf()
    start_date = start_date_entry.get_date().strftime("%Y-%m-%d")
    end_date = end_date_entry.get_date().strftime("%Y-%m-%d")
    new_file_name = f"PIT_{start_date}-{end_date}.pdf"

    #new_file_path = os.path.join(r"\\lcc-fsqb-01.lcc.local\Shares\Green Fox\QC\Logs\PIT_Scans", new_file_name)
    # Set the path to the directory containing the Excel files
    dir_path = r"C:\Users\alex\Desktop\Files\Local Shared Folder\QC_QC_QC\Logs\PIT_Scans"

    # Check if the directory exists, or try two other network locations if not
    if os.path.isdir(dir_path):
        new_file_path = os.path.join(dir_path, new_file_name)
        print(dir_path)
    else:
        network_path_1 = r"S:\Green Fox\QC_QC_QC\Logs\PIT_Scans"
        network_path_2 = r"Z:\QC_QC_QC\Logs\PIT_Scans"
        
        if os.path.isdir(network_path_1):
            dir_path = network_path_1
            new_file_path = os.path.join(dir_path, new_file_name)
            print(dir_path)
        elif os.path.isdir(network_path_2):
            dir_path = network_path_2
            new_file_path = os.path.join(dir_path, new_file_name)
            print(dir_path)
        else:
            print("Directory not found in any of the specified locations.")



    shutil.copy(pdf_file_path, new_file_path)
    messagebox.showinfo("Success", "File copied successfully.")
    #update_excel_files(start_date, end_date, new_file_path)

# TIF File Selection
tif_label = tk.Label(root, text="Select a TIF File:")
tif_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

tif_button = tk.Button(root, text="Browse", command=select_file)
tif_button.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

#convert_button = tk.Button(root, text="Convert to PDF", command=convert_to_pdf)
#convert_button.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

file_label = tk.Label(root, text="Selected File: None")
file_label.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W)

# Date Entries
start_date_label = tk.Label(root, text="Start Date:")
start_date_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)

start_date_entry = DateEntry(root, date_pattern="yyyy-mm-dd")
start_date_entry.grid(row=2, column=1, padx=5, pady=5)

end_date_label = tk.Label(root, text="End Date:")
end_date_label.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

end_date_entry = DateEntry(root, date_pattern="yyyy-mm-dd")
end_date_entry.grid(row=3, column=1, padx=5, pady=5)

# Submit Button
submit_button = tk.Button(root, text="Submit", command=submit)
submit_button.grid(row=4, column=0, padx=5, pady=5)

root.mainloop()
