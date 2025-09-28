from tkinter import filedialog, messagebox
import pandas as pd
import os

def upload_file(header):
        # Opens a file dialog to select an Excel file
        # Excel file to be uploaded is specifc for this application
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        file = pd.read_excel(file_path, header=header)
        messagebox.showinfo("Process Started", "Processing started. Please wait...")
        
        return file

def save_file(file, file_name, output_loc, messageInfo):    
    # Creates txt file which are uploaded into the system
    try:
        file_path = os.path.join(output_loc, file_name)
        file.to_csv(file_path, sep="\t", index=False, encoding="utf-8-sig")

        if os.path.exists(file_path):
            messagebox.showinfo(messageInfo["title"], messageInfo["message"])
    except Exception as e:
        messagebox.showerror("Save Failed", f"An error occurred while saving the file:\n{str(e)}")