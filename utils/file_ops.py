from tkinter import filedialog, messagebox
import pandas as pd
import os

def upload_file(file_type, reader, header=None):
        file_set = {"excel": {"name": "Excel", "suffix": ".xlsx"}, 
                    "text": {"name":"text", "suffix": ".txt"}}
        file_path = filedialog.askopenfilename(title=f"Select {file_set[file_type]["name"]} File", 
                                               filetypes=[(f"{file_set[file_type]["suffix"]}", 
                                                           f"*{file_set[file_type]["suffix"]}")])
        
        if not file_path:
            return
        
        # latin-1 is the encoding when the exported from the system
        readers = {"excel": lambda file_path: pd.read_excel(file_path, header=header), 
                   "text": lambda file_path: pd.read_csv(file_path, sep="\t", decimal=",", encoding="latin-1", dtype=str),
                   "sheets": lambda file_path: pd.read_excel(file_path, sheet_name=["ACTUALS EUR", "ACTUALS CZK"], header=header)}
        
        file = readers[reader](file_path)

        messagebox.showinfo("Process Started", "Processing started. Please wait...")
        
        return file

def save_file(df, file_name, output_loc, message_info):    
    # Creates txt file which are uploaded into the system
    try:
        file_path = os.path.join(output_loc, file_name)
        df.to_csv(file_path, sep="\t", index=False, encoding="utf-8-sig")

        if os.path.exists(file_path):
            messagebox.showinfo(message_info["title"], message_info["message"])
    except Exception as e:
        messagebox.showerror("Save Failed", f"An error occurred while saving the file:\n{str(e)}")