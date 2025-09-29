import os
import traceback
import pandas as pd
from tkinter import Frame, messagebox, ttk, filedialog

from utils.file_ops import upload_file, save_file

class PeriodeSAP(Frame):
    def __init__(self, master, menu_bar):
        super().__init__(master, highlightbackground="gray24")

        self.file_name = "SAP_Periode_Import.txt"
        self.menu_bar = menu_bar

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._init_layout()    
        
    def _init_layout(self):
        # Initializes the layout of the PeriodeSAP frame
        label_container = Frame(self)
        label_container.grid(row=0, column=0, sticky="ew")
        label_container.grid_columnconfigure(0, weight=1)
        
        ttk.Label(label_container, text="1. Select SAP .txt file", anchor="center").grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="2. App will minus 1 from PeriodeJahr", anchor="center").grid(row=1, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="").grid(row=2, column=0) # Empty label for spacing
        ttk.Button(label_container, text="Select File and Run", command=self._run).grid(row=3, column=0, padx=5)

    def _run(self):
        try:
            if not self.menu_bar.output_dir_path :
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")

            df_periode = self._select_and_process_file()
            self._save_periode_file(df_periode)
            
        except Exception as e:
            tb = traceback.format_exc()
            messagebox.showerror("Error", f"An error occurred: \n{str(e)}\n\n{tb}")
    
    def _select_and_process_file(self):
        df = upload_file(file_type="text", reader="text")
        df["PERIODJAHR"] = df["PERIODJAHR"] - 1
        df["BETRAG"] = df["BETRAG"].astype(str).str.replace(".", ",", regex=False)
        return df
    
    def _save_periode_file(self, df):
        save_file(df=df,
                  file_name=self.file_name,
                  output_loc=self.menu_bar.output_dir_path,
                  message_info={"title": "File Saved", 
                                "message": f"The import file has been saved to '{self.file_name}' in the output directory."})