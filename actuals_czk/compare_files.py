from tkinter import Frame, ttk, filedialog, messagebox, StringVar
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd
import os


class CompareFiles(Frame):

    def __init__(self, controller, menu_bar):
        super().__init__(controller, highlightbackground="gray24")
        self.columnconfigure(0, weight=1)
        
        self.menu_bar = menu_bar
        # Variable to hold the column name used for grouping
        self.column_name = StringVar()
        
        self.file1_name = StringVar()
        self.file1_path = ""
        self.file2_name = StringVar()
        self.file2_path = ""
        
        self._init_layout()
        
    def _init_layout(self):
        ttk.Label(self, text="File 1 Name: ", anchor="center").grid(row=0, column=0, padx=5, pady=10, sticky="nesw")
        ttk.Label(self, textvariable=self.file1_name, anchor="center").grid(row=0, column=1, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Select file 1", command=lambda: self._upload_action("file_1")).grid(row=1, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Label(self, text="File 2 Name: ", anchor="center").grid(row=2, column=0, padx=5, pady=10, sticky="nesw")
        ttk.Label(self, textvariable=self.file2_name, anchor="center").grid(row=2, column=1, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Select file 2", command=lambda: self._upload_action("file_2")).grid(row=3, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Compare", command=self._run).grid(row=4, column=0, padx=5, pady=5, sticky="nesw")

    def _upload_action(self, file_key):
        if file_key == "file_1":
            path_attr = "file1_path"
            file_name = self.file1_name
        elif file_key == "file_2":
            path_attr = "file2_path"
            file_name = self.file2_name
        else:
            # invalid key
            return  

        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])

        if not file_path:
            return
        
        setattr(self, path_attr, file_path)
        file_name.set(file_path.split("/")[-1])

    def _create_dataframes(self):
        return pd.read_excel(self.file1_path), pd.read_excel(self.file2_path)


    def _run(self):
        try:
            if self.menu_bar.output_dir_path:
                if self.file1_path:
                    if self.file1_path:
                        messagebox.showinfo("Process Started", "Processing started. Please wait...")
                        file1_df, file2_df = self._create_dataframes()
                        diff = file1_df.compare(file2_df)
                        diff.to_excel(os.path.join(self.menu_bar.output_dir_path, "compared.xlsx"))
                    else:
                        messagebox.showerror("Missing file", "Please select file 2.")
                else:
                    messagebox.showerror("Missing file", "Please select file 1.")
                    
            else:
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")