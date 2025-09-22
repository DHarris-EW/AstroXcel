from tkinter import Frame, ttk, filedialog, messagebox, StringVar, simpledialog
import numpy as np
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
        
        self.file1_name = StringVar(value="File 1 Name: ")
        self.file1_path = ""
        self.file2_name = StringVar(value="File 2 Name: ")
        self.file2_path = ""
        
        self._init_layout()
        
    def _init_layout(self):
        ttk.Label(self, textvariable=self.file1_name, anchor="center").grid(row=0, column=0, padx=5, pady=10, sticky="nesw")
        ttk.Button(self, text="Select file 1", command=lambda: self._upload_action("file_1")).grid(row=1, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Label(self, textvariable=self.file2_name, anchor="center").grid(row=2, column=0, padx=5, pady=10, sticky="nesw")
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
        file_name.set(file_name.get() + file_path.split("/")[-1])

    def _create_dataframes(self):
        # Creates two dataframs one for EUR and one for CZK based on the cost excel file uploaded by the user
        
        columns_to_keep = ["ACCT # (konv.)", "Unnamed: 15", "DESCRIPTION", "Date", "POHODA #. (č. dokladu)"]
    
        # Read the Excel file and create dataframes for EUR and CZK sheets
        file1_sheets = pd.read_excel(self.file1_path, sheet_name=["ACTUALS EUR", "ACTUALS CZK"], header=8)
        file2_sheets = pd.read_excel(self.file2_path, sheet_name=["ACTUALS EUR", "ACTUALS CZK"], header=8)
            
        def clean_dataframe(df):
            # Replace 0.0, "", and " " with NaN 
            df = df.replace({0.0: np.nan, "": np.nan, " ": np.nan}, regex=False)
            # Keep only the specified columns and drop completely empty columns
            df = df[[col for col in columns_to_keep if col in df.columns]]
            # Drop rows where all cells in the row are NaN. Retain rows with actual data
            df = df.dropna(how='all')
            # Fill NaN values in rows with data that used to be 0.0, "", or " "
            df = df.fillna(0.00)
            return df
        
        file1_df_eur = clean_dataframe(file1_sheets["ACTUALS EUR"])
        file1_df_czk = clean_dataframe(file1_sheets["ACTUALS CZK"])
        file2_df_eur = clean_dataframe(file2_sheets["ACTUALS EUR"])
        file2_df_czk = clean_dataframe(file2_sheets["ACTUALS CZK"])
        
        return file1_df_eur, file1_df_czk, file2_df_eur, file2_df_czk

    def _run(self):
        if self.menu_bar.output_dir_path:
            if self.file1_path:
                if self.file1_path:
                    messagebox.showinfo("Process Started", "Processing started. Please wait...")
                    file1_df_eur, file1_df_czk, file2_df_eur, file2_df_czk  = self._create_dataframes()
                    eur_comparison = file1_df_eur.merge(file2_df_eur, how="outer", indicator=True)
                    czk_comparison = file1_df_czk.merge(file2_df_czk, how="outer", indicator=True)
                    eur_unique_rows = eur_comparison[eur_comparison['_merge'] != 'both']
                    czk_unique_rows = czk_comparison[czk_comparison['_merge'] != 'both']
                    # eur_common_rows = eur_comparison[eur_comparison['_merge'] == 'both'].drop(columns=['_merge'])
                    # czk_common_rows = czk_comparison[czk_comparison['_merge'] == 'both'].drop(columns=['_merge'])

                    datasets = {"eur_new": eur_unique_rows, "czk_new": czk_unique_rows}
                    
                    for name, df in datasets.items():
                        if df.empty:
                            print(f"Skipping {name}")
                            continue
                        new_rows = []
                        jahr = simpledialog.askinteger("jahr", "Enter the Jahr")
                        periode = simpledialog.askinteger("periode", "Enter the Periode")
                        tag = simpledialog.askinteger("tag", "Enter the Tag")
                        for _, row in df.iterrows():
                            date = pd.to_datetime(row["Date"])
                            betrag = f"{row['Unnamed: 15']:.2f}".replace(".", ",")
                            
                            new_row = {
                                "Kostenart": str(row["ACCT # (konv.)"])[:-4] + "0000",
                                "Kostenstelle": "NF",
                                "Kostenträger": "1015-70108-01-1",
                                "Belegdatum": date.strftime("%d.%m.%Y"),
                                "Betrag": betrag,
                                "Belegnummer": row["POHODA #. (č. dokladu)"],
                                "Belegtext": row["DESCRIPTION"],
                                "Extra-Kosteninfo": "",
                                "Jahr": jahr if jahr else "",
                                "Periode": periode if periode else "",
                                "Tag": tag if tag else ""
                            }
                            new_rows.append(new_row)
                        if new_rows:
                            output = pd.DataFrame(new_rows)
                            file_name = f"{name}.xlsx"
                            output.to_excel(os.path.join(self.menu_bar.output_dir_path, file_name), index=False)
                            print(f"Saved {file_name}")
                        else:
                            print(f"No new rows for {name}")

                else:
                    messagebox.showerror("Missing file", "Please select file 2.")
            else:
                messagebox.showerror("Missing file", "Please select file 1.")
                
        else:
            messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")
