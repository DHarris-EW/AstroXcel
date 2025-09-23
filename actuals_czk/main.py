import os
from tkinter import Frame, messagebox, ttk
import pandas as pd
import numpy as np
from tkinter import filedialog
from actuals_czk.compare_files import CompareFiles

class ActualsCzk(Frame):
    def __init__(self, master, menu_bar):
        super().__init__(master)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.menu_bar = menu_bar

        self._init_layout(menu_bar)    
        
    def _init_layout(self, menu_bar):
        # Initializes the layout of the ActualsCzk frame
        label_container = Frame(self)
        label_container.grid(row=0, column=0, sticky="ew")
        label_container.grid_columnconfigure(0, weight=1)
        self.compare_files = CompareFiles(self, menu_bar)

        
        ttk.Label(label_container, text="1. Select Actuals CZK File", anchor="center").grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="2. Wait for prompt to say data copy is complete", anchor="center").grid(row=1, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="3. Check output folder for file", anchor="center").grid(row=2, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="").grid(row=3, column=0) # Empty label for spacing
        ttk.Button(label_container, text="Select File and Run", command=self._run).grid(row=4, column=0, padx=5)

    
    def _run(self):
        # Called when the 'Select File' button is clicked 
        try:
            if self.menu_bar.output_dir_path :
                file_path = self._upload_action()
                if not file_path:
                    return
                messagebox.showinfo("Process Started", "Processing started. Please wait...")
                # Continues after a file is selected
                df_eur_cost, df_czk_cost = self._create_cost_dataframes(file_path)
                df_eur_import, df_czk_import = self._transfer_dataframes(df_eur_cost, df_czk_cost)
                # Check for duplicates and save the final import file or duplicates file if duplicates are found
                self._check_duplicates(df_eur_import, df_czk_import)
            else:
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
    
    def _create_import_dataframes(self):
        # Creates two empty dataframes for EUR and CZK with the specified columns ready to be imported
        columns = ["Kostenart", "Kostenstelle", "Kostenträger", "Betrag", "Belegdatum", "Belegnummer", "Belegtext", "Extra-Kosteninfo"]
        return pd.DataFrame(columns=columns), pd.DataFrame(columns=columns)

    def _upload_action(self):
        # Opens a file dialog to select an Excel file
        # Excel file to be uploaded is specifc for this application
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        return file_path
        
    def _create_cost_dataframes(self, file_path):
        # Creates two dataframs one for EUR and one for CZK based on the cost excel file uploaded by the user
        
        columns_to_keep = ["ACCT # (konv.)", "Unnamed: 15", "DESCRIPTION", "Date", "POHODA #. (č. dokladu)"]
    
        # Read the Excel file and create dataframes for EUR and CZK sheets
        sheets = pd.read_excel(file_path, sheet_name=["ACTUALS EUR", "ACTUALS CZK"], header=8)
            
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
        
        df_eur_cost = clean_dataframe(sheets["ACTUALS EUR"])
        df_czk_cost = clean_dataframe(sheets["ACTUALS CZK"])
        
        return df_eur_cost, df_czk_cost
    
    def _transfer_dataframes(self, df_eur_cost, df_czk_cost):
        # Transfers data from the cost dataframes to the import dataframes
        df_eur_import, df_czk_import = self._create_import_dataframes()

        # Remove the last 4 characters from the "ACCT # (konv.)" column and append "0000"
        df_eur_import["Kostenart"] = df_eur_cost["ACCT # (konv.)"].astype(str).str[:-4] + "0000"
        df_czk_import["Kostenart"] = df_czk_cost["ACCT # (konv.)"].astype(str).str[:-4] + "0000"
    
        # Static value for "Kostenstelle"
        df_eur_import["Kostenstelle"] = "NF"
        df_czk_import["Kostenstelle"] = "NF"
    
        # Static value for "Kostenträger"
        df_eur_import["Kostenträger"] = "1015-70108-01-1"
        df_czk_import["Kostenträger"] = "1015-70108-01-1"
    
        # Format the "Unnamed: 15" column (column p) as a float with two decimal places 
        df_eur_import["Betrag"] = df_eur_cost["Unnamed: 15"].apply(lambda x: f"{x:.2f}")
        df_czk_import["Betrag"] = df_czk_cost["Unnamed: 15"].apply(lambda x: f"{x:.2f}")
    
        # Use the "Date" column for "Belegdatum"
        df_eur_import["Belegdatum"] = pd.to_datetime(df_eur_cost["Date"]).dt.strftime("%d.%m.%Y")
        df_czk_import["Belegdatum"] = pd.to_datetime(df_czk_cost["Date"]).dt.strftime("%d.%m.%Y")
    
        # Use the "POHODA #. (č. dokladu)" column for "Belegnummer"
        df_eur_import["Belegnummer"] = df_eur_cost["POHODA #. (č. dokladu)"]
        df_czk_import["Belegnummer"] = df_czk_cost["POHODA #. (č. dokladu)"]
    
        # Use the "DESCRIPTION" column for "Belegtext"
        df_eur_import["Belegtext"] = df_eur_cost["DESCRIPTION"]
        df_czk_import["Belegtext"] = df_czk_cost["DESCRIPTION"]
        
        return df_eur_import, df_czk_import

    def _check_duplicates(self, df_eur_import, df_czk_import):
        # Replace decimal points with commas in the "Betrag" column
        df_eur_import["Betrag"] = df_eur_import["Betrag"].astype(str).str.replace(".", ",", regex=False)
        df_czk_import["Betrag"] = df_czk_import["Betrag"].astype(str).str.replace(".", ",", regex=False)
        # Extracts duplicate rows
        eur_duplicates = df_eur_import[df_eur_import.duplicated(keep=False)]
        czk_duplicates = df_czk_import[df_czk_import.duplicated(keep=False)]
        
        # If duplicates are found, prompts the user to either remove them or save them to a .txt file
        def process_duplicates(file_name, sheet_name, df, df_duplicates):
            if not df_duplicates.empty:
                answer = messagebox.askyesno(f"{sheet_name} Duplicates Found", f"Warning: Duplicate {sheet_name} rows found.\n\nIf you would like to proceed and remove {sheet_name} duplicates, click 'Yes'.\n\nIf you would like to view the duplicates in a .txt file, click 'no' ?")
                if answer:
                    self._save_file(df.drop_duplicates(), f"{file_name}_Imports.txt", {"title": f"{sheet_name} Duplicates Removed", "message": f"Duplicate {sheet_name} rows have been removed. The import file has been saved to '{file_name}_Import.txt' in the output directory."})
                else:
                    self._save_file(df_duplicates, f"{file_name}_Duplicates.txt", {"title": f"{sheet_name} Duplicates Saved", "message": f"Duplicate EUR rows have been saved to '{file_name}_Duplicates.txt' in the output directory."})
            else:   
                # If no duplicates are found, returns the merged dataframe as is
                self._save_file(df, f"{file_name}_Import.txt", {"title": "File Saved", "message": f"The import file has been saved to '{file_name}_Import.txt' in the output directory."})
        
        process_duplicates("ACTUALS_EUR", "EUR", df_eur_import, eur_duplicates)
        process_duplicates("ACTUALS_CZK", "CZK", df_czk_import, czk_duplicates)

    def _save_file(self, file, file_name, messageInfo):    
        # Creates txt file which are uploaded into the system
        try:
            file_path = os.path.join(self.menu_bar.output_dir_path, file_name)
            file.to_csv(file_path, sep="\t", index=False, encoding="utf-8-sig")
    
            messagebox.showinfo(messageInfo["title"], messageInfo["message"])
        except Exception as e:
            messagebox.showerror("Save Failed", f"An error occurred while saving the file:\n{str(e)}")
