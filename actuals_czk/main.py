from tkinter import Frame, messagebox, ttk
import traceback
import pandas as pd
from utils.file_ops import save_file, upload_file
from utils.dataframe_ops import clean_dataframe

class ActualsCzk(Frame):
    def __init__(self, master, menu_bar):
        super().__init__(master)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.menu_bar = menu_bar

        self._init_layout()    
        
    def _init_layout(self):
        label_container = Frame(self)
        label_container.grid(row=0, column=0, sticky="ew")
        label_container.grid_columnconfigure(0, weight=1)
        
        ttk.Label(label_container, text="1. Select Actuals CZK File", anchor="center").grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="2. Wait for prompt to say data copy is complete", anchor="center").grid(row=1, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="3. Check output folder for file", anchor="center").grid(row=2, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="").grid(row=3, column=0) # Empty label for spacing
        ttk.Button(label_container, text="Select File and Run", command=self._run).grid(row=4, column=0, padx=5)
    
    def _run(self):
        # Called when the 'Select File' button is clicked 
        try:
            if not self.menu_bar.output_dir_path :
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")

            df_sheets = upload_file(file_type="excel", reader="sheets", header=8)
            if df_sheets is False:
                return
            df_eur_cost, df_czk_cost = self._create_cost_dataframes(df_sheets)
            df_eur_import, df_czk_import = self._transfer_dataframes(df_eur_cost, df_czk_cost)
            eur_duplicates, czk_duplicates = self._check_duplicates(df_eur_import, df_czk_import)
            self._process_duplicates_and_save("ACTUALS_EUR", "EUR", df_eur_import, eur_duplicates)
            self._process_duplicates_and_save("ACTUALS_CZK", "CZK", df_czk_import, czk_duplicates)
        except Exception as e:
            tb = traceback.format_exc()
            messagebox.showerror("Error", f"An error occurred: \n{str(e)}\n\n{tb}")
            
    def _create_import_dataframes(self):
        # Creates two empty dataframes for EUR and CZK with the specified columns ready to be imported
        columns = ["Kostenart", "Kostenstelle", "Kostenträger", "Betrag", "Belegdatum", "Belegnummer", "Belegtext", "Extra-Kosteninfo"]
        return pd.DataFrame(columns=columns), pd.DataFrame(columns=columns)
        
    def _create_cost_dataframes(self, df_sheets):
        df_eur_cost = clean_dataframe(df_sheets["ACTUALS EUR"], ["ACCT # (konv.)", "Unnamed: 15", "DESCRIPTION", "Date", "POHODA #. (č. dokladu)"])
        df_czk_cost = clean_dataframe(df_sheets["ACTUALS CZK"], ["ACCT # (konv.)", "Unnamed: 15", "DESCRIPTION", "Date", "POHODA #. (č. dokladu)"])
        
        return df_eur_cost, df_czk_cost
    
    def _map_cost_to_import(self, df_cost, df_import):
        COST_CENTRE = "NF"
        PROJECT_CODE = "1015-70108-01-1"
        # Remove the last 4 characters from the "ACCT # (konv.)" column and append "0000"
        df_import["Kostenart"] = df_cost["ACCT # (konv.)"].astype(str).str[:-4] + "0000"
        # Static value for "Kostenstelle"
        df_import["Kostenstelle"] = COST_CENTRE
        # Static value for "Kostenträger"
        df_import["Kostenträger"] = PROJECT_CODE
        # Format the "Unnamed: 15" column (column p) as a float with two decimal places 
        df_import["Betrag"] = df_cost["Unnamed: 15"].apply(lambda x: f"{x:.2f}".replace(".", ","))
        # Use the "Date" column for "Belegdatum"
        df_import["Belegdatum"] = pd.to_datetime(df_cost["Date"]).dt.strftime("%d.%m.%Y")
        # Use the "POHODA #. (č. dokladu)" column for "Belegnummer"
        df_import["Belegnummer"] = df_cost["POHODA #. (č. dokladu)"]
        # Use the "DESCRIPTION" column for "Belegtext"
        df_import["Belegtext"] = df_cost["DESCRIPTION"]
        
        return df_import
    
    def _transfer_dataframes(self, df_eur_cost, df_czk_cost):
        # Transfers data from the cost dataframes to the import dataframes
        df_eur_import, df_czk_import = self._create_import_dataframes()

        df_eur_import = self._map_cost_to_import(df_eur_cost, df_eur_import)
        df_czk_import = self._map_cost_to_import(df_czk_cost, df_czk_import)
        
        return df_eur_import, df_czk_import

    def _check_duplicates(self, df_eur_import, df_czk_import):
        # Extracts duplicate rows
        eur_duplicates = df_eur_import[df_eur_import.duplicated(keep=False)]
        czk_duplicates = df_czk_import[df_czk_import.duplicated(keep=False)]

        return eur_duplicates, czk_duplicates
        
    def _process_duplicates_and_save(self, file_name, sheet_name, df, df_duplicates):
        if not df_duplicates.empty:
            answer = messagebox.askyesno(f"{sheet_name} Duplicates Found", f"Warning: Duplicate {sheet_name} rows found.\n\nIf you would like to proceed and remove {sheet_name} duplicates, click 'Yes'.\n\nIf you would like to keep the duplicates and view them in a seperate .txt file, click 'no' ?")
            if answer:
                save_file(df.drop_duplicates(), f"{file_name}_Imports.txt", self.menu_bar.output_dir_path, {"title": f"{sheet_name} Duplicates Removed", "message": f"Duplicate {sheet_name} rows have been removed. The import file has been saved to '{file_name}_Imports.txt' in the output directory."})
            else:
                save_file(df, f"{file_name}_Imports.txt", self.menu_bar.output_dir_path, {"title": "File Saved", "message": f"The import file has been saved with its duplicates to '{file_name}_Imports.txt' in the output directory."})
                save_file(df_duplicates, f"{file_name}_Duplicates.txt", self.menu_bar.output_dir_path, {"title": f"{sheet_name} Duplicates Saved", "message": f"Duplicate {sheet_name} rows have been saved to '{file_name}_Duplicates.txt' in the output directory."})
        else:   
            # If no duplicates are found, returns the merged dataframe as is
            save_file(df, f"{file_name}_Imports.txt", self.menu_bar.output_dir_path, {"title": "File Saved", "message": f"The import file has been saved to '{file_name}_Import.txt' in the output directory."})