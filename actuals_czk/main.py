import os
from tkinter import Frame, messagebox, ttk
import pandas as pd
import numpy as np
from tkinter import filedialog
from openpyxl.utils import get_column_letter

class ActualsCzk(Frame):
    def __init__(self, master, menu_bar):
        super().__init__(master, highlightbackground="gray24")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.menu_bar = menu_bar

        self._init_layout()    
        
    def _init_layout(self):
        # Initializes the layout of the ActualsCzk frame
        label_container = Frame(self)
        label_container.grid(row=0, column=0, sticky="ew")
        label_container.grid_columnconfigure(0, weight=1)
        
        ttk.Label(label_container, text="1. Select Acutals CZK File", anchor="center").grid(row=0, column=0, padx=5, sticky="ew")
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
                # Saves the import dataframes to an Excel file
                self.save_file(df_eur_import, df_czk_import)
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
        temp_df_eur = pd.read_excel(file_path, sheet_name="ACTUALS EUR", header=8)
        temp_df_czk = pd.read_excel(file_path, sheet_name="ACTUALS CZK", header=8)
            
        # Replace 0.0, "", and " " with NaN 
        temp_df_czk = temp_df_czk.replace({0.0: np.nan, "": np.nan, " ": np.nan}, regex=False)
        temp_df_eur = temp_df_eur.replace({0.0: np.nan, "": np.nan, " ": np.nan}, regex=False)
       
        # Keep only the specified columns and drop completely empty columns
        df_eur_cost = temp_df_eur[[col for col in columns_to_keep if col in temp_df_eur.columns]]
        df_czk_cost = temp_df_czk[[col for col in columns_to_keep if col in temp_df_czk.columns]]
        
        # Drop rows where all cells in the row are NaN. Retain rows with actual data
        df_eur_cost = df_eur_cost.dropna(how='all')
        df_czk_cost = df_czk_cost.dropna(how='all')
        
        # Fill NaN values in rows with data that used to be 0.0, "", or " "
        df_eur_cost = df_eur_cost.fillna(0.00)
        df_czk_cost = df_czk_cost.fillna(0.00)
        
        return df_eur_cost, df_czk_cost
    
    def _transfer_dataframes(self, df_czk_cost, df_eur_cost):
        # Transfers data from the cost dataframes to the import dataframes
        df_eur_import, df_czk_import = self._create_import_dataframes()

        # Iterate through the columns and fill the import dataframes with the appropriate values
        for col in df_eur_import.columns:
            if col == "Kostenart":
                # Remove the last 4 characters from the "ACCT # (konv.)" column and append "0000"
                df_eur_import[col] = df_eur_cost["ACCT # (konv.)"].astype(str).str[:-4] + "0000"
                df_czk_import[col] = df_czk_cost["ACCT # (konv.)"].astype(str).str[:-4] + "0000"
            elif col == "Kostenstelle":
                # Static value for "Kostenstelle"
                df_eur_import[col] = "NF"
                df_czk_import[col] = "NF"
            elif col == "Kostenträger":
                # Static value for "Kostenträger"
                df_eur_import[col] = "1015-70108-01-1"
                df_czk_import[col] = "1015-70108-01-1"
            elif col == "Betrag":
                # Format the "Unnamed: 15" column (column p) as a float with two decimal places 
                df_eur_import[col] = df_eur_cost["Unnamed: 15"].apply(lambda x: f"{x:.2f}")
                df_czk_import[col] = df_czk_cost["Unnamed: 15"].apply(lambda x: f"{x:.2f}")
            elif col == "Belegdatum":
                # Use the "Date" column for "Belegdatum"
                df_eur_import[col] = pd.to_datetime(df_eur_cost["Date"]).dt.strftime("%d.%m.%Y")
                df_czk_import[col] = pd.to_datetime(df_czk_cost["Date"]).dt.strftime("%d.%m.%Y")
            elif col == "Belegnummer":
                # Use the "POHODA #. (č. dokladu)" column for "Belegnummer"
                df_eur_import[col] = df_eur_cost["POHODA #. (č. dokladu)"]
                df_czk_import[col] = df_czk_cost["POHODA #. (č. dokladu)"]
            elif col == "Belegtext":
                # Use the "DESCRIPTION" column for "Belegtext"
                df_eur_import[col] = df_eur_cost["DESCRIPTION"]
                df_czk_import[col] = df_czk_cost["DESCRIPTION"]
                
        return df_czk_import, df_eur_import

    def save_file(self, df_eur_import, df_czk_import):
        
        # Format 'Betrag' column with comma as decimal separator if it exists
        for df in [df_czk_import, df_eur_import]:
            df["Betrag"] = df["Betrag"].astype(str).str.replace(".", ",", regex=False)
        
        # txt file are uploaded into the system 
        df_czk_import.to_csv(os.path.join(self.menu_bar.output_dir_path, "ACTUALS CZK.txt"), sep="\t", index=False)
        df_eur_import.to_csv(os.path.join(self.menu_bar.output_dir_path, "ACTUALS EUR.txt"), sep="\t", index=False)
        
        # Saves the import dataframes to an Excel file with auto-adjusted column widths
        # Not upload to the system. Can be used for reference and verification
        file_path = os.path.join(self.menu_bar.output_dir_path, "Actual Cost_Sesam_Import.xlsx")
        try:
            with pd.ExcelWriter(file_path) as writer:
                df_czk_import.to_excel(writer, index=False, sheet_name="ACTUALS CZK")
                df_eur_import.to_excel(writer, index=False, sheet_name="ACTUALS EUR")
                
                
                worksheets = [writer.sheets["ACTUALS CZK"], writer.sheets["ACTUALS EUR"]]

                for worksheet in worksheets:
                    for col in worksheet.columns:
                        max_length = max(len(str(cell.value)) for cell in col)
                        worksheet.column_dimensions[get_column_letter(col[0].column)].width = max_length
                        
            # Show a message box to indicate the file has been saved successfully or not.
            if os.path.exists(file_path):
                messagebox.showinfo("Success", "Data Copy Successful. \nFile saved as 'Actual Cost_Sesam_Import.xlsx'")
        except Exception as e:
            messagebox.showerror("Save Failed", f"An error occurred while saving the file:\n{str(e)}")
