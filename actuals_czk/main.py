import os
from tkinter import Frame, messagebox, ttk
import pandas as pd
import numpy as np
from tkinter import filedialog

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
                df_czk_cost = self._create_cost_dataframes(file_path)
                df_czk_import = self._transfer_dataframes(df_czk_cost)
                # Saves the import dataframes to an Excel file
                self.save_file(df_czk_import)
            else:
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
    
    def _create_import_dataframes(self):
        # Creates one empty dataframes for CZK with the specified columns ready to be imported
        columns = ["Kostenart", "Kostenstelle", "Kostenträger", "Betrag", "Belegdatum", "Belegnummer", "Belegtext", "Extra-Kosteninfo"]
        return pd.DataFrame(columns=columns)

    def _upload_action(self):
        # Opens a file dialog to select an Excel file
        # Excel file to be uploaded is specifc for this application
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        return file_path
        
    def _create_cost_dataframes(self, file_path):
        # Creates one dataframe for CZK based on the cost excel file uploaded by the user
        
        columns_to_keep = ["ACCT # (konv.)", "Unnamed: 15", "DESCRIPTION", "Date", "POHODA #. (č. dokladu)"]
    
        # Read the Excel file and create dataframes for the CZK sheet
        temp_df_czk = pd.read_excel(file_path, sheet_name="ACTUALS CZK", header=8)
            
        # Replace 0.0, "", and " " with NaN 
        temp_df_czk = temp_df_czk.replace({0.0: np.nan, "": np.nan, " ": np.nan}, regex=False)
       
        # Keep only the specified columns and drop completely empty columns
        df_czk_cost = temp_df_czk[[col for col in columns_to_keep if col in temp_df_czk.columns]]
        
        # Drop rows where all cells in the row are NaN. Retain rows with actual data
        df_czk_cost = df_czk_cost.dropna(how='all')
        
        # Fill NaN values in rows with data that used to be 0.0, "", or " "
        df_czk_cost = df_czk_cost.fillna(0.00)
        
        return df_czk_cost
    
    def _transfer_dataframes(self, df_czk_cost):
        # Transfers data from the cost dataframes to the import dataframes
        df_czk_import = self._create_import_dataframes()

        # Iterate through the columns and fill the import dataframes with the appropriate values
        for col in df_czk_import.columns:
            if col == "Kostenart":
                # Remove the last 4 characters from the "ACCT # (konv.)" column and append "0000"
                df_czk_import[col] = df_czk_cost["ACCT # (konv.)"].astype(str).str[:-4] + "0000"
            elif col == "Kostenstelle":
                # Static value for "Kostenstelle"
                df_czk_import[col] = "NF"
            elif col == "Kostenträger":
                # Static value for "Kostenträger"
                df_czk_import[col] = "1015-70108-01-1"
            elif col == "Betrag":
                # Format the "Unnamed: 15" column (column p) as a float with two decimal places 
                df_czk_import[col] = df_czk_cost["Unnamed: 15"].apply(lambda x: f"{x:.2f}")
            elif col == "Belegdatum":
                # Use the "Date" column for "Belegdatum"
                df_czk_import[col] = pd.to_datetime(df_czk_cost["Date"]).dt.strftime("%d.%m.%Y")
            elif col == "Belegnummer":
                # Use the "POHODA #. (č. dokladu)" column for "Belegnummer"
                df_czk_import[col] = df_czk_cost["POHODA #. (č. dokladu)"]
            elif col == "Belegtext":
                # Use the "DESCRIPTION" column for "Belegtext"
                df_czk_import[col] = df_czk_cost["DESCRIPTION"]
                
        return df_czk_import

    def save_file(self, df_czk_import):
        
        # Format 'Betrag' column with comma as decimal separator if it exists
        df_czk_import["Betrag"] = df_czk_import["Betrag"].astype(str).str.replace(".", ",", regex=False)
        
        # txt file are uploaded into the system
        file_path = os.path.join(self.menu_bar.output_dir_path, "ACTUALS CZK.txt")
        try:
            df_czk_import.to_csv(file_path, sep="\t", index=False, encoding="utf-8-sig")
        except Exception as e:
            messagebox.showerror("Save Failed", f"An error occurred while saving the file:\n{str(e)}")
        finally:
            if os.path.exists(file_path):
                messagebox.showinfo("Success", "Data Copy Successful. \nFile saved as 'Actual Cost_Sesam_Import.xlsx'")
       