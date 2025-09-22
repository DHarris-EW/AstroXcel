import os
from tkinter import Frame, messagebox, ttk
import pandas as pd
import numpy as np
from tkinter import filedialog

class ActualsIT(Frame):
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
        
        ttk.Label(label_container, text="1. Select Acutals IT File", anchor="center").grid(row=0, column=0, padx=5, sticky="ew")
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
                df_it_cost = self._create_cost_dataframes(file_path)
                df_eur_import = self._transfer_dataframes(df_it_cost)
                self.save_file(df_eur_import, "ActualsIT_Import.txt", {"title": "File Saved", "message": "The import file has been saved to 'ActualsIT_Import.txt' in the output directory"})
            else:
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def _upload_action(self):
        # Opens a file dialog to select an Excel file
        # Excel file to be uploaded is specifc for this application
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        return file_path
        
    def _create_cost_dataframes(self, file_path):
        # Creates one dataframe for IT based on the cost excel file uploaded by the user
        
        columns_to_keep = ["Unnamed: 1", "Data Mov.", "Data Doc.", "Causale", "Unnamed: 17"]

        # Read the Excel file and create a dataframe
        temp_df_it = pd.read_excel(file_path, header=10)
        # Debug: Save the raw dataframe to inspect its structure
        # temp_df_it.to_excel(os.path.join(self.menu_bar.output_dir_path, "debug_it.xlsx"), index=False)
        
        # Replace 0.0, "", and " " with NaN 
        temp_df_it = temp_df_it.replace({0.0: np.nan, "": np.nan, " ": np.nan}, regex=False)
        
        # Keep only the specified columns and drop completely empty columns
        df_it_cost = temp_df_it[[col for col in columns_to_keep if col in temp_df_it.columns]]
        # Drop rows where all cells in the row are NaN. Retain rows with actual data
        df_it_cost = df_it_cost.dropna(how='all')
        
        # Fill NaN values in rows with data that used to be 0.0, "", or " "
        df_it_cost = df_it_cost.fillna(0.00)
        
        return df_it_cost
    
    def _format_code(self, prefix, suffix):
        return f"{int(prefix)}"[:2] + f"-{int(suffix):02d}-0000"
    
    def _transfer_dataframes(self, df_it_cost):
        # Transfers data from the cost dataframes to the import dataframes
        code_prefix = ""
        current_code = ""
        new_rows = []

        # Iterate through the columns and fill the import dataframes with the appropriate values
        for _, row in df_it_cost.iterrows():
            if pd.notna(row["Unnamed: 1"]) and row["Unnamed: 1"] != 0.0:
                if code_prefix == "":
                    code_prefix = row["Unnamed: 1"]
                else:
                    code = self._format_code(code_prefix, row["Unnamed: 1"])
                    current_code = code
                    code_prefix = ""
                # Continue to next row after setting the code
                continue
            # If there's a valid "Data Doc." and current_code is set, create a new row as all information needed is on that row
            if pd.notna(row["Data Doc."]) and row["Data Doc."] != 0.0 and current_code:
                date = pd.to_datetime(row["Data Doc."])
                periode = 12 if date.month == 1 else date.month - 1
                betrag = f"{row['Unnamed: 17']:.2f}".replace(".", ",")
                new_row = {
                    "Kostenart": current_code,
                    "Kostenstelle": "NF",
                    "Kostenträger": "1015-70108-01-1",
                    "Belegdatum": date,
                    "Betrag": betrag,
                    "Belegnummer": "",
                    "Belegtext": row["Causale"],
                    "Extra-Kosteninfo": "",
                    "Jahr": date.year,
                    "Periode": periode,
                    "Tag": date.day
                }
                new_rows.append(new_row)
            
        # If new_rows create a dataframe with the new_rows
        if new_rows:
            df_it_import = pd.DataFrame(new_rows)
        else:
            df_it_import = pd.DataFrame()
            
        return df_it_import

    def save_file(self, file, file_name, messageInfo):    
        # Creates txt file which are uploaded into the system
        try:
            file_path = os.path.join(self.menu_bar.output_dir_path, file_name)
            file.to_csv(file_path, sep="\t", index=False, encoding="utf-8-sig")
    
            if os.path.exists(file_path):
                messagebox.showinfo(messageInfo["title"], messageInfo["message"])
        except Exception as e:
            messagebox.showerror("Save Failed", f"An error occurred while saving the file:\n{str(e)}")
