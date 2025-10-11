from tkinter import Frame, messagebox, ttk
import traceback
import pandas as pd

from utils.file_ops import upload_file, save_file
from utils.dataframe_ops import clean_dataframe

class ActualsIT(Frame):
    def __init__(self, master, menu_bar):
        super().__init__(master, highlightbackground="gray24")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.menu_bar = menu_bar

        self._init_layout()    
        
    def _init_layout(self):
        label_container = Frame(self)
        label_container.grid(row=0, column=0, sticky="ew")
        label_container.grid_columnconfigure(0, weight=1)
        
        ttk.Label(label_container, text="1. Select Acutals IT File", anchor="center").grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="2. Wait for prompt to say data copy is complete", anchor="center").grid(row=1, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="3. Check output folder for file", anchor="center").grid(row=2, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="").grid(row=3, column=0) # Empty label for spacing
        ttk.Button(label_container, text="Select File and Run", command=self._run).grid(row=4, column=0, padx=5)
    
    def _run(self):
        try:
            if not self.menu_bar.output_dir_path :
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")
                return

            df_cleaned = self._select_and_clean_file()
            df_import = self._transfer_dataframes(df_cleaned)
            self._save_import_file(df_import)

        except Exception as e:
            tb = traceback.format_exc()
            messagebox.showerror("Error", f"An error occurred: \n{str(e)}\n\n{tb}")

    def _select_and_clean_file(self):
        df = upload_file(file_type="excel", reader="excel", header=10)
        return clean_dataframe(df, ["Unnamed: 1", "Data Mov.", "Date", "Data Doc.", "Causale", "Unnamed: 17"])
    
    def _save_import_file(self, df_import):
        save_file(df=df_import, 
                      file_name="ActualsIT_Import.txt", 
                      output_loc=self.menu_bar.output_dir_path, 
                      message_info={"title": "File Saved", 
                                   "message": "The import file has been saved to 'ActualsIT_Import.txt' in the output directory"})
    
    def _format_code(self, prefix, suffix):
        return f"{int(prefix)}"[:2] + f"-{int(suffix):02d}-0000"
    
    def _extract_code(self, row, current_code, code_prefix):
        if pd.notna(row["Unnamed: 1"]) and row["Unnamed: 1"] != 0.0:
                if code_prefix == "":
                    code_prefix = row["Unnamed: 1"]
                else:
                    current_code = self._format_code(code_prefix, row["Unnamed: 1"])
                    code_prefix = ""
        return current_code, code_prefix
    
    def _create_new_row(self, row, current_code):
        date = pd.to_datetime(row["Data Doc."]).strftime("%d.%m.%Y")
        periode_date = pd.to_datetime(row["Data Mov."])
        periode = 12 if periode_date.month == 1 else periode_date.month - 1
        betrag = f"{row['Unnamed: 17']:.2f}".replace(".", ",")
        PROJECT_CODE = "1015-70108-01-1"
        COST_CENTER = "NF"

        return {
            "Kostenart": current_code,
            "Kostenstelle": COST_CENTER,
            "Kostenträger": PROJECT_CODE,
            "Betrag": betrag,
            "Belegdatum": date,
            "Belegnummer": "",
            "Belegtext": row["Causale"],
            "Extra-Kosteninfo": "",
            "Jahr": periode_date.year,
            "Periode": periode,
            "Tag": periode_date.day
        }
    
    def _transfer_dataframes(self, df):
        # Transfers data from the cost dataframes to the import dataframes
        code_prefix = ""
        current_code = ""
        new_rows = []

        for _, row in df.iterrows():
            current_code, code_prefix = self._extract_code(row, current_code, code_prefix)

            if pd.notna(row["Data Doc."]) and row["Data Doc."] != 0.0 and current_code:
                new_row = self._create_new_row(row, current_code)
                new_rows.append(new_row)
            
        return pd.DataFrame(new_rows) if new_rows else pd.DataFrame()