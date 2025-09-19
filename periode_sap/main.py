import os
from tkinter import Frame, messagebox, ttk
import pandas as pd
from tkinter import filedialog

class PeriodeSAP(Frame):
    def __init__(self, master, menu_bar):
        super().__init__(master, highlightbackground="gray24")

        self.file_name="debug_periode"

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.menu_bar = menu_bar

        self._init_layout()    
        
    def _init_layout(self):
        # Initializes the layout of the PeriodeSAP frame
        label_container = Frame(self)
        label_container.grid(row=0, column=0, sticky="ew")
        label_container.grid_columnconfigure(0, weight=1)
        
        ttk.Label(label_container, text="1. Select SAP file", anchor="center").grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Label(label_container, text="").grid(row=1, column=0) # Empty label for spacing
        ttk.Button(label_container, text="Select File and Run", command=self._run).grid(row=4, column=0, padx=5)
    

    def _run(self):
        # Called when the 'Select File' button is clicked
        try:
            if self.menu_bar.output_dir_path :
                file_path = self._upload_action()
                if not file_path:
                    return
                messagebox.showinfo("Process Started", "Processing started. Please wait...")
                periode_df = self._change_periodejahr(file_path)
                self._save_file(periode_df, {"title": "File Saved", "message": f"The import file has been saved to '{self.file_name}_Import.txt' in the output directory."})
            else:
                messagebox.showerror("Output Directory Not Set", "Please select an output directory first.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def _upload_action(self):
        # Opens a file dialog to select an Excel file
        # Excel file to be uploaded is specifc for this application
        file_path = filedialog.askopenfilename(title="Select text file", filetypes=[("Text files", "*.txt")])
        if not file_path:
            return
        
        return file_path
    
    def _change_periodejahr(self, file_path):
        df = pd.read_csv(file_path, sep="\t", decimal=",")
        df["PERIODJAHR"] = df["PERIODJAHR"].apply(lambda x: x-1)
        return df
    
    def _save_file(self, df, messageInfo):
        try:
            file_path = os.path.join(self.menu_bar.output_dir_path, f"{self.file_name}.xlsx")
            df.to_excel(file_path, index=False)

            messagebox.showinfo(messageInfo["title"], messageInfo["message"])
        except Exception as e:
            messagebox.showerror("Save Failed", f"An error occurred while saving the file:\n{str(e)}")


