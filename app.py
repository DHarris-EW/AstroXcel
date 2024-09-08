import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
import sv_ttk
from tkinter import messagebox


class App(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        container = tk.Frame(self)
        container.grid(padx=20, pady=5)

        self.master_file_name = tk.StringVar()
        self.merge_file_names = tk.StringVar()
        self.output_folder_name = tk.StringVar()

        self.master_file = ""
        self.sheet = None
        self.master_df = None
        self.files_to_merge = ""
        self.output_directory_path = ""
        self.merge_files_path = []
        self.merge_ws = []
        self.column_name = tk.StringVar()

        CreateFiles(container, self).grid(row=1, column=0, pady=5, sticky="nesw")
        MergeFiles(container, self).grid(row=2, column=0, pady=10, sticky="nesw")
        Options(container, self).grid(row=0, column=0, pady=10, sticky="nesw")

    def SelectDirectory(self, event=None):
        self.output_directory_path = str(filedialog.askdirectory())
        self.output_folder_name.set(self.output_directory_path)

    def UploadAction(self, event=None):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        print(file_path)
        self.master_file_name.set(file_path.split("/")[-1])
        self.master_file = load_workbook(file_path, data_only=True)
        self.sheet = self.master_file.active
        data = self.sheet.values
        columns = next(data)
        self.master_df = pd.DataFrame(data, columns=columns)
        if "Count_ID" not in self.master_df.columns:
                self.master_df.insert(0, "Count_ID", range(1, len(self.master_df) + 1))

    def UploadActionMultiple(self, event=None):
        files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        df_files = []
        file_paths = ""

        for path in files:
            file_paths += path.split("/")[-1] + "\n"
            df_files.append(pd.read_excel(path, header=0))
            self.merge_files_path.append(path)

        self.merge_file_names.set(file_paths)
        self.files_to_merge = df_files

    def CreateMasterFile(self):
        if isinstance(self.master_df, pd.DataFrame):
            if self.output_directory_path:
                self.master_df.to_excel(self.output_directory_path + "/" + "master_file.xlsx", index=False)
            else:
                messagebox.showerror("Missing Output file", "No Output file selected.")
        else:
            messagebox.showerror("Missing File Error", "No original or master file selected.")

    def CreateWorkBooks(self):
        if isinstance(self.master_df, pd.DataFrame):
            if self.output_directory_path:
                if self.column_name.get() not in self.master_df.columns:
                    if "Zuordnung" in self.master_df.columns:
                        self.column_name.set("Zuordnung")
                    else:
                        self.column_name.set = self.master_df.columns[0]
                for index, group in self.master_df.groupby(self.column_name.get()):
                    file_name = self.output_directory_path + "/" + str(index) + ".xlsx"
                    group.to_excel(file_name, index=False)
                    wb = load_workbook(file_name)
                    ws = wb.active

                    for col in ws.columns:
                        max_length = max(len(str(cell.value)) for cell in col)

                        ws.column_dimensions[get_column_letter(col[0].column)].width = max_length
                    
                    wb.save(file_name)
            else:
                messagebox.showerror("Missing Output file", "No Output file selected.")

        else:
            messagebox.showerror("Missing File Error", "No original or master file selected.")

    def MergeWorkBooks(self):
        if self.merge_files_path:
            for path in self.merge_files_path:
                wb = load_workbook(path)
                ws = wb.active
                self.merge_ws.append(ws)

            if self.master_file and "Count_ID" in self.master_df.columns:
                print(self.master_file)
                if self.output_directory_path:
                    target_wb = load_workbook(self.master_file)
                    target_ws = target_wb.active

                    update_cells = {}

                    for ws_idx, ws in enumerate(self.merge_ws):
                        for row_idx, row in enumerate(ws.rows):
                            update_cells[row[0].value] = {"ws_idx": ws_idx, "row_idx": row_idx + 1}


                    for col in self.merge_ws[0].columns:
                        max_length = max(len(str(cell.value)) for cell in col)

                        target_ws.column_dimensions[get_column_letter(col[0].column)].width = max_length

                    for row in target_ws.rows:
                        if row[0].value in update_cells:
                            for cell_new, cell in zip(self.merge_ws[update_cells[row[0].value]["ws_idx"]][update_cells[row[0].value]["row_idx"]], row):
                                    if cell_new.has_style:
                                        cell.font = copy(cell_new.font)
                                        cell.border = copy(cell_new.border)
                                        cell.fill = copy(cell_new.fill)
                                        cell.number_format = copy(cell_new.number_format)
                                        cell.protection = copy(cell_new.protection)
                                        cell.alignment = copy(cell_new.alignment)
                                    cell.value = cell_new.value

                    target_ws.delete_cols(1)
                    target_wb.save(self.output_directory_path + "/" + 'target_file.xlsx')
                else:
                    messagebox.showerror("Missing Output file", "No Output file selected.")
            else:
                messagebox.showerror("Missing Master file", "No Master file selected. <Coud_ID> column must be included.")
        else:
            messagebox.showerror("Missing Merge files", "No files to merge selected.")


class Options(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, highlightbackground="gray24", highlightthickness=2, padx=5, pady=5)
        self.columnconfigure(0, weight=1)

        ttk.Label(self, text="Output Folder Name", anchor="center").grid(row=1, column=0, padx=5, pady=10, sticky="nesw")
        ttk.Label(self, textvariable=controller.output_folder_name, anchor="center").grid(row=2, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Select Folder", command=controller.SelectDirectory).grid(row=3, column=0, padx=5, pady=5, sticky="nswe")

class CreateFiles(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, highlightbackground="gray24", highlightthickness=2, padx=5, pady=5)
        self.columnconfigure(0, weight=1)

        ttk.Label(self, text="File Name", anchor="center").grid(row=0, column=0, padx=5, pady=10, sticky="nesw")
        ttk.Label(self, textvariable=controller.master_file_name, anchor="center").grid(row=1, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Select File", command=controller.UploadAction).grid(row=2, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Label(self, text="Column Name", anchor="center").grid(row=3, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Entry(self, textvariable=controller.column_name).grid(row=4, column=0, padx=5,pady=5, sticky="nesw")
        ttk.Button(self, text="Create Groups", command=controller.CreateWorkBooks).grid(row=5, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Create Master File", command=controller.CreateMasterFile).grid(row=6, column=0, padx=5, pady=5, sticky="nesw")

class MergeFiles(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, highlightbackground="gray24", highlightthickness=2, padx=5, pady=5)
        self.columnconfigure(0, weight=1)

        ttk.Label(self, text="File Names", anchor="center").grid(row=0, column=0, padx=5, pady=10, sticky="nesw")
        ttk.Label(self, textvariable=controller.merge_file_names, anchor="center").grid(row=1, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Select Merge Files", command=controller.UploadActionMultiple).grid(row=2, column=0, padx=5, pady=5, sticky="nesw")
        ttk.Button(self, text="Merge", command=controller.MergeWorkBooks).grid(row=3, column=0, padx=5, pady=5, sticky="nesw")

app = App()
sv_ttk.set_theme("dark")
app.title("XCelAstro")
app.resizable(False, False)
app.mainloop()