from tkinter import Menu, messagebox, StringVar, filedialog
from actuals_czk.main import ActualsCzk
from actuals_czk.compare_files import CompareFiles
from actuals_it.main import ActualsIT
from opliste.main import OPListe
from periode_sap.main import PeriodeSAP

class MenuBar(Menu):

    def __init__(self, parent):
        Menu.__init__(self, parent)
        
        # Variable to hold the output directory label and path
        # Used across the application to save files in specified directory
        self.output_dir_label = StringVar()
        self.output_dir_path = ""

        self.file_menu = Menu(self, tearoff=0)
        self.file_menu.add_command(label="Output File", command=self.select_directory)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=parent.quit)
        self.add_cascade(label="File", menu=self.file_menu)

        app_menu = Menu(self, tearoff=0)
        app_menu.add_command(label="OPListe", command=lambda: parent.show_frame(OPListe))
        app_menu.add_command(label="ActualsCzk", command=lambda: parent.show_frame(ActualsCzk))
        app_menu.add_command(label="ActualsIT", command=lambda: parent.show_frame(ActualsIT))
        app_menu.add_command(label="Periode SAP", command=lambda: parent.show_frame(PeriodeSAP))
        app_menu.add_command(label="Compare Czk", command=lambda: parent.show_frame(CompareFiles))

        self.add_cascade(label="Select App", menu=app_menu)
        
        disclaimer  = Menu(self, tearoff=0)
        disclaimer.add_command(label="Disclaimer", command=lambda: messagebox.showinfo("Privacy Notice", "This application processes all data locally.\nIt does not upload or store data externally."))
        disclaimer.add_command(label="Version 1.16.1", command="")
        self.add_cascade(label="Help", menu=disclaimer)

    def select_directory(self):
        # Selects the output directory for saving files
        if self.output_dir_path:
            self.file_menu.delete(self.file_menu.index("end"))
            
        self.output_dir_path = str(filedialog.askdirectory())
        
        # Updates the output directory label in the menu
        if self.output_dir_path:
            self.output_dir_label.set(self.output_dir_path)
            self.file_menu.insert_command(1, label=f" - ${self.output_dir_label.get()}", state="disabled")