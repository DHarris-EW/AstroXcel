from tkinter import Frame
from opliste.create_files import CreateFiles
from opliste.merge_files import MergeFiles

class OPListe(Frame):
    def __init__(self, master, menu_bar):
        super().__init__(master)
        
        self.menu_bar = menu_bar
    
        self._init_layout(menu_bar)

    def _init_layout(self, menu_bar):
        self.create_files = CreateFiles(self, menu_bar)
        self.merge_files = MergeFiles(self, menu_bar)
        
        self.merge_files.grid(row=0, column=1, pady=5, padx=5, sticky="nesw")
        self.create_files.grid(row=0, column=0, pady=5, padx=5, sticky="nesw")
