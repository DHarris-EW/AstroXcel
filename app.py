import sv_ttk
import tkinter as tk

from actuals_czk.main import ActualsCzk
from actuals_czk.compare_files import CompareFiles
from actuals_it.main import ActualsIT
from opliste.main import OPListe
from periode_sap.main import PeriodeSAP
from menubar import MenuBar

class App(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        
        sv_ttk.set_theme("dark")
        self.title("AstroXCel")
        self.resizable(False, False)
        self.geometry("500x400")
       
        self.frames = {}
        
        self._create_main_container()
        self._init_menu_bar()
        self._init_frames()

        self.show_frame(ActualsCzk)
        
    def _create_main_container(self):
        # Creates the main container for the application
        self.container = tk.Frame(self)
        self.container.grid(row=0, column=0, padx=20, pady=5)
        # Fill horizontally and vertically
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
    def _init_frames(self):
        # Initializes the frames for ActualsCzk and OPListe
        for F in (ActualsCzk, OPListe, ActualsIT, PeriodeSAP, CompareFiles):
            frame = F(self.container, self.menu_bar)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
            
    def _init_menu_bar(self):
        # Initializes the menu bar
        self.menu_bar = MenuBar(self)
        self.config(menu=self.menu_bar)

    def show_frame(self, page):
        # Raises the specified frame to the top
        frame = self.frames[page]
        frame.tkraise()

if __name__ == "__main__":
    app = App()
    app.mainloop()