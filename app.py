import tkinter as tk
from tkinter import ttk
from flexo_diecut_form import DiecutForm
from flexo_printing_form import FlexoPrintingForm
from flexo_rewinding_form import ReWindingForm
from ruling_form import RulingForm
from sheeting_form import SheetingForm
from fuel_form import FuelForm

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Manufacturing Dashboard")
        self.geometry("400x300")

        self.setup_theme()

        self.create_buttons_frame()

    def setup_theme(self):
        style = ttk.Style()

        style.theme_create("azure", parent="default", settings={
            "TButton": {
                "configure": {
                    "background": "#0078d4", 
                    "foreground": "#ffffff",  
                    "padding": [10, 5]
                },
                "map": {
                    "background": [("active", "#005a9e")] 
                }
            },
            "TFrame": {
                "configure": {
                    "background": "#f3f3f3"  
                }
            }
        })
        style.theme_use("azure")

    def create_buttons_frame(self):
        
        button_frame = ttk.Frame(self)
        button_frame.pack(pady=20, padx=20, fill='both', expand=True)

        sheeting_button = ttk.Button(button_frame, text="Open Sheeting Form", command=self.open_sheeting_form)
        sheeting_button.pack(pady=5, fill='x')

        fuel_button = ttk.Button(button_frame, text="Open Fuel Form", command=self.open_fuel_form)
        fuel_button.pack(pady=5, fill='x')

        ruling_button = ttk.Button(button_frame, text="Open Ruling Form", command=self.open_ruling_form)
        ruling_button.pack(pady=5, fill='x')

        flexo_printing_button = ttk.Button(button_frame, text="Open Flexo Printing Form", command=self.open_flexo_printing_form)
        flexo_printing_button.pack(pady=5, fill='x')

        diecut_button = ttk.Button(button_frame, text="Open Diecut Form", command=self.open_diecut_form)
        diecut_button.pack(pady=5, fill='x')

        rewinding_button = ttk.Button(button_frame, text="Open Re-Winding Form", command=self.open_rewinding_form)
        rewinding_button.pack(pady=5, fill='x')

    def open_sheeting_form(self):
        sheeting_window = tk.Toplevel(self)
        sheeting_form = SheetingForm(sheeting_window)
        sheeting_form.pack(fill='both', expand=True)

    def open_fuel_form(self):
        fuel_window = tk.Toplevel(self)
        fuel_form = FuelForm(fuel_window)
        fuel_form.pack(fill='both', expand=True)

    def open_ruling_form(self):
        ruling_window = tk.Toplevel(self)
        ruling_form = RulingForm(ruling_window)
        ruling_form.pack(fill='both', expand=True)

    def open_flexo_printing_form(self):
        flexo_printing_window = tk.Toplevel(self)
        flexo_printing_form = FlexoPrintingForm(flexo_printing_window)
        flexo_printing_form.pack(fill='both', expand=True)

    def open_diecut_form(self):
        diecut_window = tk.Toplevel(self)
        diecut_form = DiecutForm(diecut_window)
        diecut_form.pack(fill='both', expand=True)

    def open_rewinding_form(self):
        rewinding_window = tk.Toplevel(self)
        rewinding_form = ReWindingForm(rewinding_window)
        rewinding_form.pack(fill='both', expand=True)

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
