# Python 3.12.7
# Standard Libraries
import tkinter
from tkinter import ttk

# Third-party Libraries
import pandas as pd
import sv_ttk

class SpendingAnalyzerApp():
    def __init__(self):
        self.setup_ui()
    

    def setup_ui(self):
        self.root = tkinter.Tk()
        self.root.state('zoomed')

        sv_ttk.set_theme("dark")

        options = ["American Express", "Discover"]
        self.statement_type = ttk.Combobox(self.root, values=options)
        self.statement_type.grid(column=0, row=0, pady=10)

        self.process_button = ttk.Button(self.root, text="Process", command=self.process_statement)
        self.process_button.grid(column=1, row=0)

        self.root.mainloop()


    def process_statement(self):
        def _process_amex(file):
            df = pd.read_excel(file, skiprows=6)
            print(df)

        
        def _process_discover(file):
            dfs = pd.read_html(file, skiprows=6)
            df = dfs[0]
            print(df)


        match self.statement_type.get():
            case "American Express":
                _process_amex(r"C:\Users\antho\Programming\spending_analyzer\test_files\amex\activity.xlsx")
            
            case "Discover":
                _process_discover(r"C:\Users\antho\Programming\spending_analyzer\test_files\discover\Discover-Statement-20241109.xls")


app = SpendingAnalyzerApp()