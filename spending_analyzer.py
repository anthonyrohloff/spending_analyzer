# Python 3.12.7
# Standard Libraries
import tkinter
from tkinter import ttk
import sqlite3

# Third-party Libraries
from openpyxl import load_workbook
import pandas as pd
import sv_ttk

class SpendingAnalyzerApp():
    def __init__(self):
        self.setup_db()
        self.setup_ui()
    

    def setup_db(self):
        con = sqlite3.connect("spending_analyzer.db")
        cur = con.cursor()
        
        cur.execute("""CREATE TABLE IF NOT EXISTS statement_data(
                        id INTEGER PRIMARY KEY,
                        date TEXT,
                        type TEXT,
                        account INTEGER,
                        description TEXT,
                        amount REAL,
                        category TEXT
                        )""")

        con.commit()
        con.close()

    def setup_ui(self):
        self.root = tkinter.Tk()
        self.root.state('zoomed')

        sv_ttk.set_theme("dark")

        options = ["American Express", "Discover"]
        self.statement_type = ttk.Combobox(self.root, values=options)
        self.statement_type.grid(column=0, row=0, pady=10, padx=10)

        self.process_button = ttk.Button(self.root, text="Process", command=self.process_statement)
        self.process_button.grid(column=0, row=1, pady=10, padx=10)

        self.root.mainloop()


    def process_statement(self):
        def _process_amex(file):
            df = pd.read_excel(file, 
                               skiprows=6,
                               usecols=["Date", 
                                        "Appears On Your Statement As", 
                                        "Amount", 
                                        "Category"])
            df.rename(columns={"Date": "date", "Appears On Your Statement As": "description", "Amount": "amount", "Category": "category"}, inplace=True)
            
            workbook = load_workbook(file)
            sheet = workbook.active
            df["account"] = sheet["A5"].value[-5:]
            df["type"] = "American Express"
            
            df.to_sql('statement_data', con, if_exists='append', index=False)

        
        def _process_discover(file):
            dfs = pd.read_html(file, skiprows=6)
            df = dfs[0] 
            print(df)


        con = sqlite3.connect("spending_analyzer.db")
        cur = con.cursor()

        # TODO: Add auto-detection
        match self.statement_type.get():
            case "American Express":
                _process_amex(r"C:\Users\antho\Programming\spending_analyzer\test_files\amex\activity.xlsx")
            
            case "Discover":
                _process_discover(r"C:\Users\antho\Programming\spending_analyzer\test_files\discover\Discover-Statement-20241109.xls")

        # cur.execute("SELECT * FROM statement_data")
        # results = cur.fetchall()
        # for result in results:
        #     print(result)

        con.commit()
        con.close()

app = SpendingAnalyzerApp()