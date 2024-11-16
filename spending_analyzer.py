# Python 3.12.7
# Standard Libraries
import tkinter
from tkinter import ttk, filedialog
import sqlite3

# Third-party Libraries
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import pandas as pd
import sv_ttk

class SpendingAnalyzerApp():
    def __init__(self):
        self.setup_db()
        self.setup_ui()
        self.file = None
    

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
        # self.root.state('zoomed')

        sv_ttk.set_theme("dark")

        self.upload_button = ttk.Button(self.root, text="Upload File", command=self.upload_file)
        self.upload_button.grid(column=0, row=0, pady=10, padx=10)

        options = ["American Express", "Discover"]
        self.statement_type = ttk.Combobox(self.root, values=options, state="readonly")
        self.statement_type.grid(column=0, row=1, pady=10, padx=10)

        self.process_button = ttk.Button(self.root, text="Process", command=self.process_statement)
        self.process_button.grid(column=0, row=2, pady=10, padx=10)

        self.update_visualizations()

        self.root.mainloop()


    def update_visualizations(self):
        con = sqlite3.connect("spending_analyzer.db")
        cur = con.cursor()

        cur.execute("SELECT DISTINCT category FROM statement_data")
        labels = cur.fetchall()
        print(labels)

        # TODO: loop through all labels^^^ 
        # select all amounts for each category, sum them
        # append to 'sizes' list
        labels = ['Apples', 'Bananas', 'Cherries', 'Dates']
        sizes = [35, 25, 25, 15]
        colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
        explode = (0, 0, 0, 0)  

        fig = Figure(figsize=(5, 4), dpi=100)
        ax = fig.add_subplot(111)
        ax.pie(
            sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140
        )
        ax.axis('equal')  

        canvas = FigureCanvasTkAgg(fig, master=self.root)
        canvas.draw()
        canvas.get_tk_widget().grid(column=0, row=3, pady=10, padx=10)

        toolbar_frame = ttk.Frame(self.root)
        toolbar_frame.grid(column=0, row=4, pady=10, padx=10)

        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
        toolbar.update()

        con.commit()
        con.close()


    def upload_file(self):
        self.file = filedialog.askopenfilename()


    def process_statement(self):
        # TODO: normalize categories
        def _process_amex(self):
            df = pd.read_excel(self.file, 
                               skiprows=6,
                               usecols=["Date", 
                                        "Appears On Your Statement As", 
                                        "Amount", 
                                        "Category"])
            df.rename(columns={"Date": "date", "Appears On Your Statement As": "description", "Amount": "amount", "Category": "category"}, inplace=True)
            
            workbook = load_workbook(self.file)
            sheet = workbook.active
            df["account"] = sheet["A5"].value[-5:]
            df["type"] = "American Express"
            
            df.to_sql('statement_data', con, if_exists='append', index=False)

        
        def _process_discover(self):
            dfs = pd.read_html(self.file, skiprows=6)
            df = dfs[0]
            df = df.iloc[:, [0, 2, 3, 4]]
            df.rename(columns={0: "date", 2: "description", 3: "amount", 4: "category"}, inplace=True)
            
            dfs_account = pd.read_html(self.file)
            df_account = dfs_account[0]
            account = df_account[0][2][-4:]
            df["account"] = account
            df["type"] = "Discover"

            df.to_sql('statement_data', con, if_exists='append', index=False)


        con = sqlite3.connect("spending_analyzer.db")
        cur = con.cursor()

        # TODO: Add auto-detection
        match self.statement_type.get():
            case "American Express":
                _process_amex(self)
            
            case "Discover":
                _process_discover(self)

        cur.execute("SELECT * FROM statement_data")
        results = cur.fetchall()
        for result in results:
            print(result)

        con.commit()
        con.close()


app = SpendingAnalyzerApp()