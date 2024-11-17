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
        """Calls functions to set up database and UI, and creates self.file for file uploading"""
        self.setup_db()
        self.setup_ui()
        self.file = None
    

    def setup_db(self):
        """Creates database if needed
        
        DB info:
        id -- unique row id
        date -- date of row item from statement
        type -- type of account (American Express, Discover, etc...)
        account -- account number / last few digits of card
        description -- description of charge
        amount -- amount of charge
        category -- category of charge
        """
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
        """Sets up UI window and widgets"""
        self.root = tkinter.Tk()
        # self.root.state('zoomed')

        sv_ttk.set_theme("dark")

        # Set up widgets
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
        """Updates visualizations to most recent values"""
        def _update_pie(self, parameter, col, row):
            con = sqlite3.connect("spending_analyzer.db")
            cur = con.cursor()

            # Fetch unique values for the parameter
            cur.execute(f"SELECT DISTINCT {parameter} FROM statement_data")
            labels = [label[0] for label in cur.fetchall()]

            if not labels:
                print(f"No data found for parameter: {parameter}")
                con.close()
                return

            label_totals = {}
            for label in labels:
                # Fetch sum of amounts for each label based on parameter
                cur.execute(f"SELECT SUM(amount) FROM statement_data WHERE {parameter} = ? AND amount > 0", (label,))
                total = cur.fetchone()[0] or 0
                label_totals[label] = total

            labels = label_totals.keys()
            sizes = label_totals.values()

            # Create pie chart
            fig = Figure(figsize=(5, 4), dpi=100)
            ax = fig.add_subplot(111)
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', shadow=True, startangle=140)
            ax.axis('equal')

            # Render pie chart on Tkinter canvas
            canvas = FigureCanvasTkAgg(fig, master=self.root)
            canvas.draw()
            canvas.get_tk_widget().grid(column=col, row=row, pady=10, padx=10)

            # Add toolbar
            toolbar_frame = ttk.Frame(self.root)
            toolbar_frame.grid(column=col, row=row + 1, pady=10, padx=10)
            toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
            toolbar.update()

            con.close()


        _update_pie(self, "category", 0, 3)
        _update_pie(self, "account", 0, 5)

    def upload_file(self):
        """Sets self.file to point to uploaded file"""
        self.file = filedialog.askopenfilename()


    def process_statement(self):
        """Processes uploaded file based on type of statement"""
        # TODO: normalize categories
        def _process_amex(self):
            """Processes an American Express statement"""
            # Read in data from excel sheet
            df = pd.read_excel(self.file, 
                               skiprows=6,
                               usecols=["Date", 
                                        "Appears On Your Statement As", 
                                        "Amount", 
                                        "Category"])
            df.rename(columns={"Date": "date", "Appears On Your Statement As": "description", "Amount": "amount", "Category": "category"}, inplace=True)
            
            # Get account number
            workbook = load_workbook(self.file)
            sheet = workbook.active
            df["account"] = sheet["A5"].value[-5:]
            df["type"] = "American Express"
            
            # Send to db
            df.to_sql('statement_data', con, if_exists='append', index=False)

        
        def _process_discover(self):
            """Processes a Discover Card statement"""
            # Read in data from excel sheet
            dfs = pd.read_html(self.file, skiprows=6)
            df = dfs[0]
            df = df.iloc[:, [0, 2, 3, 4]]
            df.rename(columns={0: "date", 2: "description", 3: "amount", 4: "category"}, inplace=True)
            
            # Get account number
            dfs_account = pd.read_html(self.file)
            df_account = dfs_account[0]
            account = df_account[0][2][-4:]
            df["account"] = account
            df["type"] = "Discover"

            # Send to db
            df.to_sql('statement_data', con, if_exists='append', index=False)

        con = sqlite3.connect("spending_analyzer.db")

        # Manually set the upload file type
        # TODO: Add auto-detection
        match self.statement_type.get():
            case "American Express":
                _process_amex(self)
            
            case "Discover":
                _process_discover(self)

        con.commit()
        con.close()
        
        self.update_visualizations()


app = SpendingAnalyzerApp()