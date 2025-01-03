    def CheckHistory(self):
        # Back Button Function
        def ButtonBack():
            # Destroying History view widgets
            self.History.destroy()
            self.tree.destroy()
            self.BackButton.destroy()

            # Recreating Main Menu Buttons
            self.button1 = tk.Button(self.LowerFrame, text="+ Add Expense", font=("Arial Black", 8), fg="white", bg="#61876E", command=self.AddExpense, activebackground="#0081C9")
            self.button1.grid(column=0, row=0, padx=15, pady=10)

            self.button2 = tk.Button(self.LowerFrame, text="+ Add Profit", font=("Arial Black", 8), fg="white", bg="#61876E", command=self.AddProfit, activebackground="#0081C9")
            self.button2.grid(column=1, row=0, padx=15, pady=10)

            self.button3 = tk.Button(self.LowerFrame, text="Check History", font=("Arial Black", 8), fg="white", bg="#61876E", command=self.CheckHistory, activebackground="#0081C9")
            self.button3.grid(column=2, row=0, padx=15, pady=10)

            # Reinitializing Pie Chart
            self.Graph()

        # Destroy Main Menu Widgets
        self.button1.destroy()
        self.button2.destroy()
        self.button3.destroy()
        self.label1.destroy()
        self.label2.destroy()
        self.label3.destroy()
        self.label4.destroy()

        # Label for History Title
        self.History = tk.Label(self.LowerFrame, text="HISTORY", font=('Arial Black', 12), fg="white", bg="#3C6255")
        self.History.grid(column=1, row=0, pady=(10, 20))

        # Read Expense Data from Excel
        df = pd.read_excel('Expenses.xlsx')

        # Treeview Widget for History
        columns = ("Type", "Name", "Cost", "Date")
        self.tree = ttk.Treeview(self.LowerFrame, columns=columns, show="headings", height=8)
        self.tree.grid(columnspan=3, row=1, padx=20, pady=10)

        # Define Treeview Columns and Headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)

        # Populate Treeview with DataFrame Rows
        for index, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

        # Back Button
        self.BackButton = tk.Button(self.LowerFrame, text="BACK", font=("Arial Black", 8), fg="white", bg="#61876E", command=ButtonBack, activebackground="#0081C9")
        self.BackButton.grid(column=1, row=2, pady=(20, 10))