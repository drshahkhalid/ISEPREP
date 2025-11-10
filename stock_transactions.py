import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from db import connect_db
from language_manager import lang
import csv
import logging
from popup_utils import custom_popup, custom_askyesno, custom_dialog

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class StockTransactions(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.role = app.role
        self.pack(fill="both", expand=True)
        self.render_transactions_page()

    def render_transactions_page(self):
        # Apply modern theme
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", rowheight=25, font=('Helvetica', 10), background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=('Helvetica', 11, 'bold'), background="#d3d3d3", foreground="#333333")
        style.map("Treeview", background=[('selected', '#0077CC'), ('!selected', '#ffffff')], foreground=[('selected', 'white')])
        style.configure("Accent.TButton", font=("Helvetica", 10), background="#2980B9", foreground="white")
        style.configure("Search.TEntry", font=("Helvetica", 10))
        style.configure("Label.TLabel", font=("Helvetica", 10), background="#F5F5F5", foreground="#333333")

        # Title
        tk.Label(
            self,
            text=lang.t("stock_transactions.title", fallback="Stock Transactions"),
            font=("Helvetica", 20, "bold"),
            bg="#F5F5F5",
            fg="#2C3E50"
        ).pack(pady=10)

        # Search frame
        search_frame = tk.Frame(self, bg="#F5F5F5")
        search_frame.pack(fill="x", pady=5, padx=10)

        ttk.Label(search_frame, text=lang.t("search", fallback="Search"), style="Label.TLabel").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(search_frame, width=30, style="Search.TEntry")
        self.search_entry.pack(side="left", padx=5)
        ttk.Button(search_frame, text=lang.t("filter", fallback="Filter"), command=self.search_transactions, style="Accent.TButton").pack(side="left", padx=5)
        ttk.Button(search_frame, text=lang.t("clear", fallback="Clear"), command=self.load_transactions, style="Accent.TButton").pack(side="left", padx=5)
        ttk.Button(search_frame, text=lang.t("export_excel", fallback="Export to Excel"), command=self.export_to_excel, style="Accent.TButton").pack(side="right", padx=5)

        # Treeview frame with scrollbars
        tree_frame = tk.Frame(self, bg="#F5F5F5")
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Table columns
        columns = [
            "Date", "Time", "unique_id", "code", "Description",
            "Expiry_date", "Batch_Number", "Scenario", "Kit", "Module",
            "Qty_IN", "IN_Type", "Qty_Out", "Out_Type",
            "Third_Party", "End_User", "Discrepancy", "Remarks", "Movement_Type"
        ]
        display_columns = [
            "Date", "Time", "code", "Description",
            "Expiry_date", "Batch_Number", "Scenario", "Kit", "Module",
            "Qty_IN", "IN_Type", "Qty_Out", "Out_Type",
            "Third_Party", "End_User", "Discrepancy", "Remarks", "Movement_Type"
        ]

        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20, displaycolumns=display_columns)
        headers = {
            "Date": "Date",
            "Time": "Time",
            "unique_id": "Unique ID",
            "code": "Code",
            "Description": "Description",
            "Expiry_date": "Expiry Date",
            "Batch_Number": "Batch Number",
            "Scenario": "Scenario",
            "Kit": "Kit",
            "Module": "Module",
            "Qty_IN": "Qty IN",
            "IN_Type": "IN Type",
            "Qty_Out": "Qty OUT",
            "Out_Type": "Out Type",
            "Third_Party": "Third Party",
            "End_User": "End User",
            "Discrepancy": "Discrepancy",
            "Remarks": "Remarks",
            "Movement_Type": "Movement Type"
        }
        widths = {
            "Date": 100,
            "Time": 80,
            "unique_id": 0,  # Hidden
            "code": 100,
            "Description": 480,
            "Expiry_date": 100,
            "Batch_Number": 100,
            "Scenario": 100,
            "Kit": 100,
            "Module": 100,
            "Qty_IN": 80,
            "IN_Type": 100,
            "Qty_Out": 80,
            "Out_Type": 100,
            "Third_Party": 100,
            "End_User": 100,
            "Discrepancy": 80,
            "Remarks": 150,
            "Movement_Type": 100
        }
        alignments = {
            "Date": "w",
            "Time": "w",
            "code": "w",
            "Description": "w",
            "Expiry_date": "w",
            "Batch_Number": "w",
            "Scenario": "w",
            "Kit": "w",
            "Module": "w",
            "Qty_IN": "e",
            "IN_Type": "w",
            "Qty_Out": "e",
            "Out_Type": "w",
            "Third_Party": "w",
            "End_User": "w",
            "Discrepancy": "e",
            "Remarks": "w",
            "Movement_Type": "w"
        }

        for col in columns:
            self.tree.heading(col, text=headers.get(col, col))
            self.tree.column(col, width=widths.get(col, 100), anchor=alignments.get(col, "center"))

        # Scrollbars
        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        # Row styling
        self.tree.tag_configure('oddrow', background='#F5F5F5')
        self.tree.tag_configure('evenrow', background='#E8ECEF')

        self.load_transactions()

    def load_transactions(self):
        """Load all transactions into treeview."""
        self.tree.delete(*self.tree.get_children())
        conn = connect_db()
        cursor = conn.cursor()

        try:
            cursor.execute("""
                SELECT
                    Date, Time, unique_id, code, Description,
                    Expiry_date, Batch_Number, Scenario, Kit, Module,
                    Qty_IN, IN_Type, Qty_Out, Out_Type,
                    Third_Party, End_User, Discrepancy, Remarks, Movement_Type
                FROM stock_transactions
                ORDER BY Date DESC, Time DESC
            """)

            for i, row in enumerate(cursor.fetchall()):
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        row[0], row[1], row[2], row[3], row[4],
                        row[5], row[6], row[7], row[8], row[9],
                        row[10], row[11], row[12], row[13],
                        row[14], row[15], row[16], row[17], row[18]
                    ),
                    tags=(tag,)
                )
            logging.debug("Successfully loaded transactions")
        except Exception as e:
            logging.error(f"Error loading transactions: {str(e)}")
            messagebox.showerror("Error", f"Failed to load transactions: {str(e)}", parent=self)
        finally:
            cursor.close()
            conn.close()

    def search_transactions(self):
        """Search by code or movement type."""
        query = self.search_entry.get().strip()
        if not query:
            self.load_transactions()
            return

        self.tree.delete(*self.tree.get_children())
        conn = connect_db()
        cursor = conn.cursor()

        try:
            cursor.execute("""
                SELECT
                    Date, Time, unique_id, code, Description,
                    Expiry_date, Batch_Number, Scenario, Kit, Module,
                    Qty_IN, IN_Type, Qty_Out, Out_Type,
                    Third_Party, End_User, Discrepancy, Remarks, Movement_Type
                FROM stock_transactions
                WHERE code LIKE ? OR Movement_Type LIKE ?
                ORDER BY Date DESC, Time DESC
            """, (f"%{query}%", f"%{query}%"))

            for i, row in enumerate(cursor.fetchall()):
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        row[0], row[1], row[2], row[3], row[4],
                        row[5], row[6], row[7], row[8], row[9],
                        row[10], row[11], row[12], row[13],
                        row[14], row[15], row[16], row[17], row[18]
                    ),
                    tags=(tag,)
                )
            logging.debug(f"Successfully searched transactions for query: {query}")
        except Exception as e:
            logging.error(f"Error searching transactions: {str(e)}")
            messagebox.showerror("Error", f"Failed to search transactions: {str(e)}", parent=self)
        finally:
            cursor.close()
            conn.close()

    def export_to_excel(self):
        """Export displayed data to CSV."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        try:
            with open(file_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                # Write header
                writer.writerow([
                    "Date", "Time", "Unique ID", "Code", "Description",
                    "Expiry Date", "Batch Number", "Scenario", "Kit", "Module",
                    "Qty IN", "IN Type", "Qty OUT", "Out Type",
                    "Third Party", "End User", "Discrepancy", "Remarks", "Movement Type"
                ])
                # Write rows
                for item in self.tree.get_children():
                    writer.writerow(self.tree.item(item)["values"])
            logging.debug(f"Successfully exported transactions to {file_path}")
            messagebox.showinfo("Export", "Transactions exported successfully!", parent=self)
        except Exception as e:
            logging.error(f"Error exporting transactions: {str(e)}")
            messagebox.showerror("Error", f"Failed to export transactions: {str(e)}", parent=self)
