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

    # ---------------- Translation helpers (canonical EN -> display) ----------------
    def _to_display_in_type(self, canonical: str) -> str:
        """
        Localize IN_Type (canonical English stored in DB) to the active language.
        Falls back gracefully to the canonical string when no translation exists.
        """
        if canonical is None:
            return ""
        return lang.enum_to_display("stock_in.in_types_map", canonical, fallback=canonical)

    def _to_display_movement(self, canonical: str) -> str:
        """
        Localize Movement_Type (canonical English stored in DB) to the active language.
        Uses 'stock_transactions.movement_types_map' from the translation files.
        Falls back to a prettified canonical (title case, '_' → space) if missing.
        """
        if not canonical:
            return ""
        fallback = canonical.replace("_", " ").title()
        return lang.enum_to_display("stock_transactions.movement_types_map", canonical, fallback=fallback)

    def _to_display_remarks(self, text: str) -> str:
        """
        Optionally localize standardized remarks via 'stock_transactions.remarks_map'.
        If the text doesn't match a known key, show the original text.
        """
        if text is None:
            return ""
        return lang.enum_to_display("stock_transactions.remarks_map", text, fallback=text)

    # ---------------- Render page ----------------
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

        ttk.Label(search_frame, text=lang.t("stock_transactions.search", fallback="Search"), style="Label.TLabel").pack(side="left", padx=5)
        self.search_entry = ttk.Entry(search_frame, width=30, style="Search.TEntry")
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<Return>", lambda event: self.search_transactions())
        ttk.Button(search_frame, text=lang.t("stock_transactions.filter", fallback="Filter"), command=self.search_transactions, style="Accent.TButton").pack(side="left", padx=5)
        ttk.Button(search_frame, text=lang.t("stock_transactions.clear", fallback="Clear"), command=self.load_transactions, style="Accent.TButton").pack(side="left", padx=5)
        ttk.Button(search_frame, text=lang.t("stock_transactions.export_excel", fallback="Export to Excel"), command=self.export_to_excel, style="Accent.TButton").pack(side="right", padx=5)

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

        # Translated headers
        headers = {
            "Date": lang.t("stock_transactions.date", fallback="Date"),
            "Time": lang.t("stock_transactions.time", fallback="Time"),
            "unique_id": lang.t("stock_transactions.unique_id", fallback="Unique ID"),
            "code": lang.t("stock_transactions.code", fallback="Code"),
            "Description": lang.t("stock_transactions.description", fallback="Description"),
            "Expiry_date": lang.t("stock_transactions.expiry_date", fallback="Expiry Date"),
            "Batch_Number": lang.t("stock_transactions.batch_number", fallback="Batch Number"),
            "Scenario": lang.t("stock_transactions.scenario", fallback="Scenario"),
            "Kit": lang.t("stock_transactions.kit", fallback="Kit"),
            "Module": lang.t("stock_transactions.module", fallback="Module"),
            "Qty_IN": lang.t("stock_transactions.qty_in", fallback="Qty IN"),
            "IN_Type": lang.t("stock_transactions.in_type", fallback="IN Type"),
            "Qty_Out": lang.t("stock_transactions.qty_out", fallback="Qty OUT"),
            "Out_Type": lang.t("stock_transactions.out_type", fallback="Out Type"),
            "Third_Party": lang.t("stock_transactions.third_party", fallback="Third Party"),
            "End_User": lang.t("stock_transactions.end_user", fallback="End User"),
            "Discrepancy": lang.t("stock_transactions.discrepancy", fallback="Discrepancy"),
            "Remarks": lang.t("stock_transactions.remarks", fallback="Remarks"),
            "Movement_Type": lang.t("stock_transactions.movement_type", fallback="Movement Type")
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

    # ---------------- Data loading with localized display ----------------
    def load_transactions(self):
        """Load all transactions into treeview with localized display for selected fields."""
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
                # Localize selected columns (IN_Type, Remarks, Movement_Type)
                in_type_disp = self._to_display_in_type(row[11])
                movement_disp = self._to_display_movement(row[18])
                remarks_disp = self._to_display_remarks(row[17])

                values = (
                    row[0],            # Date
                    row[1],            # Time
                    row[2],            # unique_id (hidden)
                    row[3],            # code
                    row[4],            # Description
                    row[5],            # Expiry_date
                    row[6],            # Batch_Number
                    row[7],            # Scenario
                    row[8],            # Kit
                    row[9],            # Module
                    row[10],           # Qty_IN
                    in_type_disp,      # IN_Type (localized)
                    row[12],           # Qty_Out
                    row[13],           # Out_Type (left as-is; add mapping if needed)
                    row[14],           # Third_Party
                    row[15],           # End_User
                    row[16],           # Discrepancy
                    remarks_disp,      # Remarks (localized if mapped)
                    movement_disp      # Movement_Type (localized)
                )

                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                self.tree.insert("", "end", values=values, tags=(tag,))
            logging.debug("Successfully loaded transactions")
        except Exception as e:
            logging.error(f"Error loading transactions: {str(e)}")
            messagebox.showerror(
                lang.t("dialog_titles.error", fallback="Error"),
                lang.t("stock_transactions.load_error", fallback="Failed to load transactions: {error}").format(error=str(e)),
                parent=self
            )
        finally:
            cursor.close()
            conn.close()

    # ---------------- Search (supports localized movement type input) ----------------
    def search_transactions(self):
        """Search by code or movement type (supports localized input for movement type)."""
        query = (self.search_entry.get() or "").strip()
        if not query:
            self.load_transactions()
            return

        self.tree.delete(*self.tree.get_children())
        conn = connect_db()
        cursor = conn.cursor()

        try:
            # Convert localized movement type to canonical English (case-insensitive) for searching
            canonical_mt = lang.enum_to_canonical("stock_transactions.movement_types_map", query, fallback=query)

            cursor.execute("""
                SELECT 
                    Date, Time, unique_id, code, Description,
                    Expiry_date, Batch_Number, Scenario, Kit, Module,
                    Qty_IN, IN_Type, Qty_Out, Out_Type,
                    Third_Party, End_User, Discrepancy, Remarks, Movement_Type
                FROM stock_transactions
                WHERE code LIKE ?
                   OR Movement_Type LIKE ?
                   OR Movement_Type = ?
                ORDER BY Date DESC, Time DESC
            """, (f"%{query}%", f"%{canonical_mt}%", canonical_mt))

            for i, row in enumerate(cursor.fetchall()):
                in_type_disp = self._to_display_in_type(row[11])
                movement_disp = self._to_display_movement(row[18])
                remarks_disp = self._to_display_remarks(row[17])

                values = (
                    row[0], row[1], row[2], row[3], row[4],
                    row[5], row[6], row[7], row[8], row[9],
                    row[10], in_type_disp, row[12], row[13],
                    row[14], row[15], row[16], remarks_disp, movement_disp
                )

                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                self.tree.insert("", "end", values=values, tags=(tag,))
            logging.debug(f"Successfully searched transactions for query: {query}")
        except Exception as e:
            logging.error(f"Error searching transactions: {str(e)}")
            messagebox.showerror(
                lang.t("dialog_titles.error", fallback="Error"),
                lang.t("stock_transactions.search_error", fallback="Failed to search transactions: {error}").format(error=str(e)),
                parent=self
            )
        finally:
            cursor.close()
            conn.close()

    # ---------------- Export displayed data ----------------
    def export_to_excel(self):
        """Export displayed data to CSV."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not file_path:
            return

        try:
            with open(file_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                # Write header with translated labels
                writer.writerow([
                    lang.t("stock_transactions.date", fallback="Date"),
                    lang.t("stock_transactions.time", fallback="Time"),
                    lang.t("stock_transactions.unique_id", fallback="Unique ID"),
                    lang.t("stock_transactions.code", fallback="Code"),
                    lang.t("stock_transactions.description", fallback="Description"),
                    lang.t("stock_transactions.expiry_date", fallback="Expiry Date"),
                    lang.t("stock_transactions.batch_number", fallback="Batch Number"),
                    lang.t("stock_transactions.scenario", fallback="Scenario"),
                    lang.t("stock_transactions.kit", fallback="Kit"),
                    lang.t("stock_transactions.module", fallback="Module"),
                    lang.t("stock_transactions.qty_in", fallback="Qty IN"),
                    lang.t("stock_transactions.in_type", fallback="IN Type"),
                    lang.t("stock_transactions.qty_out", fallback="Qty OUT"),
                    lang.t("stock_transactions.out_type", fallback="Out Type"),
                    lang.t("stock_transactions.third_party", fallback="Third Party"),
                    lang.t("stock_transactions.end_user", fallback="End User"),
                    lang.t("stock_transactions.discrepancy", fallback="Discrepancy"),
                    lang.t("stock_transactions.remarks", fallback="Remarks"),
                    lang.t("stock_transactions.movement_type", fallback="Movement Type")
                ])
                # Write rows
                for item in self.tree.get_children():
                    writer.writerow(self.tree.item(item)["values"])
            logging.debug(f"Successfully exported transactions to {file_path}")
            messagebox.showinfo(
                lang.t("dialog_titles.success", fallback="Success"),
                lang.t("stock_transactions.export_success", fallback="Transactions exported successfully!"),
                parent=self
            )
        except Exception as e:
            logging.error(f"Error exporting transactions: {str(e)}")
            messagebox.showerror(
                lang.t("dialog_titles.error", fallback="Error"),
                lang.t("stock_transactions.export_error", fallback="Failed to export transactions: {error}").format(error=str(e)),
                parent=self
            )