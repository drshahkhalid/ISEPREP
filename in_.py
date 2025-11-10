import tkinter as tk
from tkinter import ttk, filedialog
from datetime import datetime
import sqlite3
import re
from calendar import monthrange
from language_manager import lang
from db import connect_db
from transaction_utils import log_transaction
from stock_data import StockData
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from popup_utils import custom_popup, custom_askyesno, custom_dialog
import os

# Unified popup utilities (consistent with receive_kit / dispatch_kit)
from popup_utils import custom_popup, custom_askyesno

# ---------------------- Utility / Parsing ---------------------- #
def parse_expiry(date_str):
    """
    Parse various date formats; for MM/YYYY or YYYY/MM return last day of month.
    Returns datetime.date or None.
    """
    if not date_str or str(date_str).lower() == 'none':
        return None
    date_str = date_str.strip()
    date_str = re.sub(r'[-/\s]+', '/', date_str)
    formats = [
        r'(\d{1,2})/(\d{1,2})/(\d{4})',  # DD/MM/YYYY
        r'(\d{4})/(\d{1,2})/(\d{1,2})',  # YYYY/MM/DD
        r'(\d{1,2})/(\d{4})',            # MM/YYYY
        r'(\d{4})/(\d{1,2})'             # YYYY/MM
    ]
    from calendar import monthrange as mr
    try:
        for fmt in formats:
            m = re.match(fmt, date_str)
            if not m:
                continue
            g = m.groups()
            if len(g) == 3:
                if fmt.startswith(r'(\d{1,2})'):  # DD/MM/YYYY
                    day, month, year = map(int, g)
                else:  # YYYY/MM/DD
                    year, month, day = map(int, g)
                return datetime(year, month, day).date()
            elif len(g) == 2:
                if fmt.startswith(r'(\d{1,2})'):
                    month, year = map(int, g)
                else:
                    year, month = map(int, g)
                _, last_day = mr(year, month)
                return datetime(year, month, last_day).date()
        return None
    except ValueError:
        return None

def get_active_designation(code):
    """
    Returns description in active language (fallback chain).
    """
    conn = connect_db()
    if conn is None:
        return lang.t("stock_in.no_description", "No Description")
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT designation, designation_en, designation_fr, designation_sp
            FROM items_list WHERE code=?
        """, (code,))
        row = cur.fetchone()
        if not row:
            return lang.t("stock_in.no_description", "No Description")
        lang_code = (lang.lang_code or "en").lower()
        mapping = {
            "en": "designation_en",
            "fr": "designation_fr",
            "es": "designation_sp",
            "sp": "designation_sp"
        }
        active_col = mapping.get(lang_code, "designation_en")
        if row[active_col]:
            return row[active_col]
        if row["designation_en"]:
            return row["designation_en"]
        return row["designation"] if row["designation"] else lang.t("stock_in.no_description", "No Description")
    finally:
        cur.close()
        conn.close()

def check_expiry_required(code):
    """
    Returns True if 'exp' found in items_list.remarks.
    """
    conn = connect_db()
    if conn is None:
        return False
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute("SELECT remarks FROM items_list WHERE code=?", (code,))
        row = cur.fetchone()
        return bool(row and row['remarks'] and 'exp' in row['remarks'].lower())
    finally:
        cur.close()
        conn.close()

def fetch_project_details():
    """
    Returns (project_name, project_code).
    """
    conn = connect_db()
    if conn is None:
        return lang.t("stock_in.unknown_project", "Unknown Project"), lang.t("stock_in.unknown_code", "Unknown Code")
    cur = conn.cursor()
    try:
        cur.execute("SELECT project_name, project_code FROM project_details LIMIT 1")
        row = cur.fetchone()
        return (row[0] if row and row[0] else lang.t("stock_in.unknown_project", "Unknown Project"),
                row[1] if row and row[1] else lang.t("stock_in.unknown_code", "Unknown Code"))
    finally:
        cur.close()
        conn.close()

def fetch_qty_needed(self, code, scenario_name, scenario_id, std_qty):
    """
    Compute qty_needed = qty_to_order_per_scenario - (session qty_in).
    Fallback to std_qty if no stored target.
    """
    conn = connect_db()
    if conn is None:
        return 0
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT COALESCE(qty_to_order_per_scenario, 0)
            FROM stock_data
            WHERE item=? AND scenario=?
        """, (code, scenario_name))
        row = cur.fetchone()
        base_needed = int(row[0]) if row else None
        if base_needed is None:
            base_needed = int(std_qty) if str(std_qty).isdigit() else 0
        session_in = 0
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if vals[0] == code and vals[2] == scenario_name:
                qin = vals[7]
                if qin and str(qin).isdigit():
                    session_in += int(qin)
        return max(base_needed - session_in, 0)
    finally:
        cur.close()
        conn.close()

# ---------------------- Main Class ---------------------- #
class StockIn(tk.Frame):
    def __init__(self, parent: tk.Widget, root: tk.Tk, role: str = "supervisor"):
        super().__init__(parent)
        self.root = root
        self.role = role.lower()
        self.editing_cell = None
        self.scenario_map = self.fetch_scenario_map()
        self.reverse_scenario_map = {v: k for k, v in self.scenario_map.items()}
        self.user_inputs = {}
        self.row_data = {}             # item_id -> {'unique_id','unique_id_2'}
        self.context_menu_row = None
        self.current_document_number = None  # store doc number for export
        self.pack(fill="both", expand=True)
        self.render_ui()

    # ---------- Unified popup wrappers ---------- #
    def show_error(self, msg_key_default, default_text, **fmt):
        custom_popup(self, lang.t("dialog_titles.error", "Error"),
                     lang.t(msg_key_default, default_text, **fmt), "error")

    def show_info(self, msg_key_default, default_text, **fmt):
        custom_popup(self, lang.t("dialog_titles.success", "Success"),
                     lang.t(msg_key_default, default_text, **fmt), "info")

    def ask_yes_no(self, msg_key_default, default_text, **fmt):
        return custom_askyesno(self, lang.t("dialog_titles.confirm", "Confirm"),
                               lang.t(msg_key_default, default_text, **fmt)) == "yes"

    # -------- Document Number Generation -------- #
    def generate_document_number(self, in_type_text: str) -> str:
        """
        Format: YYYY/MM/<PROJECT_CODE>/<ABBR>/<SERIAL>
        ABBR from mapping; dynamic fallback if unknown.
        SERIAL increments per prefix (4 digits).
        """
        project_name, project_code = fetch_project_details()
        project_code = (project_code or "PRJ").strip().upper()

        base_map = {
            "In MSF": "IMSF",
            "In Local Purchase": "ILP",
            "In from Quarantine": "IFQ",
            "In Donation": "IDN",
            "Return from End User": "IREU",
            "In Supply Non-MSF": "ISNM",
            "In Borrowing": "IBR",
            "In Return of Loan": "IRL",
            "In Correction of Previous Transaction": "ICOR"
        }
        raw = (in_type_text or "").strip()
        import re
        norm = re.sub(r'[^a-z0-9]+', '', raw.lower())
        abbr = None
        for k, v in base_map.items():
            if re.sub(r'[^a-z0-9]+', '', k.lower()) == norm:
                abbr = v
                break
        if not abbr:
            tokens = re.split(r'\s+', raw.upper())
            stop = {"OF", "FROM", "THE", "AND", "DE", "DU", "DES", "LA", "LE", "LES"}
            parts = []
            for t in tokens:
                if not t or t in stop:
                    continue
                if t == "MSF":
                    parts.append("MSF")
                else:
                    parts.append(t[0])
            if not parts:
                abbr = (raw[:4].upper() or "DOC").replace(" ", "")
            else:
                abbr = "".join(parts)
            if len(abbr) > 8:
                abbr = abbr[:8]

        now = datetime.now()
        prefix = f"{now.year:04d}/{now.month:02d}/{project_code}/{abbr}"
        serial = 1
        conn = connect_db()
        if conn:
            cur = conn.cursor()
            try:
                cur.execute("""
                    SELECT document_number
                      FROM stock_transactions
                     WHERE document_number LIKE ?
                     ORDER BY document_number DESC
                     LIMIT 1
                """, (prefix + "/%",))
                row = cur.fetchone()
                if row and row[0]:
                    tail = row[0].rsplit("/", 1)[-1]
                    if tail.isdigit():
                        serial = int(tail) + 1
            except Exception:
                pass
            finally:
                cur.close()
                conn.close()
        doc = f"{prefix}/{serial:04d}"
        self.current_document_number = doc
        return doc

    # -------- Data Fetch Helpers -------- #
    def fetch_scenario_map(self):
        conn = connect_db()
        if conn is None:
            return {}
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("SELECT scenario_id, name FROM scenarios ORDER BY name")
            return {str(r['scenario_id']): r['name'] for r in cur.fetchall()}
        finally:
            cur.close()
            conn.close()

    def fetch_search_results(self, query):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        lang_code = (lang.lang_code or "en").lower()
        mapping = {
            "en": "designation_en",
            "fr": "designation_fr",
            "es": "designation_sp",
            "sp": "designation_sp"
        }
        active_des = mapping.get(lang_code, "designation_en")
        sql = f"""
            SELECT DISTINCT c.code
              FROM compositions c
              LEFT JOIN items_list i ON c.code = i.code
             WHERE c.quantity > 0
               AND (
                    LOWER(c.code) LIKE ?
                 OR LOWER(i.{active_des}) LIKE ?
                 OR (i.{active_des} IS NULL AND LOWER(i.designation_en) LIKE ?)
                 OR (i.{active_des} IS NULL AND i.designation_en IS NULL AND LOWER(i.designation) LIKE ?)
               )
        """
        params = [f"%{query.lower()}%"] * 4
        if self.scenario_var.get() != lang.t("stock_in.all_scenarios", "All Scenarios"):
            sql += " AND c.scenario_id = ?"
            params.append(self.reverse_scenario_map.get(self.scenario_var.get()))
        sql += " ORDER BY c.code"
        try:
            cur.execute(sql, params)
            rows = cur.fetchall()
            return [{'code': r['code'], 'description': get_active_designation(r['code'])} for r in rows]
        finally:
            cur.close()
            conn.close()

    def fetch_all_compositions(self):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        base = """
            SELECT c.code, c.quantity, s.name AS scenario_name, c.scenario_id
              FROM compositions c
              LEFT JOIN scenarios s ON c.scenario_id = s.scenario_id
             WHERE c.quantity > 0
        """
        params = []
        if self.scenario_var.get() != lang.t("stock_in.all_scenarios", "All Scenarios"):
            base += " AND c.scenario_id = ?"
            params.append(self.reverse_scenario_map.get(self.scenario_var.get()))
        base += " ORDER BY c.code, s.name"
        try:
            cur.execute(base, params)
            rows = cur.fetchall()
            out = []
            for r in rows:
                code = r['code']
                scenario_name = r['scenario_name']
                out.append({
                    'unique_id': f"{scenario_name}/None/-----/{code}",
                    'unique_id_2': f"{scenario_name}/None/-----/{code}",
                    'scenario_id': r['scenario_id'],
                    'scenario_name': scenario_name,
                    'kit_code': '-----',
                    'module_code': '-----',
                    'code': code,
                    'quantity': r['quantity'],
                    'description': get_active_designation(code)
                })
            return out
        finally:
            cur.close()
            conn.close()

    def fetch_compositions_for_code(self, code):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        sql = """
            SELECT c.code, c.quantity, s.name AS scenario_name, c.scenario_id
              FROM compositions c
              LEFT JOIN scenarios s ON c.scenario_id = s.scenario_id
             WHERE c.quantity > 0 AND c.code = ?
        """
        params = [code]
        if self.scenario_var.get() != lang.t("stock_in.all_scenarios", "All Scenarios"):
            sql += " AND c.scenario_id = ?"
            params.append(self.reverse_scenario_map.get(self.scenario_var.get()))
        sql += " ORDER BY c.code, s.name"
        try:
            cur.execute(sql, params)
            rows = cur.fetchall()
            out = []
            for r in rows:
                out.append({
                    'unique_id': f"{r['scenario_name']}/None/-----/{r['code']}",
                    'unique_id_2': f"{r['scenario_name']}/None/-----/{r['code']}",
                    'scenario_id': r['scenario_id'],
                    'scenario_name': r['scenario_name'],
                    'kit_code': '-----',
                    'module_code': '-----',
                    'code': r['code'],
                    'quantity': r['quantity'],
                    'description': get_active_designation(r['code'])
                })
            return out
        finally:
            cur.close()
            conn.close()

    def fetch_third_parties(self):
        conn = connect_db()
        if conn is None:
            return []
        cur = conn.cursor()
        try:
            cur.execute("SELECT name FROM third_parties ORDER BY name")
            return [r[0] for r in cur.fetchall()]
        finally:
            cur.close()
            conn.close()

    def fetch_end_users(self):
        conn = connect_db()
        if conn is None:
            return []
        cur = conn.cursor()
        try:
            cur.execute("SELECT name FROM end_users ORDER BY name")
            return [r[0] for r in cur.fetchall()]
        finally:
            cur.close()
            conn.close()

    # -------- UI Rendering -------- #
    def render_ui(self):
        for w in self.winfo_children():
            try:
                w.destroy()
            except:
                pass

        bg = "#F0F4F8"
        self.configure(bg=bg)

        tk.Label(self, text=lang.t("stock_in.title", "Stock In"),
                 font=("Helvetica", 20, "bold"), bg=bg).pack(pady=10)

        btn_frame = tk.Frame(self, bg=bg)
        btn_frame.pack(pady=5, fill="x")
        can_edit = self.role in ["admin", "manager"]
        tk.Button(btn_frame, text=lang.t("stock_in.add_button", "Add Stock"),
                  bg="#27AE60", fg="white", command=self.save_all,
                  activebackground="#1E874B",
                  state="normal" if can_edit else "disabled").pack(side="left", padx=5)
        tk.Button(btn_frame, text=lang.t("stock_in.clear_all", "Clear All"),
                  bg="#7F8C8D", fg="white", activebackground="#666E70",
                  command=self.clear_form).pack(side="left", padx=5)
        tk.Button(btn_frame, text=lang.t("stock_in.export", "Export"),
                  bg="#2980B9", fg="white", activebackground="#1F6390",
                  command=self.export_data).pack(side="left", padx=5)

        filter_frame = tk.Frame(self, bg=bg)
        filter_frame.pack(pady=5, fill="x")
        tk.Label(filter_frame, text=lang.t("stock_in.scenario", "Scenario:"),
                 bg=bg).grid(row=0, column=0, padx=5, sticky="w")
        self.scenario_var = tk.StringVar(value=lang.t("stock_in.all_scenarios", "All Scenarios"))
        scenarios = list(self.scenario_map.values()) + [lang.t("stock_in.all_scenarios", "All Scenarios")]
        self.scenario_cb = ttk.Combobox(filter_frame, textvariable=self.scenario_var,
                                        values=scenarios, state="readonly", width=30)
        self.scenario_cb.grid(row=0, column=1, padx=5, pady=5)
        self.scenario_cb.bind("<<ComboboxSelected>>", self.on_scenario_selected)

        type_frame = tk.Frame(self, bg=bg)
        type_frame.pack(pady=5, fill="x")
        tk.Label(type_frame, text=lang.t("stock_in.in_type", "IN Type:"), bg=bg).grid(row=0, column=0, padx=5, sticky="w")
        self.trans_type_var = tk.StringVar()
        self.trans_type_cb = ttk.Combobox(type_frame, textvariable=self.trans_type_var,
                                          values=[
                                              lang.t("stock_in.in_msf", "In MSF"),
                                              lang.t("stock_in.in_local_purchase", "In Local Purchase"),
                                              lang.t("stock_in.in_from_quarantine", "In from Quarantine"),
                                              lang.t("stock_in.in_donation", "In Donation"),
                                              lang.t("stock_in.return_from_end_user", "Return from End User"),
                                              lang.t("stock_in.in_supply_non_msf", "In Supply Non-MSF"),
                                              lang.t("stock_in.in_borrowing", "In Borrowing"),
                                              lang.t("stock_in.in_return_loan", "In Return of Loan"),
                                              lang.t("stock_in.in_correction", "In Correction of Previous Transaction")
                                          ],
                                          state="readonly", width=30)
        self.trans_type_cb.grid(row=0, column=1, padx=5, pady=5)
        self.trans_type_cb.bind("<<ComboboxSelected>>", self.update_dropdown_visibility)

        tk.Label(type_frame, text=lang.t("stock_in.end_user", "End User:"), bg=bg).grid(row=0, column=2, padx=5, sticky="w")
        self.end_user_var = tk.StringVar()
        self.end_user_cb = ttk.Combobox(type_frame, textvariable=self.end_user_var, state="disabled", width=30)
        self.end_user_cb['values'] = self.fetch_end_users()
        self.end_user_cb.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(type_frame, text=lang.t("stock_in.third_party", "Third Party:"), bg=bg).grid(row=0, column=4, padx=5, sticky="w")
        self.third_party_var = tk.StringVar()
        self.third_party_cb = ttk.Combobox(type_frame, textvariable=self.third_party_var, state="disabled", width=30)
        self.third_party_cb['values'] = self.fetch_third_parties()
        self.third_party_cb.grid(row=0, column=5, padx=5, pady=5)

        tk.Label(type_frame, text=lang.t("stock_in.remarks", "Remarks:"), bg=bg).grid(row=0, column=6, padx=5, sticky="w")
        self.remarks_entry = tk.Entry(type_frame, width=40, state="disabled")
        self.remarks_entry.grid(row=0, column=7, padx=5, pady=5)

        search_frame = tk.Frame(self, bg=bg)
        search_frame.pack(pady=10, fill="x")
        tk.Label(search_frame, text=lang.t("stock_in.item_code", "Item Code"), bg=bg).grid(row=0, column=0, padx=5, sticky="w")
        self.code_entry = tk.Entry(search_frame)
        self.code_entry.grid(row=0, column=1, padx=5, pady=5)
        self.code_entry.bind("<KeyRelease>", self.search_items)
        self.code_entry.bind("<Return>", self.select_first_result)
        tk.Button(search_frame, text=lang.t("stock_in.clear_search", "Clear Search"),
                  bg="#7F8C8D", fg="white", activebackground="#666E70",
                  command=self.clear_search).grid(row=0, column=2, padx=5, pady=5)
        self.search_listbox = tk.Listbox(search_frame, height=5, width=80)
        self.search_listbox.grid(row=0, column=3, padx=5, pady=5, sticky="we")
        self.search_listbox.bind("<<ListboxSelect>>", self.fill_code_from_search)

        tree_frame = tk.Frame(self, bg=bg)
        tree_frame.pack(expand=True, fill="both", pady=10)
        self.cols = ("code", "description", "scenario_name", "kit_code", "module_code",
                     "std_qty", "qty_needed", "qty_in", "expiry_date", "batch_no")
        self.tree = ttk.Treeview(tree_frame, columns=self.cols, show="headings", height=18)
        self.tree.tag_configure("light_red", background="#FF9999")

        self.headers = {
            "code": lang.t("stock_in.code", "Code"),
            "description": lang.t("stock_in.description", "Description"),
            "scenario_name": lang.t("stock_in.scenario_name", "Scenario Name"),
            "kit_code": lang.t("stock_in.kit_code", "Kit"),
            "module_code": lang.t("stock_in.module_code", "Module"),
            "std_qty": lang.t("stock_in.std_qty", "Std Qty"),
            "qty_needed": lang.t("stock_in.qty_needed", "Qty Needed"),
            "qty_in": lang.t("stock_in.qty_in", "Qty In"),
            "expiry_date": lang.t("stock_in.expiry_date", "Expiry Date"),
            "batch_no": lang.t("stock_in.batch_no", "Batch No")
        }
        self.widths = {
            "code": 100, "description": 335, "scenario_name": 120, "kit_code": 120,
            "module_code": 120, "std_qty": 80, "qty_needed": 90, "qty_in": 90,
            "expiry_date": 110, "batch_no": 120
        }
        for c in self.cols:
            self.tree.heading(c, text=self.headers[c])
            self.tree.column(c, width=self.widths[c], stretch=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(
            label=lang.t("stock_in.add_another_line", "Add another line for this item"),
            command=self.add_another_line
        )
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Double-1>", self.start_edit)

        self.status_var = tk.StringVar(value=lang.t("stock_in.ready", "Ready"))
        tk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w",
                 bg=bg).pack(fill="x", pady=(5, 0))

        self.populate_table()

    # -------- Context Menu -------- #
    def show_context_menu(self, event):
        self.context_menu_row = self.tree.identify_row(event.y)
        if self.context_menu_row:
            self.tree.selection_set(self.context_menu_row)
            self.context_menu.post(event.x_root, event.y_root)

    def add_another_line(self):
        if not self.context_menu_row:
            return
        vals = self.tree.item(self.context_menu_row, "values")
        item_id = self.context_menu_row
        unique_id_2 = self.row_data[item_id]['unique_id_2']
        code = vals[0]
        description = vals[1]
        scenario_name = vals[2]
        kit_code = vals[3]
        module_code = vals[4]
        std_qty = int(vals[5]) if vals[5] and str(vals[5]).isdigit() else 0
        scenario_id = self.reverse_scenario_map.get(scenario_name)
        qty_needed = fetch_qty_needed(self, code, scenario_name, scenario_id, std_qty)
        new_unique_id = f"{scenario_name}/None/-----/{code}/None"
        input_key = f"{scenario_name}/{code}/{std_qty}"
        user_input = self.user_inputs.get(input_key, {"qty_in": "", "expiry_date": "", "batch_no": ""})
        index = self.tree.index(self.context_menu_row) + 1
        new_item_id = self.tree.insert("", index, values=(
            code, description, scenario_name, kit_code, module_code,
            std_qty, qty_needed, user_input["qty_in"], user_input["expiry_date"], user_input["batch_no"]
        ))
        self.row_data[new_item_id] = {'unique_id': new_unique_id, 'unique_id_2': unique_id_2}
        self.status_var.set(lang.t("stock_in.added_line", "Added new line for {code}").format(code=code))
        for item_id in self.tree.get_children():
            item_vals = self.tree.item(item_id, "values")
            if item_vals[0] == code and item_vals[2] == scenario_name:
                item_std_qty = int(item_vals[5]) if item_vals[5] and str(item_vals[5]).isdigit() else 0
                new_qty_needed = fetch_qty_needed(self, code, scenario_name, scenario_id, item_std_qty)
                self.tree.set(item_id, "qty_needed", new_qty_needed)

    def update_tree_columns(self, event=None):
        for c in self.cols:
            self.tree.heading(c, text=self.headers[c])
            self.tree.column(c, width=self.widths[c], stretch=True)
        self.tree.bind("<Double-1>", self.start_edit)

    def on_scenario_selected(self, event=None):
        self.populate_table()

    def update_dropdown_visibility(self, event=None):
        ttype = self.trans_type_var.get()
        self.end_user_cb.config(state="disabled")
        self.third_party_cb.config(state="disabled")
        self.remarks_entry.config(state="disabled")
        self.end_user_var.set("")
        self.third_party_var.set("")
        self.remarks_entry.delete(0, tk.END)
        if ttype in [
            lang.t("stock_in.in_donation", "In Donation"),
            lang.t("stock_in.in_borrowing", "In Borrowing"),
            lang.t("stock_in.in_return_loan", "In Return of Loan")
        ]:
            self.third_party_cb.config(state="readonly")
        elif ttype == lang.t("stock_in.return_from_end_user", "Return from End User"):
            self.end_user_cb.config(state="readonly")
        elif ttype == lang.t("stock_in.in_correction", "In Correction of Previous Transaction"):
            self.remarks_entry.config(state="normal")

    def search_items(self, event=None):
        query = self.code_entry.get().strip()
        self.search_listbox.delete(0, tk.END)
        if not query:
            self.populate_table()
            return
        results = self.fetch_search_results(query)
        for r in results:
            display = f"{r['code']} - {r['description'] or lang.t('stock_in.no_description', 'No Description')}"
            self.search_listbox.insert(tk.END, display)
        self.status_var.set(
            lang.t("stock_in.found_items", "Found {size} items").format(size=self.search_listbox.size())
        )

    def clear_search(self):
        self.code_entry.delete(0, tk.END)
        self.search_listbox.delete(0, tk.END)
        self.populate_table()

    def select_first_result(self, event=None):
        if self.search_listbox.size() > 0:
            self.search_listbox.selection_set(0)
            self.fill_code_from_search()

    def fill_code_from_search(self, event=None):
        sel = self.search_listbox.curselection()
        if not sel:
            return
        code = self.search_listbox.get(sel[0]).split(" - ")[0]
        self.code_entry.delete(0, tk.END)
        self.code_entry.insert(0, code)
        self.search_listbox.delete(0, tk.END)
        self.save_user_inputs()
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.status_var.set(lang.t("stock_in.loading", "Loading..."))
        self.update_idletasks()
        data = self.fetch_compositions_for_code(code)
        if not data:
            self.status_var.set(
                lang.t("stock_in.no_items", "No items found for code {code}").format(code=code)
            )
        else:
            for row in data:
                unique_id = row['unique_id']
                unique_id_2 = row['unique_id_2']
                scenario_name = row['scenario_name']
                kit_code = row['kit_code']
                module_code = row['module_code']
                code_ = row['code']
                description = row['description']
                std_qty = int(row['quantity']) if row['quantity'] and str(row['quantity']).isdigit() else 0
                qty_needed = fetch_qty_needed(self, code_, scenario_name, row['scenario_id'], std_qty)
                input_key = f"{scenario_name}/{code_}/{std_qty}"
                user_input = self.user_inputs.get(input_key, {"qty_in": "", "expiry_date": "", "batch_no": ""})
                values = (
                    code_, description, scenario_name, kit_code, module_code,
                    std_qty, qty_needed, user_input["qty_in"], user_input["expiry_date"], user_input["batch_no"]
                )
                item_id = self.tree.insert("", "end", values=values)
                self.row_data[item_id] = {'unique_id': unique_id, 'unique_id_2': unique_id_2}
                if user_input["qty_in"] and user_input["qty_in"].isdigit() and check_expiry_required(code_):
                    if not self.validate_expiry_for_save(code_, user_input["qty_in"], user_input["expiry_date"]):
                        self.tree.item(item_id, tags=("light_red",))
            self.status_var.set(
                lang.t("stock_in.loaded", "Loaded {count} records for code {code}")
                .format(count=len(self.tree.get_children()), code=code)
            )

    def save_user_inputs(self):
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            scenario_name = vals[2]
            code = vals[0]
            std_qty = vals[5]
            key = f"{scenario_name}/{code}/{std_qty}"
            self.user_inputs[key] = {
                "qty_in": vals[7],
                "expiry_date": vals[8],
                "batch_no": vals[9]
            }

    def populate_table(self, event=None):
        self.save_user_inputs()
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        data = self.fetch_all_compositions()
        for row in data:
            unique_id = row['unique_id']
            unique_id_2 = row['unique_id_2']
            scenario_name = row['scenario_name']
            kit_code = row['kit_code']
            module_code = row['module_code']
            code = row['code']
            description = row['description']
            std_qty = int(row['quantity']) if row['quantity'] and str(row['quantity']).isdigit() else 0
            qty_needed = fetch_qty_needed(self, code, scenario_name, row['scenario_id'], std_qty)
            key = f"{scenario_name}/{code}/{std_qty}"
            user_input = self.user_inputs.get(key, {"qty_in": "", "expiry_date": "", "batch_no": ""})
            values = (
                code, description, scenario_name, kit_code, module_code,
                std_qty, qty_needed, user_input["qty_in"], user_input["expiry_date"], user_input["batch_no"]
            )
            item_id = self.tree.insert("", "end", values=values)
            self.row_data[item_id] = {'unique_id': unique_id, 'unique_id_2': unique_id_2}
            if user_input["qty_in"] and user_input["qty_in"].isdigit() and check_expiry_required(code):
                if not self.validate_expiry_for_save(code, user_input["qty_in"], user_input["expiry_date"]):
                    self.tree.item(item_id, tags=("light_red",))
        self.status_var.set(
            lang.t("stock_in.loaded", "Loaded {count} records").format(count=len(self.tree.get_children()))
        )

    def clear_form(self):
        self.code_entry.delete(0, tk.END)
        self.search_listbox.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.trans_type_var.set("")
        self.end_user_var.set("")
        self.third_party_var.set("")
        self.remarks_entry.delete(0, tk.END)
        self.end_user_cb.config(state="disabled")
        self.third_party_cb.config(state="disabled")
        self.remarks_entry.config(state="disabled")
        self.user_inputs = {}
        self.current_document_number = None
        self.populate_table()
        self.status_var.set(lang.t("stock_in.ready", "Ready"))

    def validate_expiry_for_save(self, code, qty_in, expiry_date):
        if not qty_in:
            return True
        if isinstance(qty_in, str) and not qty_in.isdigit():
            return True
        if not check_expiry_required(code):
            return True
        parsed_expiry = parse_expiry(expiry_date) if expiry_date else None
        current_date = datetime.now().date()
        if not parsed_expiry or parsed_expiry <= current_date:
            return False
        return True

    # -------- Editing (Cell) -------- #
    def start_edit(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id or not col:
            return
        col_index = int(col.replace("#", "")) - 1
        if col_index not in [7, 8, 9]:
            return
        if self.editing_cell:
            self.editing_cell.destroy()
            self.editing_cell = None
        bbox = self.tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox
        value = self.tree.set(row_id, self.cols[col_index])
        code = self.tree.set(row_id, "code")
        scenario_name = self.tree.set(row_id, "scenario_name")
        scenario_id = self.reverse_scenario_map.get(scenario_name)
        try:
            std_qty = int(self.tree.set(row_id, "std_qty"))
        except:
            std_qty = 0
        expiry_required = check_expiry_required(code)

        entry = tk.Entry(self.tree, font=("Helvetica", 10))
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.focus()
        self.editing_cell = entry

        def save_edit(event=None, row_id=row_id, col_index=col_index, code=code,
                      scenario_name=scenario_name, scenario_id=scenario_id,
                      std_qty=std_qty, expiry_required=expiry_required):
            if row_id not in self.tree.get_children():
                entry.destroy()
                self.editing_cell = None
                return
            new_val = entry.get().strip()

            if col_index == 7:
                if new_val and not new_val.isdigit():
                    custom_popup(self, lang.t("dialog_titles.error", "Error"),
                                 lang.t("stock_in.invalid_qty", "Qty In must be an integer"), "error")
                    entry.focus_set()
                    return
            if col_index == 8:
                if new_val:
                    parsed = parse_expiry(new_val)
                    if not parsed:
                        custom_popup(self, lang.t("dialog_titles.error", "Error"),
                                     lang.t("stock_in.invalid_expiry_format", "Invalid expiry date format"), "error")
                        entry.focus_set()
                        return
                    if parsed <= datetime.now().date():
                        custom_popup(self, lang.t("dialog_titles.error", "Error"),
                                     lang.t("stock_in.expiry_future", "Expiry date must be in the future"), "error")
                        entry.focus_set()
                        return
                    new_val = parsed.strftime("%Y-%m-%d")
            if col_index == 9:
                if new_val and len(new_val) > 30:
                    custom_popup(self, lang.t("dialog_titles.error", "Error"),
                                 lang.t("stock_in.batch_no_length", "Batch No must be 30 characters or less"), "error")
                    entry.focus_set()
                    return

            self.tree.set(row_id, self.cols[col_index], new_val)
            if col_index == 8:
                unique_id_2_base = self.row_data[row_id]['unique_id_2'].rsplit('/', 1)[0]
                new_unique_id_2 = f"{unique_id_2_base}/{new_val or 'None'}"
                self.row_data[row_id]['unique_id_2'] = new_unique_id_2
                self.row_data[row_id]['unique_id'] = new_unique_id_2

            qty_in = self.tree.set(row_id, "qty_in")
            expiry_date = self.tree.set(row_id, "expiry_date")
            if qty_in and qty_in.isdigit() and expiry_required and not self.validate_expiry_for_save(code, qty_in, expiry_date):
                self.tree.item(row_id, tags=("light_red",))
            else:
                self.tree.item(row_id, tags=())

            scenario_name2 = self.tree.set(row_id, "scenario_name")
            code2 = self.tree.set(row_id, "code")
            std_qty2 = int(self.tree.set(row_id, "std_qty")) if self.tree.set(row_id, "std_qty") and self.tree.set(row_id, "std_qty").isdigit() else 0
            input_key = f"{scenario_name2}/{code2}/{std_qty2}"
            self.user_inputs[input_key] = {
                "qty_in": self.tree.set(row_id, "qty_in"),
                "expiry_date": self.tree.set(row_id, "expiry_date"),
                "batch_no": self.tree.set(row_id, "batch_no")
            }

            for item_id in self.tree.get_children():
                item_vals = self.tree.item(item_id, "values")
                if item_vals[0] == code2 and item_vals[2] == scenario_name2:
                    item_std_qty = int(item_vals[5]) if item_vals[5] and str(item_vals[5]).isdigit() else 0
                    new_qty_needed = fetch_qty_needed(self, code2, scenario_name2, scenario_id, item_std_qty)
                    self.tree.set(item_id, "qty_needed", new_qty_needed)

            entry.destroy()
            self.editing_cell = None
            self.tree.update_idletasks()
            if event and event.keysym == "Tab":
                self.move_to_next_editable_cell(row_id, col_index)

        entry.bind("<Return>", save_edit)
        entry.bind("<Tab>", lambda e: save_edit(e, row_id, col_index, code, scenario_name, scenario_id, std_qty, expiry_required))
        entry.bind("<FocusOut>", lambda e: save_edit(e, row_id, col_index, code, scenario_name, scenario_id, std_qty, expiry_required))

    def move_to_next_editable_cell(self, current_row, current_col_index):
        editable_columns = [7, 8, 9]
        rows = self.tree.get_children()
        if not rows:
            return
        current_row_index = list(rows).index(current_row)
        for col_index in editable_columns:
            if col_index > current_col_index:
                if self.start_edit_cell(current_row, col_index):
                    return
        next_row_index = (current_row_index + 1) % len(rows)
        next_row = rows[next_row_index]
        self.start_edit_cell(next_row, editable_columns[0])

    def start_edit_cell(self, row_id, col_index):
        col = f"#{col_index + 1}"
        bbox = self.tree.bbox(row_id, col)
        if not bbox:
            return False
        x, y, width, height = bbox
        value = self.tree.set(row_id, self.cols[col_index])
        code = self.tree.set(row_id, "code")
        scenario_name = self.tree.set(row_id, "scenario_name")
        scenario_id = self.reverse_scenario_map.get(scenario_name)
        try:
            std_qty = int(self.tree.set(row_id, "std_qty"))
        except:
            std_qty = 0
        expiry_required = check_expiry_required(code)

        if self.editing_cell:
            self.editing_cell.destroy()
            self.editing_cell = None

        entry = tk.Entry(self.tree, font=("Helvetica", 10))
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, value)
        entry.focus()
        self.editing_cell = entry

        def save_edit(event=None, row_id=row_id, col_index=col_index, code=code,
                      scenario_name=scenario_name, scenario_id=scenario_id,
                      std_qty=std_qty, expiry_required=expiry_required):
            if row_id not in self.tree.get_children():
                entry.destroy()
                self.editing_cell = None
                return
            new_val = entry.get().strip()
            if col_index == 7 and new_val and not new_val.isdigit():
                custom_popup(self, lang.t("dialog_titles.error", "Error"),
                             lang.t("stock_in.invalid_qty", "Qty In must be an integer"), "error")
                entry.focus_set()
                return
            if col_index == 8 and new_val:
                parsed = parse_expiry(new_val)
                if not parsed:
                    custom_popup(self, lang.t("dialog_titles.error", "Error"),
                                 lang.t("stock_in.invalid_expiry_format", "Invalid expiry date format"), "error")
                    entry.focus_set()
                    return
                if parsed <= datetime.now().date():
                    custom_popup(self, lang.t("dialog_titles.error", "Error"),
                                 lang.t("stock_in.expiry_future", "Expiry date must be in the future"), "error")
                    entry.focus_set()
                    return
                new_val = parsed.strftime("%Y-%m-%d")
            if col_index == 9 and new_val and len(new_val) > 30:
                custom_popup(self, lang.t("dialog_titles.error", "Error"),
                             lang.t("stock_in.batch_no_length", "Batch No must be 30 characters or less"), "error")
                entry.focus_set()
                return
            self.tree.set(row_id, self.cols[col_index], new_val)
            if col_index == 8:
                unique_id_2_base = self.row_data[row_id]['unique_id_2'].rsplit('/', 1)[0]
                new_unique_id_2 = f"{unique_id_2_base}/{new_val or 'None'}"
                self.row_data[row_id]['unique_id_2'] = new_unique_id_2
                self.row_data[row_id]['unique_id'] = new_unique_id_2

            qty_in = self.tree.set(row_id, "qty_in")
            expiry_date = self.tree.set(row_id, "expiry_date")
            if qty_in and qty_in.isdigit() and expiry_required and not self.validate_expiry_for_save(code, qty_in, expiry_date):
                self.tree.item(row_id, tags=("light_red",))
            else:
                self.tree.item(row_id, tags=())

            scenario_name2 = self.tree.set(row_id, "scenario_name")
            code2 = self.tree.set(row_id, "code")
            std_qty2 = int(self.tree.set(row_id, "std_qty")) if self.tree.set(row_id, "std_qty") and self.tree.set(row_id, "std_qty").isdigit() else 0
            input_key = f"{scenario_name2}/{code2}/{std_qty2}"
            self.user_inputs[input_key] = {
                "qty_in": self.tree.set(row_id, "qty_in"),
                "expiry_date": self.tree.set(row_id, "expiry_date"),
                "batch_no": self.tree.set(row_id, "batch_no")
            }
            for item_id in self.tree.get_children():
                item_vals = self.tree.item(item_id, "values")
                if item_vals[0] == code2 and item_vals[2] == scenario_name2:
                    item_std_qty = int(item_vals[5]) if item_vals[5] and str(item_vals[5]).isdigit() else 0
                    new_qty_needed = fetch_qty_needed(self, code2, scenario_name2, scenario_id, item_std_qty)
                    self.tree.set(item_id, "qty_needed", new_qty_needed)
            entry.destroy()
            self.editing_cell = None
            if event and event.keysym == "Tab":
                self.move_to_next_editable_cell(row_id, col_index)
        entry.bind("<Return>", save_edit)
        entry.bind("<Tab>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        return True

    # -------- Save / Persist -------- #
    def save_all(self):
        if self.role.lower() not in ["admin", "manager"]:
            self.show_error("stock_in.no_permission", "Only admin or manager roles can save changes.")
            return
        rows = self.tree.get_children()
        if not rows:
            self.show_error("stock_in.no_rows", "No rows to save.")
            return
        ttype = self.trans_type_var.get()
        remarks = self.remarks_entry.get().strip()
        end_user = self.end_user_var.get().strip()
        third_party = self.third_party_var.get().strip()

        if not ttype:
            self.show_error("stock_in.no_in_type", "IN Type is required.")
            return
        if ttype == lang.t("stock_in.in_donation", "In Donation"):
            if not third_party or third_party not in self.third_party_cb['values']:
                self.show_error("stock_in.invalid_third_party", "A valid Third Party must be selected for In Donation.")
                return
        if ttype == lang.t("stock_in.return_from_end_user", "Return from End User"):
            if not end_user or end_user not in self.end_user_cb['values']:
                self.show_error("stock_in.invalid_end_user", "A valid End User must be selected for Return from End User.")
                return
        if ttype == lang.t("stock_in.in_correction", "In Correction of Previous Transaction"):
            if len(remarks) < 10 or len(remarks) > 300:
                self.show_error("stock_in.remarks_length", "Remarks must be between 10 and 300 characters for In Correction of Previous Transaction.")
                return

        # Generate Document Number
        doc_number = self.generate_document_number(ttype)
        self.status_var.set(
            lang.t("stock_in.generating_document", "Generated Document Number: {doc}")
               .format(doc=doc_number)
        )

        invalid_items = []
        exported_rows = []
        for iid in rows:
            vals = self.tree.item(iid, "values")
            code = vals[0]
            description = vals[1]
            scenario_name = vals[2]
            kit_code = vals[3]
            module_code = vals[4]
            std_qty = vals[5]
            qty_needed = vals[6]
            qty_in = vals[7]
            expiry_date = vals[8]
            batch_no = vals[9]

            if not qty_in or not qty_in.isdigit():
                continue
            qty_in_int = int(qty_in)
            parsed_exp = parse_expiry(expiry_date)
            expiry_fmt = parsed_exp.strftime("%Y-%m-%d") if parsed_exp else None

            if not self.validate_expiry_for_save(code, qty_in_int, expiry_fmt):
                invalid_items.append(code)
                self.tree.item(iid, tags=("light_red",))
                continue

            exp_part = expiry_fmt or "None"
            six_layer_unique_id = f"{scenario_name}/{kit_code if kit_code != '-----' else 'None'}/{module_code if module_code != '-----' else 'None'}/{code}/{std_qty}/{exp_part}"

            try:
                # log_transaction now supports document_number if DB column exists
                log_transaction(
                    unique_id=six_layer_unique_id,
                    code=code,
                    Description=description,
                    Expiry_date=expiry_fmt,
                    Batch_Number=batch_no,
                    Scenario=scenario_name,
                    Kit=kit_code if kit_code != "-----" else None,
                    Module=module_code if module_code != "-----" else None,
                    Qty_IN=qty_in_int,
                    IN_Type=ttype,
                    Third_Party=third_party if third_party else None,
                    End_User=end_user if end_user else None,
                    Remarks=remarks,
                    Movement_Type="stock_in",
                    document_number=doc_number
                )

                StockData.add_or_update(unique_id=six_layer_unique_id, qty_in=qty_in_int, qty_out=0, exp_date=expiry_fmt)
                exported_rows.append({
                    'code': code,
                    'description': description,
                    'scenario_name': scenario_name,
                    'kit_code': kit_code,
                    'module_code': module_code,
                    'std_qty': std_qty,
                    'qty_needed': qty_needed,
                    'qty_in': qty_in_int,
                    'expiry_date': expiry_fmt,
                    'batch_no': batch_no
                })
            except Exception as e:
                custom_popup(self, lang.t("dialog_titles.error", "Error"),
                             lang.t("stock_in.save_failed", "Failed to save row: {error}").format(error=str(e)), "error")
                continue

        if invalid_items:
            self.show_error("stock_in.invalid_expiry",
                            "Valid future expiry dates are required for items: {items}",
                            items=', '.join(invalid_items))
            return

        self.show_info("stock_in.saved", "Stock IN saved successfully.")
        # Offer export
        if exported_rows and self.ask_yes_no("stock_in.save_excel_prompt", "Do you want to save the stock issuance to Excel?"):
            self.export_data(exported_rows)

        self.clear_form()

    # -------- Export -------- #
    def export_data(self, rows_to_export=None):
        try:
            default_dir = "D:/ISEPREP"
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)

            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ttype = self.trans_type_var.get() or lang.t("stock_in.unknown", "Unknown")
            safe_time = current_time.replace(":", "-").replace(" ", "_")
            file_name = f"IsEPREP_Stock-In_{ttype.replace(' ', '_')}_{safe_time}.xlsx"

            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title=lang.t("stock_in.save_excel", "Save Excel"),
                initialfile=file_name,
                initialdir=default_dir
            )
            if not path:
                self.status_var.set(lang.t("stock_in.export_cancelled", "Export cancelled"))
                return

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = lang.t("stock_in.stock_in", "Stock In")

            project_name, project_code = fetch_project_details()
            doc_number = getattr(self, "current_document_number", None)

            # A1: Date + Doc number
            if doc_number:
                ws['A1'] = f"Date: {current_time}{' ' * 8}Document Number: {doc_number}"
            else:
                ws['A1'] = f"Date: {current_time}"
            ws['A1'].font = Font(name="Helvetica", size=10)
            ws['A1'].alignment = Alignment(horizontal="left")

            # A2: Title
            ws['A2'] = lang.t("stock_in.stock_in", "Stock In")
            ws['A2'].font = Font(name="Tahoma", size=14)
            ws['A2'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A2:J2')

            # A3: Project
            ws['A3'] = f"{project_name} - {project_code}"
            ws['A3'].font = Font(name="Tahoma", size=14)
            ws['A3'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A3:J3')

            # A4: IN Type
            ws['A4'] = f"{lang.t('stock_in.in_type', 'In Type')}: {ttype}"
            ws['A4'].font = Font(name="Tahoma")
            ws['A4'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A4:J4')

            ws.append([])  # Blank row

            headers = [
                lang.t("stock_in.code", "Code"),
                lang.t("stock_in.description", "Description"),
                lang.t("stock_in.scenario_name", "Scenario Name"),
                lang.t("stock_in.kit_code", "Kit"),
                lang.t("stock_in.module_code", "Module"),
                lang.t("stock_in.std_qty", "Std Qty"),
                lang.t("stock_in.qty_needed", "Qty Needed"),
                lang.t("stock_in.qty_in", "Qty In"),
                lang.t("stock_in.expiry_date", "Expiry Date"),
                lang.t("stock_in.batch_no", "Batch No")
            ]
            ws.append(headers)

            kit_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
            module_fill = PatternFill(start_color="F0FFF0", end_color="F0FFF0", fill_type="solid")

            rows_data = rows_to_export or [
                {k: v for k, v in zip(
                    ['code', 'description', 'scenario_name', 'kit_code', 'module_code', 'std_qty', 'qty_needed', 'qty_in', 'expiry_date', 'batch_no'],
                    self.tree.item(i)["values"]
                )}
                for i in self.tree.get_children()
                if self.tree.item(i)["values"][7]
            ]

            start_row = ws.max_row + 1
            for offset, r in enumerate(rows_data):
                row_idx = start_row + offset
                ws.append([
                    r['code'], r['description'], r['scenario_name'], r['kit_code'], r['module_code'],
                    r['std_qty'], r['qty_needed'], r['qty_in'], r['expiry_date'], r['batch_no']
                ])
                # Type highlight
                conn = connect_db()
                cur = conn.cursor()
                cur.execute("SELECT type FROM items_list WHERE code=?", (r['code'],))
                type_row = cur.fetchone()
                cur.close()
                if conn:
                    conn.close()
                item_type = type_row[0].upper() if type_row and type_row[0] else None
                if r['kit_code'] != '-----' and item_type == "KIT":
                    for col_cells in ws[f"A{row_idx}:J{row_idx}"]:
                        for cell in col_cells:
                            cell.fill = kit_fill
                elif r['module_code'] != '-----' and item_type == "MODULE":
                    for col_cells in ws[f"A{row_idx}:J{row_idx}"]:
                        for cell in col_cells:
                            cell.fill = module_fill

            ws.column_dimensions['A'].width = 100 / 7
            ws.column_dimensions['B'].width = 335 / 7
            ws.column_dimensions['C'].width = 120 / 7
            ws.column_dimensions['D'].width = 120 / 7
            ws.column_dimensions['E'].width = 120 / 7
            ws.column_dimensions['F'].width = 80 / 7
            ws.column_dimensions['G'].width = 90 / 7
            ws.column_dimensions['H'].width = 90 / 7
            ws.column_dimensions['I'].width = 110 / 7
            ws.column_dimensions['J'].width = 120 / 7

            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToHeight = 0
            ws.page_setup.fitToWidth = 1

            wb.save(path)
            wb.close()
            msg = lang.t("stock_in.export_success", "Export successful: {path}").format(path=path)
            custom_popup(self, lang.t("dialog_titles.success", "Success"), msg, "info")
            self.status_var.set(msg)
        except Exception as e:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         lang.t("stock_in.export_failed", "Export failed: {error}").format(error=str(e)), "error")
            self.status_var.set(lang.t("stock_in.export_error", "Export error: {error}").format(error=str(e)))

# ---------------------- Runner ---------------------- #
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Stock In")
    app = tk.Tk()
    app.role = "admin"
    StockIn(root, app, role="admin")
    root.mainloop()
