import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import sqlite3
import openpyxl
from db import connect_db
from language_manager import lang
from item_families import ItemFamilyManager
from popup_utils import custom_popup, custom_askyesno, custom_dialog

# ============================================================
# THEME / STYLE
# ============================================================
BG_MAIN        = "#F0F4F8"
BG_PANEL       = "#FFFFFF"
COLOR_PRIMARY  = "#2C3E50"
COLOR_ACCENT   = "#2563EB"
COLOR_BORDER   = "#D0D7DE"
ROW_ALT_COLOR  = "#F7FAFC"
ROW_NORM_COLOR = "#FFFFFF"
BTN_ADD        = "#27AE60"
BTN_EDIT       = "#F39C12"
BTN_DELETE     = "#C0392B"
BTN_IMPORT     = "#2980B9"
BTN_EXPORT     = "#2980B9"
BTN_CLEAR      = "#8E44AD"
BTN_MISC       = "#7F8C8D"
BTN_DISABLED   = "#94A3B8"


# ============================================================
# EDIT PERMISSIONS
# Restrict modification for ~ and $
# Canonical names OR symbols are both checked. 
# ============================================================
RESTRICTED_MODIFY = {"manager", "supervisor", "~", "$"}

# ============================================================
# DB bootstrap
# ============================================================
def ensure_table():
    conn = connect_db()
    if conn is None:
        return
    cursor = conn.cursor()
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS items_list (
                item_id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT UNIQUE NOT NULL,
                pack TEXT,
                price_per_pack_euros REAL,
                unit_price_euros REAL,
                weight_per_pack_kg REAL,
                volume_per_pack_dm3 REAL,
                shelf_life_months INTEGER,
                remarks TEXT,
                account_code TEXT,
                type TEXT,
                designation TEXT,
                designation_en TEXT,
                designation_fr TEXT,
                designation_sp TEXT,
                unique_id_1 TEXT
            )
        """)
        conn.commit()
    finally:
        cursor.close()
        conn.close()

# ============================================================
# TYPE / DESCRIPTION HELPERS
# ============================================================
def detect_type(code, designation):
    if not code: 
        return "Item"
    code = str(code).strip()
    designation = str(designation or "").lower()
    if code. upper().startswith("K"):
        if designation.startswith("kit") or "modules" in designation:
            return "Kit"
        if "module" in designation and not designation.startswith("kit"):
            return "Module"
    return "Item"

def generate_unique_id(code):
    return code

def get_item_description(code):
    conn = connect_db()
    if conn is None: 
        return "No Description"
    cursor = conn.cursor()
    try:
        cursor.execute(
            "SELECT designation, designation_en, designation_fr, designation_sp FROM items_list WHERE code=?",
            (code,))
        row = cursor.fetchone()
        if not row:
            return "No Description"
        row_dict = {
            "designation": row[0],
            "designation_en":  row[1],
            "designation_fr": row[2],
            "designation_sp": row[3]
        }
        lang_code = lang.lang_code. lower()
        mapping = {"en": "designation_en", "fr": "designation_fr", "es": "designation_sp", "sp": "designation_sp"}
        active_col = mapping.get(lang_code, "designation_en")
        if row_dict. get(active_col):
            return row_dict[active_col]
        if row_dict.get("designation_en"):
            return row_dict["designation_en"]
        if row_dict.get("designation_fr"):
            return row_dict["designation_fr"]
        if row_dict.get("designation_sp"):
            return row_dict["designation_sp"]
        return row_dict. get("designation") or "No Description"
    finally:
        cursor.close()
        conn.close()

def get_family_remarks(code):
    if not code or len(code) < 4:
        return None
    try:
        family_manager = ItemFamilyManager()
        return family_manager.get_remarks_by_item_code(code)
    except Exception: 
        return None

# ============================================================
# MAIN CLASS
# ============================================================
class ManageItems(tk.Frame):
    """
    Items management (read/write) with permission control: 
      - Users with role symbol '~' (manager) or '$' (supervisor), or
        canonical names 'manager' / 'supervisor' are READ-ONLY. 
      - All other roles retain existing functionality.
    """
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG_MAIN)
        self.app = app
        self.role = getattr(app, "role", "supervisor")
        self.tree = None
        self.search_var = tk.StringVar()
        self.total_label = None
        ensure_table()
        self._configure_styles()
        self._build_ui()
        self.load_data()

    # --------------- Permission helpers ---------------
    def _can_modify(self):
        return (self.role or "").strip().lower() not in RESTRICTED_MODIFY

    def _deny(self):
        custom_popup(self,
                     lang.t("dialog_titles.restricted", fallback="Restricted"),
                     self. t("no_modify_permission", fallback="You do not have permission to modify items."),
                     "warning")

    # ---------------- Translation helper ----------------
    def t(self, key, fallback=None, **kwargs):
        """Translate a key under 'items' section"""
        return lang.t(f"items.{key}", fallback=fallback, **kwargs)

    # ---------------- Language-specific column mapping ----------------
    def get_active_lang_column(self):
        """Return the designation column name for the active language."""
        lang_code = lang. lang_code.lower()
        mapping = {
            "en": "designation_en",
            "fr": "designation_fr",
            "es": "designation_sp",
            "sp": "designation_sp"
        }
        return mapping. get(lang_code, "designation_en")

    # ---------------- Style configuration ----------------
    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style. configure(
            "Items. Treeview",
            background=BG_PANEL,
            fieldbackground=BG_PANEL,
            foreground=COLOR_PRIMARY,
            rowheight=26,
            font=("Helvetica", 10),
            bordercolor=COLOR_BORDER,
            relief="flat"
        )
        style.map("Items.Treeview",
                  background=[("selected", COLOR_ACCENT)],
                  foreground=[("selected", "#FFFFFF")])
        style.configure(
            "Items.Treeview.Heading",
            background="#E5E8EB",
            foreground=COLOR_PRIMARY,
            font=("Helvetica", 11, "bold"),
            relief="flat",
            bordercolor=COLOR_BORDER
        )
        style.configure("Items.TEntry", font=("Helvetica", 10))
        style.configure("Items.TCombobox", font=("Helvetica", 10))

    # ---------------- Safe cleaners ----------------
    def clean_str(self, value):
        if value is None:
            return None
        v = str(value).strip()
        if v.lower() in ("nan", "none", ""):
            return None
        return v

    def safe_float(self, value):
        try:
            if value in (None, "", "nan"):
                return None
            return float(value)
        except Exception:
            return None

    def safe_int(self, value):
        try:
            if value in (None, "", "nan"):
                return None
            return int(value)
        except Exception: 
            return None

    def determine_type(self, code, designation, db_type=None):
        if db_type and str(db_type).strip():
            return db_type
        return detect_type(code, designation)

    def get_active_designation(self, row):
        """Get designation in order of priority:  active language > en > fr > sp."""
        lang_code = lang. lang_code.lower()
        mapping = {"en": "designation_en", "fr": "designation_fr", "es": "designation_sp", "sp": "designation_sp"}
        active_col = mapping.get(lang_code, "designation_en")
        
        # First try active language
        if row. get(active_col):
            return row[active_col]
        # Fallback priority:  en > fr > sp
        if row.get("designation_en"):
            return row["designation_en"]
        if row.get("designation_fr"):
            return row["designation_fr"]
        if row.get("designation_sp"):
            return row["designation_sp"]
        return row.get("designation") or ""

    # ---------------- UI Construction ----------------
    def _build_ui(self):
        tk.Label(
            self,
            text=self.t("title", fallback="Manage Items"),
            font=("Helvetica", 20, "bold"),
            bg=BG_MAIN,
            fg=COLOR_PRIMARY,
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(12, 6))

        # Search panel
        search_frame = tk. Frame(self, bg=BG_MAIN)
        search_frame.pack(fill="x", padx=12, pady=(0, 8))

        tk.Label(search_frame,
                 text=self.t("search_label", fallback="Search: "),
                 bg=BG_MAIN, fg=COLOR_PRIMARY,
                 font=("Helvetica", 10),
                 anchor="w").pack(side="left", padx=(0, 6))

        search_entry = tk.Entry(search_frame,
                                textvariable=self.search_var,
                                width=40,
                                font=("Helvetica", 10),
                                relief="solid",
                                bd=1)
        search_entry.pack(side="left", padx=(0, 6))
        search_entry.bind("<Return>", lambda e: self.search_items())

        tk.Button(search_frame,
                  text=self.t("search_button", fallback="Search"),
                  command=self. search_items,
                  bg=COLOR_ACCENT, fg="#FFFFFF",
                  activebackground="#1D4ED8",
                  relief="flat", padx=12, pady=5,
                  font=("Helvetica", 10, "bold")).pack(side="left", padx=4)

        tk.Button(search_frame,
                  text=self.t("clear_button", fallback="Clear"),
                  command=self. load_data,
                  bg=BTN_MISC, fg="#FFFFFF",
                  activebackground="#64748B",
                  relief="flat", padx=12, pady=5,
                  font=("Helvetica", 10, "bold")).pack(side="left", padx=4)

        # Tree frame
        tree_outer = tk.Frame(self, bg=COLOR_BORDER, bd=1, relief="solid")
        tree_outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        columns = (
            "code", "designation", "type", "pack", "price_per_pack_euros",
            "unit_price_euros", "weight_per_pack_kg", "volume_per_pack_dm3",
            "shelf_life_months", "remarks", "account_code"
        )
        self.tree = ttk.Treeview(tree_outer,
                                 columns=columns,
                                 show="headings",
                                 height=20,
                                 style="Items.Treeview")
        col_config = {
            "code": {"width":  140, "anchor": "w"},
            "designation": {"width": 380, "anchor": "w"},
            "type": {"width": 80, "anchor": "w"},
            "pack": {"width": 80, "anchor": "e"},
            "price_per_pack_euros": {"width": 110, "anchor": "e"},
            "unit_price_euros":  {"width": 110, "anchor": "e"},
            "weight_per_pack_kg": {"width": 110, "anchor":  "e"},
            "volume_per_pack_dm3": {"width": 110, "anchor":  "e"},
            "shelf_life_months": {"width": 110, "anchor": "e"},
            "remarks": {"width":  160, "anchor": "w"},
            "account_code": {"width": 120, "anchor": "e"}
        }
        header_names = {
            "code": self.t("col_code", fallback="Code"),
            "designation": self.t("col_designation", fallback="Designation"),
            "type":  self.t("col_type", fallback="Type"),
            "pack": self.t("col_pack", fallback="Pack"),
            "price_per_pack_euros": self.t("col_price_per_pack", fallback="Price/Pack"),
            "unit_price_euros":  self.t("col_unit_price", fallback="Unit Price"),
            "weight_per_pack_kg": self.t("col_weight_pack", fallback="Weight/Pack (kg)"),
            "volume_per_pack_dm3": self.t("col_volume_pack", fallback="Volume/Pack (dm³)"),
            "shelf_life_months": self.t("col_shelf_life", fallback="Shelf Life (m)"),
            "remarks": self.t("col_remarks", fallback="Remarks"),
            "account_code": self.t("col_account_code", fallback="Account Code")
        }
        for col in columns:
            self.tree. heading(col, text=header_names. get(col, col. title()))
            self.tree.column(col, **col_config[col])
        self.tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(tree_outer, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        # Buttons (with permission control)
        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill="x", padx=12, pady=(0, 6))
        can_modify = self._can_modify()

        def mk_btn(label, cmd, color, allow_readonly=False):
            is_allowed = can_modify or allow_readonly
            return tk.Button(btn_frame,
                             text=label,
                             command=cmd if is_allowed else self._deny,
                             bg=color if is_allowed else BTN_DISABLED,
                             fg="#FFFFFF",
                             activebackground=color if is_allowed else BTN_DISABLED,
                             relief="flat",
                             padx=14, pady=6,
                             font=("Helvetica", 10, "bold"))

        btn_import = mk_btn(self.t("import_button", fallback="Import"), self.import_excel, BTN_IMPORT)
        btn_export = mk_btn(self.t("export_button", fallback="Export"), self.export_excel, BTN_EXPORT, allow_readonly=True)
        btn_add = mk_btn(self.t("add_button", fallback="Add"), self.add_item, BTN_ADD)
        btn_edit = mk_btn(self.t("edit_button", fallback="Edit"), self.edit_item, BTN_EDIT)
        btn_delete = mk_btn(self. t("delete_button", fallback="Delete"), self.delete_item, BTN_DELETE)
        btn_clear_all = mk_btn(self.t("clear_all_button", fallback="Clear All"), self.clear_all, BTN_CLEAR)

        for b in [btn_import, btn_export, btn_add, btn_edit, btn_delete, btn_clear_all]:
            b. pack(side="left", padx=4)

        if not can_modify:
            custom_popup(self,
                         lang.t("dialog_titles. restricted", fallback="Restricted"),
                         self.t("read_only_mode", fallback="Read-only mode:  you cannot modify items."),
                         "warning")

        # Status bar
        status_frame = tk.Frame(self, bg=BG_MAIN)
        status_frame.pack(fill="x", padx=12, pady=(0, 8))
        self.total_label = tk.Label(status_frame,
                                    text=self.t("total_items", count=0, fallback="Total Items:  0"),
                                    bg=BG_MAIN, fg=COLOR_PRIMARY,
                                    anchor="w",
                                    font=("Helvetica", 10))
        self.total_label.pack(side="left")

    # ---------------- Data Loading ----------------
    def load_data(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self.t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT * FROM items_list")
            cols = [d[0] for d in cursor. description]
            rows = cursor.fetchall()
            for idx, row in enumerate(rows):
                row_dict = dict(zip(cols, row))
                designation = self.get_active_designation(row_dict)
                item_type = self.determine_type(row_dict['code'], designation, row_dict. get('type'))
                display_tuple = (
                    row_dict["code"],
                    designation or "",
                    item_type or "",
                    row_dict["pack"] or "",
                    self._blank_if_none(row_dict["price_per_pack_euros"]),
                    self._blank_if_none(row_dict["unit_price_euros"]),
                    self._blank_if_none(row_dict["weight_per_pack_kg"]),
                    self._blank_if_none(row_dict["volume_per_pack_dm3"]),
                    self._blank_if_none(row_dict["shelf_life_months"]),
                    row_dict["remarks"] or "",
                    row_dict["account_code"] or ""
                )
                tag = "alt" if idx % 2 else "norm"
                self.tree.insert("", "end", values=display_tuple, tags=(tag,))
            self._configure_row_tags()
            cursor.execute("SELECT COUNT(*) FROM items_list")
            total = cursor.fetchone()[0]
            self.total_label.config(text=self.t("total_items", count=total, fallback=f"Total Items: {total}"))
        finally:
            cursor.close()
            conn.close()
        self.search_var.set("")

    def _configure_row_tags(self):
        self.tree.tag_configure("norm", background=ROW_NORM_COLOR)
        self.tree.tag_configure("alt", background=ROW_ALT_COLOR)

    def _blank_if_none(self, value):
        return "" if value is None else value

    # ---------------- Search ----------------
    def search_items(self):
        text = self.search_var.get().strip()
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self. t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cursor = conn.cursor()
        try:
            if not text:
                cursor.execute("SELECT * FROM items_list")
                count_query = "SELECT COUNT(*) FROM items_list"
                count_params = ()
            else:
                like = f"%{text}%"
                query = """
                    SELECT * FROM items_list
                    WHERE UPPER(code) LIKE UPPER(?)
                       OR UPPER(unique_id_1) LIKE UPPER(?)
                       OR UPPER(designation_en) LIKE UPPER(?)
                       OR UPPER(designation_fr) LIKE UPPER(?)
                       OR UPPER(designation_sp) LIKE UPPER(?)
                """
                cursor.execute(query, (like, like, like, like, like))
                count_query = """
                    SELECT COUNT(*) FROM items_list
                    WHERE UPPER(code) LIKE UPPER(?)
                       OR UPPER(unique_id_1) LIKE UPPER(?)
                       OR UPPER(designation_en) LIKE UPPER(?)
                       OR UPPER(designation_fr) LIKE UPPER(?)
                       OR UPPER(designation_sp) LIKE UPPER(?)
                """
                count_params = (like, like, like, like, like)

            cols = [d[0] for d in cursor.description]
            rows = cursor.fetchall()
            for idx, row in enumerate(rows):
                row_dict = dict(zip(cols, row))
                designation = self.get_active_designation(row_dict)
                item_type = self.determine_type(row_dict['code'], designation, row_dict.get('type'))
                display_tuple = (
                    row_dict["code"],
                    designation or "",
                    item_type or "",
                    row_dict["pack"] or "",
                    self._blank_if_none(row_dict["price_per_pack_euros"]),
                    self._blank_if_none(row_dict["unit_price_euros"]),
                    self._blank_if_none(row_dict["weight_per_pack_kg"]),
                    self._blank_if_none(row_dict["volume_per_pack_dm3"]),
                    self._blank_if_none(row_dict["shelf_life_months"]),
                    row_dict["remarks"] or "",
                    row_dict["account_code"] or ""
                )
                tag = "alt" if idx % 2 else "norm"
                self.tree.insert("", "end", values=display_tuple, tags=(tag,))
            self._configure_row_tags()
            cursor.execute(count_query, count_params)
            total = cursor.fetchone()[0]
            self.total_label.config(text=self.t("total_items", count=total, fallback=f"Total Items: {total}"))
        finally:
            cursor.close()
            conn.close()
        self.search_var.set("")

    # ---------------- Export ----------------
    def export_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title=self.t("export_dialog_title", fallback="Save Items List")
        )
        if not file_path:
            return
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self. t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT * FROM items_list")
            cols = [d[0] for d in cursor.description]
            rows = cursor.fetchall()
            data = [dict(zip(cols, r)) for r in rows]

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Items List"
            
            # Export with active language designation
            headers = [
                self.t("excel_header_code", fallback="Code"),
                self.t("excel_header_designation", fallback="Designation"),
                self.t("excel_header_type", fallback="Type"),
                self.t("excel_header_pack", fallback="Pack"),
                self. t("excel_header_price_per_pack_euros", fallback="Price/Pack[Euros]"),
                self.t("excel_header_unit_price_euros", fallback="Unit price[Euros]"),
                self.t("excel_header_weight_per_pack_kg", fallback="Weight/pack[kg]"),
                self.t("excel_header_volume_per_pack_dm3", fallback="Volume/pack[dm3]"),
                self.t("excel_header_shelf_life_months", fallback="Shelf life (months)"),
                self.t("excel_header_remarks", fallback="Remarks"),
                self. t("excel_header_account_code", fallback="Account code")
            ]
            ws.append(headers)

            for d in data:
                designation = self.get_active_designation(d)
                item_type = self.determine_type(d["code"], designation, d.get("type"))
                ws.append([
                    d["code"],
                    designation or "",
                    item_type,
                    d["pack"] or "",
                    d["price_per_pack_euros"] if d["price_per_pack_euros"] is not None else "",
                    d["unit_price_euros"] if d["unit_price_euros"] is not None else "",
                    d["weight_per_pack_kg"] if d["weight_per_pack_kg"] is not None else "",
                    d["volume_per_pack_dm3"] if d["volume_per_pack_dm3"] is not None else "",
                    d["shelf_life_months"] if d["shelf_life_months"] is not None else "",
                    d["remarks"] or "",
                    d["account_code"] or ""
                ])

            for col in ws.columns:
                max_len = 0
                letter = col[0].column_letter
                for cell in col: 
                    try:
                        l = len(str(cell.value)) if cell.value is not None else 0
                        if l > max_len: 
                            max_len = l
                    except Exception:
                        pass
                ws.column_dimensions[letter]. width = min(max_len + 2, 60)

            wb.save(file_path)
            custom_popup(self, lang.t("dialog_titles.success", fallback="Success"),
                         self.t("export_success", fallback="Exported to {path}").format(path=file_path),
                         "info")
        except Exception as e:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self. t("export_failed", fallback="Export failed: {err}").format(err=str(e)),
                         "error")
        finally:
            cursor.close()
            conn.close()

    # ---------------- Clear All ----------------
    def clear_all(self):
        if not self._can_modify():
            self._deny()
            return
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self. t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT COUNT(*) FROM compositions")
            count = cursor.fetchone()[0]
            if count > 0:
                custom_popup(
                    self,
                    lang.t("dialog_titles.warning", fallback="Warning"),
                    self.t("cannot_clear_items", fallback="Cannot clear Items List while Standard List has data.  Clear it first."),
                    "warning"
                )
                return
            ans = custom_askyesno(self,
                                  lang.t("dialog_titles.confirm", fallback="Confirm"),
                                  self. t("confirm_clear_all",
                                         fallback="Clear all items?  This cannot be undone."))
            if ans != "yes":
                return
            cursor.execute("DELETE FROM items_list")
            conn.commit()
            self.load_data()
            custom_popup(self, lang. t("dialog_titles.success", fallback="Success"),
                         self.t("cleared_items", fallback="All items cleared.", count=0),
                         "info")
        finally:
            cursor.close()
            conn.close()

    # ---------------- Import ----------------
    def import_excel(self):
        if not self._can_modify():
            self._deny()
            return
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title=self.t("import_dialog_title", fallback="Import Items from Excel")
        )
        if not file_path:
            return
        ans = custom_askyesno(
            self,
            lang.t("dialog_titles.confirm", fallback="Confirm"),
            self.t("confirm_import", fallback="Import and merge items from this file? ")
        )
        if ans != "yes":
            return
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self.t("import_failed", fallback="Failed to read file: {err}").format(err=str(e)),
                         "error")
            return
        
        # Extended column mapping to include all designation variants
        column_mapping = {
            "Code": "code",
            "Pack": "pack",
            "Price/pack[Euros]": "price_per_pack_euros",
            "Unit price[Euros]":  "unit_price_euros",
            "Weight/pack[kg]": "weight_per_pack_kg",
            "Volume/pack[dm3]": "volume_per_pack_dm3",
            "Shelf life (months)": "shelf_life_months",
            "Remarks": "remarks",
            "Account code": "account_code",
            "Designation": "designation",
            "Designation_EN": "designation_en",
            "Designation_FR": "designation_fr",
            "Designation_SP": "designation_sp",
            # Alternative column names
            "Designation EN": "designation_en",
            "Designation FR": "designation_fr",
            "Designation SP": "designation_sp",
            "Designation ES": "designation_sp",
        }
        
        # Add missing columns with None
        for col in column_mapping: 
            if col not in df. columns:
                df[col] = None
        
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self. t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cursor = conn.cursor()
        success = 0
        try:
            for _, r in df.iterrows():
                code = self. clean_str(r["Code"])
                if not code:
                    continue
                
                # Collect all designation fields from import
                designation_data = {
                    "designation":  self.clean_str(r. get("Designation")),
                    "designation_en": self.clean_str(r.get("Designation_EN")) or self.clean_str(r.get("Designation EN")),
                    "designation_fr": self.clean_str(r.get("Designation_FR")) or self.clean_str(r.get("Designation FR")),
                    "designation_sp": self.clean_str(r.get("Designation_SP")) or self.clean_str(r.get("Designation SP")) or self.clean_str(r.get("Designation ES")),
                }
                
                # If no language-specific designation provided, use generic "Designation" for all
                if not designation_data["designation_en"] and designation_data["designation"]:
                    designation_data["designation_en"] = designation_data["designation"]
                
                # Priority fallback for main designation field
                main_designation = (designation_data["designation_en"] or 
                                  designation_data["designation_fr"] or 
                                  designation_data["designation_sp"] or 
                                  designation_data["designation"])
                
                data = {
                    "pack": self.clean_str(r["Pack"]),
                    "price_per_pack_euros":  self.safe_float(r["Price/pack[Euros]"]),
                    "unit_price_euros":  self.safe_float(r["Unit price[Euros]"]),
                    "weight_per_pack_kg": self.safe_float(r["Weight/pack[kg]"]),
                    "volume_per_pack_dm3": self.safe_float(r["Volume/pack[dm3]"]),
                    "shelf_life_months": self.safe_int(r["Shelf life (months)"]),
                    "remarks": self.clean_str(r["Remarks"]),
                    "account_code": self.clean_str(r["Account code"]),
                    "designation": main_designation,
                    "designation_en": designation_data["designation_en"],
                    "designation_fr": designation_data["designation_fr"],
                    "designation_sp": designation_data["designation_sp"],
                }
                
                data["type"] = self.determine_type(code, main_designation)
                unique_id_1 = generate_unique_id(code)

                family_remarks = get_family_remarks(code)
                if family_remarks:
                    if data["remarks"]:
                        data["remarks"] = f"{data['remarks']}, {family_remarks}"
                    else:
                        data["remarks"] = family_remarks

                cursor.execute("SELECT code FROM items_list WHERE code=?", (code,))
                exists = cursor.fetchone()
                if exists:
                    cursor. execute("""
                        UPDATE items_list SET
                            pack=?, price_per_pack_euros=?, unit_price_euros=?,
                            weight_per_pack_kg=?, volume_per_pack_dm3=?, shelf_life_months=?,
                            remarks=?, account_code=?, designation=?, designation_en=?,
                            designation_fr=?, designation_sp=?, type=?, unique_id_1=? 
                        WHERE code=?
                    """, (
                        data["pack"], data["price_per_pack_euros"], data["unit_price_euros"],
                        data["weight_per_pack_kg"], data["volume_per_pack_dm3"], data["shelf_life_months"],
                        data["remarks"], data["account_code"], data["designation"], data["designation_en"],
                        data["designation_fr"], data["designation_sp"], data["type"], unique_id_1, code
                    ))
                else:
                    if len(code) < 8 or not main_designation or len(main_designation) < 15:
                        continue
                    cursor.execute("""
                        INSERT INTO items_list (
                            code, pack, price_per_pack_euros, unit_price_euros,
                            weight_per_pack_kg, volume_per_pack_dm3, shelf_life_months,
                            remarks, account_code, designation, designation_en,
                            designation_fr, designation_sp, type, unique_id_1
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        code, data["pack"], data["price_per_pack_euros"], data["unit_price_euros"],
                        data["weight_per_pack_kg"], data["volume_per_pack_dm3"], data["shelf_life_months"],
                        data["remarks"], data["account_code"], data["designation"], data["designation_en"],
                        data["designation_fr"], data["designation_sp"], data["type"], unique_id_1
                    ))
                success += 1
            conn.commit()
            custom_popup(self,
                         lang.t("dialog_titles.success", fallback="Success"),
                         self.t("import_complete",
                                fallback="Imported {success} / {total} rows",
                                success=success, total=len(df)),
                         "info")
            self.load_data()
        except Exception as e:
            conn.rollback()
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self. t("import_failed", fallback="Import failed: {err}").format(err=str(e)),
                         "error")
        finally:
            cursor.close()
            conn.close()

    # ---------------- CRUD:  Add/Edit/Delete ----------------
    def add_item(self):
        if not self._can_modify():
            self._deny()
            return
        self._item_form(self.t("add_title", fallback="Add Item"))

    def edit_item(self):
        if not self._can_modify():
            self._deny()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self, lang.t("dialog_titles.warning", fallback="Warning"),
                         self.t("select_edit", fallback="Select an item to edit"),
                         "warning")
            return
        values = self.tree.item(sel[0])["values"]
        code = values[0]
        
        # Fetch full record from DB to get all designation fields
        conn = connect_db()
        if conn is None: 
            custom_popup(self, lang.t("dialog_titles. error", fallback="Error"),
                         self.t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cursor = conn. cursor()
        try:
            cursor.execute("SELECT * FROM items_list WHERE code=?", (code,))
            cols = [d[0] for d in cursor.description]
            row = cursor.fetchone()
            if row:
                full_record = dict(zip(cols, row))
                self._item_form(self.t("edit_title", fallback="Edit Item"), values, full_record)
            else:
                custom_popup(self, lang.t("dialog_titles. error", fallback="Error"),
                             self.t("item_not_found", fallback="Item not found"),
                             "error")
        finally:
            cursor.close()
            conn.close()

    def delete_item(self):
        if not self._can_modify():
            self._deny()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self, lang.t("dialog_titles.warning", fallback="Warning"),
                         self.t("select_delete", fallback="Select an item to delete"),
                         "warning")
            return
        code = self.tree.item(sel[0])["values"][0]
        ans = custom_askyesno(
            self,
            lang.t("dialog_titles.confirm", fallback="Confirm"),
            self.t("confirm_delete", fallback="Delete item {code}?").format(code=code)
        )
        if ans != "yes":
            return
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", fallback="Error"),
                         self.t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM items_list WHERE code=? ", (code,))
            conn.commit()
            self.load_data()
            custom_popup(self, lang.t("dialog_titles.success", fallback="Success"),
                         self.t("delete_success", fallback="Item deleted"),
                         "info")
        finally:
            cursor.close()
            conn.close()

    # ---------------- Item Form (Add/Edit) ----------------
    def _item_form(self, title, values=None, full_record=None):
        # Extra safety
        if not self._can_modify():
            self._deny()
            return
        form = tk. Toplevel(self)
        form.title(title)
        form.configure(bg=BG_MAIN)
        form.geometry("520x840")
        form.transient(self)
        form.grab_set()

        tk.Label(form,
                 text=title,
                 font=("Helvetica", 16, "bold"),
                 fg=COLOR_PRIMARY,
                 bg=BG_MAIN,
                 anchor="w").pack(fill="x", padx=16, pady=(14, 8))

        fields = [
            ("code", True),
            ("designation", True),
            ("type", False),
            ("pack", False),
            ("price_per_pack_euros", False),
            ("unit_price_euros", False),
            ("weight_per_pack_kg", False),
            ("volume_per_pack_dm3", False),
            ("shelf_life_months", False),
            ("remarks", False),
            ("account_code", False)
        ]
        entries = {}

        def add_field(fname, required):
            # Translate field labels
            label_key = f"field_{fname}"
            label_text = self.t(label_key, fallback=fname.replace('_', ' ').title())
            tk.Label(form,
                     text=f"{label_text}{' *' if required else ''}:",
                     font=("Helvetica", 10),
                     fg=COLOR_PRIMARY,
                     bg=BG_MAIN,
                     anchor="w").pack(fill="x", padx=18, pady=(6, 0))
            ent = tk.Entry(form,
                           font=("Helvetica", 11),
                           relief="solid",
                           bd=1)
            ent.pack(fill="x", padx=18, pady=(0, 6))
            entries[fname] = ent

        for fname, req in fields:
            add_field(fname, req)

        if values and full_record:
            # When editing, show the active language's designation
            active_designation = self.get_active_designation(full_record)
            entries["code"].insert(0, values[0] if values[0] is not None else "")
            entries["designation"].insert(0, active_designation)
            entries["type"].insert(0, values[2] if values[2] is not None else "")
            entries["pack"].insert(0, values[3] if values[3] is not None else "")
            entries["price_per_pack_euros"].insert(0, values[4] if values[4] is not None and values[4] != "" else "")
            entries["unit_price_euros"].insert(0, values[5] if values[5] is not None and values[5] != "" else "")
            entries["weight_per_pack_kg"].insert(0, values[6] if values[6] is not None and values[6] != "" else "")
            entries["volume_per_pack_dm3"]. insert(0, values[7] if values[7] is not None and values[7] != "" else "")
            entries["shelf_life_months"].insert(0, values[8] if values[8] is not None and values[8] != "" else "")
            entries["remarks"].insert(0, values[9] if values[9] is not None else "")
            entries["account_code"].insert(0, values[10] if values[10] is not None else "")
            entries["code"].config(state="disabled")

        def save():
            code = entries["code"].get().strip()
            designation = entries["designation"].get().strip()
            if not code or not designation:
                custom_popup(form, lang.t("dialog_titles.error", fallback="Error"),
                             self.t("required_fields", fallback="Code and Designation are required"),
                             "error")
                return
            if not values:  # Adding new item
                if len(code) < 8:
                    custom_popup(form, lang.t("dialog_titles.error", fallback="Error"),
                                 self.t("code_length", fallback="Code must be at least 8 characters"),
                                 "error")
                    return
                if len(designation) < 15:
                    custom_popup(form, lang.t("dialog_titles.error", fallback="Error"),
                                 self.t("designation_length", fallback="Designation must be at least 15 characters"),
                                 "error")
                    return

            provided_type = entries["type"].get().strip()
            final_type = self.determine_type(code, designation, provided_type)
            unique_id_1 = generate_unique_id(code)

            payload = {}
            for fname, _ in fields:
                raw = entries[fname].get().strip()
                payload[fname] = raw if raw else None

            payload["type"] = final_type

            # Get active language column
            active_lang_col = self.get_active_lang_column()
            
            # When editing, fetch existing designations to preserve other languages
            if full_record:
                payload["designation_en"] = full_record.get("designation_en")
                payload["designation_fr"] = full_record.get("designation_fr")
                payload["designation_sp"] = full_record.get("designation_sp")
            else:
                # New item - initialize all to None
                payload["designation_en"] = None
                payload["designation_fr"] = None
                payload["designation_sp"] = None
            
            # Update the active language designation
            if active_lang_col == "designation_en":
                payload["designation_en"] = designation
            elif active_lang_col == "designation_fr":
                payload["designation_fr"] = designation
            elif active_lang_col == "designation_sp":
                payload["designation_sp"] = designation
            
            # Set main designation field (priority:  en > fr > sp)
            payload["designation"] = (payload["designation_en"] or 
                                     payload["designation_fr"] or 
                                     payload["designation_sp"] or 
                                     designation)

            fam_rem = get_family_remarks(code)
            if fam_rem:
                if payload["remarks"]:
                    payload["remarks"] = f"{payload['remarks']}, {fam_rem}"
                else:
                    payload["remarks"] = fam_rem

            conn = connect_db()
            if conn is None:
                custom_popup(form, lang.t("dialog_titles.error", fallback="Error"),
                             self. t("db_error", fallback="Database connection failed"),
                             "error")
                return
            cursor = conn.cursor()
            try:
                if values:  # Update
                    cursor.execute("""
                        UPDATE items_list
                           SET designation=?,
                               designation_en=?,
                               designation_fr=?,
                               designation_sp=?,
                               pack=?,
                               price_per_pack_euros=?,
                               unit_price_euros=?,
                               weight_per_pack_kg=?,
                               volume_per_pack_dm3=?,
                               shelf_life_months=?,
                               remarks=?,
                               account_code=?,
                               type=?,
                               unique_id_1=?
                         WHERE code=?
                    """, (
                        payload["designation"],
                        payload["designation_en"],
                        payload["designation_fr"],
                        payload["designation_sp"],
                        payload["pack"],
                        self.safe_float(payload["price_per_pack_euros"]),
                        self.safe_float(payload["unit_price_euros"]),
                        self.safe_float(payload["weight_per_pack_kg"]),
                        self.safe_float(payload["volume_per_pack_dm3"]),
                        self. safe_int(payload["shelf_life_months"]),
                        payload["remarks"],
                        payload["account_code"],
                        payload["type"],
                        unique_id_1,
                        code
                    ))
                else:   # Insert
                    cursor.execute("""
                        INSERT INTO items_list (
                            code, designation, designation_en, designation_fr, designation_sp,
                            pack, price_per_pack_euros, unit_price_euros,
                            weight_per_pack_kg, volume_per_pack_dm3,
                            shelf_life_months, remarks, account_code,
                            type, unique_id_1
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (
                        code,
                        payload["designation"],
                        payload["designation_en"],
                        payload["designation_fr"],
                        payload["designation_sp"],
                        payload["pack"],
                        self.safe_float(payload["price_per_pack_euros"]),
                        self.safe_float(payload["unit_price_euros"]),
                        self.safe_float(payload["weight_per_pack_kg"]),
                        self. safe_float(payload["volume_per_pack_dm3"]),
                        self.safe_int(payload["shelf_life_months"]),
                        payload["remarks"],
                        payload["account_code"],
                        payload["type"],
                        unique_id_1
                    ))
                conn.commit()
                form.destroy()
                self.load_data()
                custom_popup(
                    self,
                    lang.t("dialog_titles.success", fallback="Success"),
                    self. t("save_success", fallback="Item saved successfully. "),
                    "info"
                )
            except sqlite3.IntegrityError:
                custom_popup(form, lang.t("dialog_titles.error", fallback="Error"),
                             self.t("duplicate_code", fallback="Code already exists"),
                             "error")
            except Exception as e:
                custom_popup(form, lang.t("dialog_titles.error", fallback="Error"),
                             self.t("save_failed", fallback="Save failed: {err}").format(err=str(e)),
                             "error")
            finally: 
                cursor.close()
                conn.close()

        tk.Button(form,
                  text=self.t("save_button", fallback="Save"),
                  command=save,
                  bg=BTN_ADD, fg="#FFFFFF",
                  relief="flat",
                  font=("Helvetica", 11, "bold"),
                  padx=14, pady=8,
                  activebackground="#1E874B").pack(pady=20, padx=18, fill="x")

    # Public refresh
    def refresh(self):
        self.load_data()


if __name__ == "__main__": 
    root = tk.Tk()
    class DummyApp:  pass
    dummy = DummyApp()
    # Try roles:  "admin", "manager", "~", "$"
    dummy.role = "admin"  # Full access for testing
    root.title("Manage Items")
    ManageItems(root, dummy)
    root.geometry("1200x780")
    root.mainloop()