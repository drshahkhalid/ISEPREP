"""
losses.py  v1.1

Losses Report

Purpose:
  Track all losses (negative outbound adjustments) by loss Out_Type categories:
     Expired Items, Damaged Items, Cold Chain Break, Batch Recall, Theft, Other Losses
  Aggregation key: (Date, Code, Out_Type)
  Summarizes quantity (sum Qty_Out) and provides contextual metadata:
     scenarios, kit numbers, module numbers, expiry dates, document numbers, remarks.

Filters (Drop Downs / Inputs):
  - Scenario (scenarios.name)
  - Kit Number (from stock_data.kit_number)        (All / specific)
  - Module Number (from stock_data.module_number)  (All / specific)
  - Type (All, Kit, Module, Item)  (using detect_type)
  - Type of Losses (All + list of loss categories)
  - Item Search (code or translated description; only applied to Items)
  - Date Range: From / To (multi-format parsing; To capped at today)
  - Document Number substring
  - Simple / Detailed mode toggle

Columns:
  Simple Mode:
    date, code, description, quantity, type_of_loss
  Detailed Mode:
    date, scenarios, kits, modules, type, code, description,
    quantity, type_of_loss, expiry_dates, documents, remarks

  * scenarios / kits / modules / expiry_dates / documents / remarks
    are comma-separated unique sorted lists aggregated inside each (date, code, out_type) group.

Quantity Definition:
  Sum of Qty_Out for rows whose Out_Type is one of the defined loss categories
  (and matches the Type of Losses filter if not 'All').

Expiry Dates:
  Pulled primarily from stock_data.exp_date (left join via unique_id).
  If stock_data.exp_date is null, fallback to stock_transactions.Expiry_date.

Remarks:
  Aggregated distinct stock_transactions.Remarks values.

Document Number:
  Comma-separated documents for the grouped rows.

Keyboard:
  - Enter in search/document/date fields: Refresh
  - Esc in an Entry: Clear that field & refresh
  - Esc globally (outside Entry): Clear all filters

Export:
  - Excel (includes filter summary + current mode)
  - Rows shaded (KIT green / MODULE light blue) based on detect_type

Dependencies:
  - db.connect_db
  - manage_items.get_item_description / detect_type
  - language_manager.lang
  - popup_utils.custom_popup
  - tkcalendar (optional)

Future Enhancements (not yet implemented):
  - Graphical trend lines
  - Drill-down double-click to open underlying transactions
  - Pivot by scenario or third party (if required later)

"""

import tkinter as tk
from tkinter import ttk, filedialog
import sqlite3
import re
from datetime import date, datetime
from calendar import monthrange
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# Optional calendar
try:
    from tkcalendar import DateEntry
    TKCAL_AVAILABLE = True
except Exception:
    TKCAL_AVAILABLE = False

from db import connect_db
from manage_items import get_item_description, detect_type
from language_manager import lang
from popup_utils import custom_popup

# ============================================================
# IMPORT CENTRALIZED THEME (NEW)
# ============================================================
from theme_config import AppTheme, configure_tree_tags

# ============================================================
# REMOVED OLD GENERAL COLOR CONSTANTS - Now using AppTheme
# ============================================================
# OLD (REMOVED):
# BG_MAIN        = "#F0F4F8"
# BG_PANEL       = "#FFFFFF"
# COLOR_PRIMARY  = "#2C3E50"
# COLOR_BORDER   = "#D0D7DE"
# ROW_ALT_COLOR  = "#F7FAFC"
# ROW_NORM_COLOR = "#FFFFFF"
# BTN_EXPORT     = "#2980B9"
# BTN_REFRESH    = "#2563EB"
# BTN_CLEAR      = "#7F8C8D"
# BTN_TOGGLE     = "#8E44AD"

# ============================================================
# KEPT VISUALIZATION-SPECIFIC COLORS (Excel fill specific)
# ============================================================
KIT_FILL_COLOR    = "228B22"    # Excel fill color (no #)
MODULE_FILL_COLOR = "ADD8E6"    # Excel fill color (no #)

LOSS_TYPES = [
    "Expired Items",
    "Damaged Items",
    "Cold Chain Break",
    "Batch Recall",
    "Theft",
    "Other Losses"
]
LOSS_TYPES_SET = set(LOSS_TYPES)

# ---------------- Date Parsing ----------------
DATE_REGEXES = [
    ("%Y-%m-%d", re.compile(r"^\d{4}-\d{2}-\d{2}$")),
    ("%d/%m/%Y", re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")),
    ("%d-%m-%Y", re.compile(r"^\d{1,2}-\d{1,2}-\d{4}$")),
    ("%Y/%m/%d", re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$")),
    ("%d %b %Y", re.compile(r"^\d{1,2}\s+[A-Za-z]{3}\s+\d{4}$")),
    ("%d %B %Y", re.compile(r"^\d{1,2}\s+[A-Za-z]+\s+\d{4}$")),
]
MONTH_ONLY_REGEXES = [
    ("%Y-%m", re.compile(r"^\d{4}-\d{1,2}$")),
    ("%m/%Y", re.compile(r"^\d{1,2}/\d{4}$")),
    ("%b-%Y", re.compile(r"^[A-Za-z]{3}-\d{4}$")),
    ("%B-%Y", re.compile(r"^[A-Za-z]+-\d{4}$")),
]

def parse_user_date(text: str, role: str):
    if not text:
        return None
    raw = text.strip()
    for fmt, rx in DATE_REGEXES:
        if rx.match(raw):
            try:
                return datetime.strptime(raw, fmt).date()
            except Exception:
                pass
    for fmt, rx in MONTH_ONLY_REGEXES:
        if rx.match(raw):
            try:
                dt = datetime.strptime(raw, fmt)
                y, m = dt.year, dt.month
                return date(y, m, 1) if role == "from" else date(y, m, monthrange(y, m)[1])
            except Exception:
                pass
    if re.match(r"^\d{4}$", raw):
        y = int(raw)
        return date(y, 1, 1) if role == "from" else date(y, 12, 31)
    return None

# ---------------- Filter Value Providers ----------------
def fetch_scenarios():
    conn = connect_db()
    if conn is None: return []
    cur = conn.cursor()
    try:
        cur.execute("SELECT name FROM scenarios ORDER BY name")
        return [r[0] for r in cur.fetchall()]
    except sqlite3.Error:
        return []
    finally:
        cur.close(); conn.close()

def fetch_kit_numbers():
    """Kit numbers from stock_data.kit_number"""
    conn = connect_db()
    if conn is None: return []
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT DISTINCT kit_number FROM stock_data
            WHERE kit_number IS NOT NULL AND kit_number!='None' AND kit_number!=''
            ORDER BY kit_number
        """)
        return [r[0] for r in cur.fetchall()]
    except sqlite3.Error:
        return []
    finally:
        cur.close(); conn.close()

def fetch_module_numbers():
    """Module numbers from stock_data.module_number"""
    conn = connect_db()
    if conn is None: return []
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT DISTINCT module_number FROM stock_data
            WHERE module_number IS NOT NULL AND module_number!='None' AND module_number!=''
            ORDER BY module_number
        """)
        return [r[0] for r in cur.fetchall()]
    except sqlite3.Error:
        return []
    finally:
        cur.close(); conn.close()

# ---------------- Aggregation ----------------
def aggregate_losses(filters):
    scenario = filters.get("scenario")
    kit_number = filters.get("kit")
    module_number = filters.get("module")
    type_filter = filters.get("type")
    item_search = filters.get("item_search")
    doc_search = filters.get("doc_number")
    loss_filter = filters.get("loss_type")
    date_from = filters.get("date_from")
    date_to = filters.get("date_to")

    where = ["t.Out_Type IN ({})".format(",".join("?" for _ in LOSS_TYPES_SET))]
    params = list(LOSS_TYPES_SET)

    if scenario and scenario.lower() != "all":
        where.append("t.Scenario = ?")
        params.append(scenario)
    if doc_search:
        where.append("t.document_number LIKE ?")
        params.append(f"%{doc_search}%")
    if date_from:
        where.append("t.Date >= ?")
        params.append(date_from.strftime("%Y-%m-%d"))
    if date_to:
        where.append("t.Date <= ?")
        params.append(date_to.strftime("%Y-%m-%d"))
    if loss_filter and loss_filter.lower() != "all":
        where.append("t.Out_Type = ?")
        params.append(loss_filter)
    if kit_number and kit_number.lower() != "all":
        where.append("sd.kit_number = ?")
        params.append(kit_number)
    if module_number and module_number.lower() != "all":
        where.append("sd.module_number = ?")
        params.append(module_number)

    where_sql = " AND ".join(where)

    sql = f"""
        SELECT
            t.Date, t.code, t.Scenario, t.Qty_Out, t.Out_Type,
            t.document_number, t.Remarks, t.unique_id,
            COALESCE(sd.exp_date, t.Expiry_date) as expiry_date,
            sd.kit_number, sd.module_number
        FROM stock_transactions t
        LEFT JOIN stock_data sd ON t.unique_id = sd.unique_id
        WHERE {where_sql}
    """

    conn = connect_db()
    if conn is None:
        return []

    cur = conn.cursor()
    try:
        cur.execute(sql, params)
        fetched = cur.fetchall()
    except sqlite3.Error:
        cur.close(); conn.close()
        return []

    groups = {}
    scen_map = defaultdict(set)
    kit_map = defaultdict(set)
    mod_map = defaultdict(set)
    doc_map = defaultdict(set)
    expiry_map = defaultdict(set)
    remarks_map = defaultdict(set)

    for (dt_str, code, scen, qty_out, out_type, doc, remarks, unique_id,
         exp_date, kn, mn) in fetched:
        if not code:
            continue
        desc = get_item_description(code)
        dtype = detect_type(code, desc)

        if type_filter and type_filter.lower() != "all":
            if dtype.lower() != type_filter.lower():
                continue

        if item_search:
            if dtype.lower() != "item":
                continue
            if (item_search.lower() not in code.lower()
                and item_search.lower() not in desc.lower()):
                continue

        key = (dt_str, code, out_type)
        rec = groups.setdefault(key, {
            "date": dt_str,
            "code": code,
            "type": dtype,
            "description": desc,
            "out_type": out_type,
            "quantity": 0
        })

        if qty_out:
            try:
                rec["quantity"] += int(qty_out)
            except Exception:
                pass

        if scen: scen_map[key].add(scen)
        if kn and kn not in ("None",""): kit_map[key].add(kn)
        if mn and mn not in ("None",""): mod_map[key].add(mn)
        if doc: doc_map[key].add(doc)
        if exp_date and exp_date not in ("None",""): expiry_map[key].add(exp_date)
        if remarks: remarks_map[key].add(remarks)

    rows = []
    for key, rec in groups.items():
        rec["scenarios"] = ", ".join(sorted(scen_map[key])) if scen_map[key] else ""
        rec["kits"] = ", ".join(sorted(kit_map[key])) if kit_map[key] else ""
        rec["modules"] = ", ".join(sorted(mod_map[key])) if mod_map[key] else ""
        rec["documents"] = ", ".join(sorted(doc_map[key])) if doc_map[key] else ""
        rec["expiry_dates"] = ", ".join(sorted(expiry_map[key])) if expiry_map[key] else ""
        rec["remarks"] = ", ".join(sorted(remarks_map[key])) if remarks_map[key] else ""
        rows.append(rec)

    rows.sort(key=lambda r: (r["date"], r["type"], r["code"], r["out_type"]))
    cur.close(); conn.close()
    return rows

# ---------------- UI Class (UPDATED: All color references use AppTheme) ----------------
class Losses(tk.Frame):
    SIMPLE_COLS = ["date","code","description","quantity","out_type"]
    DETAIL_COLS = [
        "date","scenarios","kits","modules","type","code","description",
        "quantity","out_type","expiry_dates","documents","remarks"
    ]

    def __init__(self, parent, app, *args, **kwargs):
        super().__init__(parent, bg=AppTheme.BG_MAIN, *args, **kwargs)  # UPDATED: Use AppTheme
        self.app = app
        self.rows = []
        self.simple_mode = False

        all_lbl = self._all_label()
        self.scenario_var = tk.StringVar(value=all_lbl)
        self.kit_var = tk.StringVar(value=all_lbl)
        self.module_var = tk.StringVar(value=all_lbl)
        self.type_var = tk.StringVar(value=all_lbl)
        self.loss_type_var = tk.StringVar(value=all_lbl)
        self.item_search_var = tk.StringVar()
        self.doc_var = tk.StringVar()
        self.from_var = tk.StringVar()
        self.to_var = tk.StringVar()

        self.status_var = tk.StringVar(value=lang.t("losses.ready","Ready"))
        self.tree = None

        self._build_ui()
        self.populate_filters()
        self._set_default_dates()
        self.refresh()

    def _all_label(self):
        return lang.t("losses.all", "All")

    def _norm(self, val):
        all_lbl = self._all_label()
        return "All" if (val is None or val == "" or val == all_lbl) else val

    def _norm_type(self, val):
        if val in (None, "", self._all_label()):
            return "All"
        if val == lang.t("losses.type_kit","Kit"):
            return "Kit"
        if val == lang.t("losses.type_module","Module"):
            return "Module"
        if val == lang.t("losses.type_item","Item"):
            return "Item"
        return val

    def _loss_display_pairs(self):
        return [
            ("Expired Items",   lang.t("losses.loss_expired","Expired Items")),
            ("Damaged Items",   lang.t("losses.loss_damaged","Damaged Items")),
            ("Cold Chain Break",lang.t("losses.loss_cold_chain","Cold Chain Break")),
            ("Batch Recall",    lang.t("losses.loss_batch_recall","Batch Recall")),
            ("Theft",           lang.t("losses.loss_theft","Theft")),
            ("Other Losses",    lang.t("losses.loss_other","Other Losses")),
        ]

    def _norm_loss_type(self, display_val):
        if display_val in (None, "", self._all_label()):
            return "All"
        for canonical, disp in self._loss_display_pairs():
            if display_val == disp:
                return canonical
        return display_val

    # ---------- UI (UPDATED: All color references use AppTheme) ----------
    def _build_ui(self):
        header = tk.Frame(self, bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        header.pack(fill="x", padx=12, pady=(12,6))

        tk.Label(header,
                 text=lang.t("menu.reports.losses","Losses"),
                 font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HUGE, "bold"),  # UPDATED: Use AppTheme
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY).pack(side="left")  # UPDATED: Use AppTheme

        self.toggle_btn = tk.Button(header,
                                    text=lang.t("losses.detailed","Detailed"),
                                    bg=AppTheme.BTN_TOGGLE,  # UPDATED: Use AppTheme
                                    fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                                    relief="flat", padx=14, pady=6,
                                    command=self.toggle_mode)
        self.toggle_btn.pack(side="right", padx=(6,0))

        tk.Button(header, text=lang.t("generic.export","Export"),
                  bg=AppTheme.BTN_EXPORT,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14, pady=6, command=self.export_excel)\
            .pack(side="right", padx=(6,0))
        tk.Button(header, text=lang.t("generic.clear","Clear"),
                  bg=AppTheme.BTN_NEUTRAL,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14, pady=6, command=self.clear_filters)\
            .pack(side="right", padx=(6,0))
        tk.Button(header, text=lang.t("generic.refresh","Refresh"),
                  bg=AppTheme.BTN_REFRESH,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14, pady=6, command=self.refresh)\
            .pack(side="right", padx=(6,0))

        filters = tk.Frame(self, bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        filters.pack(fill="x", padx=12, pady=(0,8))

        all_lbl = self._all_label()
        type_all_lbl   = all_lbl
        type_kit_lbl   = lang.t("losses.type_kit","Kit")
        type_mod_lbl   = lang.t("losses.type_module","Module")
        type_item_lbl  = lang.t("losses.type_item","Item")

        # Row 1
        r1 = tk.Frame(filters, bg=AppTheme.BG_MAIN); r1.pack(fill="x", pady=2)  # UPDATED: Use AppTheme

        tk.Label(r1, text=lang.t("generic.scenario","Scenario"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=0,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.scenario_cb = ttk.Combobox(r1, textvariable=self.scenario_var,
                                        state="readonly", width=20)
        self.scenario_cb.grid(row=0,column=1,padx=(0,12))

        tk.Label(r1, text=lang.t("generic.kit_number","Kit Number"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=2,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.kit_cb = ttk.Combobox(r1, textvariable=self.kit_var,
                                   state="readonly", width=16)
        self.kit_cb.grid(row=0,column=3,padx=(0,12))

        tk.Label(r1, text=lang.t("generic.module_number","Module Number"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=4,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.module_cb = ttk.Combobox(r1, textvariable=self.module_var,
                                      state="readonly", width=16)
        self.module_cb.grid(row=0,column=5,padx=(0,12))

        tk.Label(r1, text=lang.t("generic.type","Type"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=6,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.type_cb = ttk.Combobox(r1, textvariable=self.type_var,
                                    state="readonly", width=12,
                                    values=[type_all_lbl, type_kit_lbl, type_mod_lbl, type_item_lbl])
        self.type_cb.grid(row=0,column=7,padx=(0,12))

        tk.Label(r1, text=lang.t("losses.type_of_loss","Type of Losses"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=8,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        loss_display = [disp for _, disp in self._loss_display_pairs()]
        self.loss_type_cb = ttk.Combobox(r1, textvariable=self.loss_type_var,
                                         state="readonly", width=18,
                                         values=[all_lbl] + loss_display)
        self.loss_type_cb.grid(row=0,column=9,padx=(0,4))

        # Row 2
        r2 = tk.Frame(filters, bg=AppTheme.BG_MAIN); r2.pack(fill="x", pady=2)  # UPDATED: Use AppTheme

        tk.Label(r2, text=lang.t("losses.item_search","Item Search"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=0,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.item_entry = tk.Entry(r2, textvariable=self.item_search_var, width=20)
        self.item_entry.grid(row=0,column=1,padx=(0,12))
        self.item_entry.bind("<Return>", lambda e: self.refresh())
        self.item_entry.bind("<Escape>", lambda e: self._clear_field(self.item_search_var))

        tk.Label(r2, text=lang.t("generic.document_number","Document Number"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=2,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.doc_entry = tk.Entry(r2, textvariable=self.doc_var, width=20)
        self.doc_entry.grid(row=0,column=3,padx=(0,12))
        self.doc_entry.bind("<Return>", lambda e: self.refresh())
        self.doc_entry.bind("<Escape>", lambda e: self._clear_field(self.doc_var))

        tk.Label(r2, text=lang.t("generic.from_date","From Date"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=4,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.from_entry = self._date_widget(r2, self.from_var)
        self.from_entry.grid(row=0,column=5,padx=(0,12))

        tk.Label(r2, text=lang.t("generic.to_date","To Date"), bg=AppTheme.BG_MAIN)\
            .grid(row=0,column=6,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.to_entry = self._date_widget(r2, self.to_var)
        self.to_entry.grid(row=0,column=7,padx=(0,12))

        # Status bar (UPDATED: Colors use AppTheme)
        tk.Label(self, textvariable=self.status_var, anchor="w",
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                 relief="sunken")\
            .pack(fill="x", padx=12, pady=(0,8))

        # Table (UPDATED: Colors use AppTheme)
        table_frame = tk.Frame(self, bg=AppTheme.COLOR_BORDER, bd=1, relief="solid")  # UPDATED: Use AppTheme
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0,12))

        self.tree = ttk.Treeview(table_frame, columns=(), show="headings", height=22)
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        hsb.pack(fill="x", padx=12, pady=(0,12))
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Style configuration (UPDATED: Removed theme_use, using AppTheme)
        style = ttk.Style()
        # REMOVED: style.theme_use("clam") - already applied globally
        style.configure("Treeview",
                        background=AppTheme.BG_PANEL,  # UPDATED: Use AppTheme
                        fieldbackground=AppTheme.BG_PANEL,  # UPDATED: Use AppTheme
                        foreground=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                        rowheight=24,
                        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL))  # UPDATED: Use AppTheme
        style.configure("Treeview.Heading",
                        background="#E5E8EB",
                        foreground=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                        font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HEADING, "bold"))  # UPDATED: Use AppTheme
        self.tree.tag_configure("alt", background=AppTheme.ROW_ALT)  # UPDATED: Use AppTheme
        self.tree.tag_configure("kitrow", background=AppTheme.KIT_COLOR, foreground=AppTheme.TEXT_WHITE)  # UPDATED: Use AppTheme
        self.tree.tag_configure("modrow", background=AppTheme.MODULE_COLOR)  # UPDATED: Use AppTheme

        # Global ESC (clear all) unless inside Entry
        self.bind_all("<Escape>", self._esc_global)

    def _esc_global(self, event):
        if isinstance(event.widget, tk.Entry):
            return
        self.clear_filters()

    def _date_widget(self, parent, var):
        if TKCAL_AVAILABLE:
            w = DateEntry(parent, textvariable=var, width=12,
                          date_pattern="yyyy-mm-dd",
                          showweeknumbers=False,
                          background=AppTheme.COLOR_ACCENT,  # UPDATED: Use AppTheme
                          foreground=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                          borderwidth=1)
            w.bind("<Return>", lambda e: self.refresh())
            w.bind("<Escape>", lambda e: self._clear_field(var))
            return w
        else:
            e = tk.Entry(parent, textvariable=var, width=12)
            e.bind("<Return>", lambda ev: self.refresh())
            e.bind("<Escape>", lambda ev: self._clear_field(var))
            return e

    def _clear_field(self, tk_var):
        tk_var.set("")
        self.refresh()

    def _set_default_dates(self):
        today = date.today()
        try:
            past = date(today.year - 1, today.month, today.day)
        except Exception:
            past = today
        self.from_var.set(past.strftime("%Y-%m-%d"))
        self.to_var.set(today.strftime("%Y-%m-%d"))

    # ---------- Filter population ----------
    def populate_filters(self):
        scenarios = fetch_scenarios()
        kits = fetch_kit_numbers()
        modules = fetch_module_numbers()
        all_lbl = self._all_label()

        self.scenario_cb['values'] = [all_lbl] + scenarios
        self.kit_cb['values'] = [all_lbl] + kits
        self.module_cb['values'] = [all_lbl] + modules

        for var, cb in [
            (self.scenario_var, self.scenario_cb),
            (self.kit_var, self.kit_cb),
            (self.module_var, self.module_cb)
        ]:
            if var.get() not in cb['values']:
                var.set(all_lbl)

    # ---------- Mode toggle ----------
    def toggle_mode(self):
        self.simple_mode = not self.simple_mode
        self.toggle_btn.config(
            text=lang.t("losses.simple","Simple") if self.simple_mode
            else lang.t("losses.detailed","Detailed")
        )
        self._populate_tree()

    # ---------- Refresh ----------
    def refresh(self):
        filters = {
            "scenario": self._norm(self.scenario_var.get()),
            "kit": self._norm(self.kit_var.get()),
            "module": self._norm(self.module_var.get()),
            "type": self._norm_type(self.type_var.get()),
            "loss_type": self._norm_loss_type(self.loss_type_var.get()),
            "item_search": self.item_search_var.get().strip(),
            "doc_number": self.doc_var.get().strip(),
            "date_from": parse_user_date(self.from_var.get().strip(), "from") if self.from_var.get().strip() else None,
            "date_to": parse_user_date(self.to_var.get().strip(), "to") if self.to_var.get().strip() else None
        }
        if filters["date_to"] and filters["date_to"] > date.today():
            filters["date_to"] = date.today()
        self.rows = aggregate_losses(filters)
        self._populate_tree()
        self.status_var.set(lang.t("losses.loaded","Loaded {n} rows").format(n=len(self.rows)))

    def clear_filters(self):
        all_lbl = self._all_label()
        self.scenario_var.set(all_lbl)
        self.kit_var.set(all_lbl)
        self.module_var.set(all_lbl)
        self.type_var.set(all_lbl)
        self.loss_type_var.set(all_lbl)
        self.item_search_var.set("")
        self.doc_var.set("")
        self._set_default_dates()
        self.refresh()

    # ---------- Columns ----------
    def _current_columns(self):
        return self.SIMPLE_COLS if self.simple_mode else self.DETAIL_COLS

    def _populate_tree(self):
        cols = self._current_columns()
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols

        headings = {
            "date": lang.t("generic.date","Date"),
            "scenarios": lang.t("generic.scenario","Scenario(s)"),
            "kits": lang.t("generic.kit_number","Kit Number(s)"),
            "modules": lang.t("generic.module_number","Module Number(s)"),
            "type": lang.t("generic.type","Type"),
            "code": lang.t("generic.code","Code"),
            "description": lang.t("generic.description","Description"),
            "quantity": lang.t("losses.quantity","Quantity"),
            "out_type": lang.t("losses.type_of_loss","Type of Loss"),
            "expiry_dates": lang.t("losses.expiry_dates","Expiry Date(s)"),
            "documents": lang.t("generic.document_number","Document Number(s)"),
            "remarks": lang.t("losses.remarks","Remarks")
        }

        for c in cols:
            width = 120
            if c == "description":
                width = 340
            elif c == "code":
                width = 160
            elif c in ("scenarios","kits","modules","documents","remarks","expiry_dates"):
                width = 220
            elif c in ("quantity","out_type"):
                width = 130
            self.tree.heading(c, text=headings.get(c, c))
            self.tree.column(c, width=width, anchor="w", stretch=True)

        for idx, r in enumerate(self.rows):
            if self.simple_mode:
                values = [
                    r["date"], r["code"], r["description"],
                    r["quantity"], r["out_type"]
                ]
            else:
                values = [
                    r["date"], r["scenarios"], r["kits"], r["modules"], r["type"],
                    r["code"], r["description"],
                    r["quantity"], r["out_type"],
                    r["expiry_dates"], r["documents"], r["remarks"]
                ]
            tag = "alt" if idx % 2 else ""
            if r["type"].upper() == "KIT":
                tag = "kitrow"
            elif r["type"].upper() == "MODULE":
                tag = "modrow"
            self.tree.insert("", "end", values=values, tags=(tag,))

    # ---------- Export ----------
    def export_excel(self):
        if not self.rows:
            custom_popup(self, lang.t("generic.info","Info"),
                         lang.t("losses.no_data_export","Nothing to export."), "warning")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            title=lang.t("losses.export_title","Save Losses Report"),
            initialfile=f"Losses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Losses"

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws.append([lang.t("generic.generated","Generated"), now_str])
            ws.append([lang.t("generic.filters_used","Filters Used")])
            ws.append(["Scenario", self.scenario_var.get(),
                       "Kit", self.kit_var.get(),
                       "Module", self.module_var.get()])
            ws.append(["Type", self.type_var.get(),
                       "Loss Type", self.loss_type_var.get(),
                       "Document", self.doc_var.get()])
            ws.append(["From", self.from_var.get(),
                       "To", self.to_var.get(),
                       "Mode", "Simple" if self.simple_mode else "Detailed"])
            ws.append([])

            cols = self._current_columns()
            headers = [c.replace("_"," ").title() for c in cols]
            ws.append(headers)

            kit_fill = PatternFill(start_color=KIT_FILL_COLOR, end_color=KIT_FILL_COLOR, fill_type="solid")
            module_fill = PatternFill(start_color=MODULE_FILL_COLOR, end_color=MODULE_FILL_COLOR, fill_type="solid")

            for r in self.rows:
                if self.simple_mode:
                    line = [
                        r["date"], r["code"], r["description"],
                        r["quantity"], r["out_type"]
                    ]
                else:
                    line = [
                        r["date"], r["scenarios"], r["kits"], r["modules"], r["type"],
                        r["code"], r["description"],
                        r["quantity"], r["out_type"],
                        r["expiry_dates"], r["documents"], r["remarks"]
                    ]
                ws.append(line)
                dtype = r["type"].upper()
                if dtype == "KIT":
                    for c in ws[ws.max_row]: c.fill = kit_fill
                elif dtype == "MODULE":
                    for c in ws[ws.max_row]: c.fill = module_fill

            for col in ws.columns:
                max_len = 0
                letter = get_column_letter(col[0].column)
                for cell in col:
                    v = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(v))
                ws.column_dimensions[letter].width = min(max_len + 2, 60)

            wb.save(path)
            custom_popup(self, lang.t("generic.success","Success"),
                         lang.t("losses.export_success","Export completed: {f}").format(f=path),
                         "info")
        except Exception as e:
            custom_popup(self, lang.t("generic.error","Error"),
                         lang.t("losses.export_fail","Export failed: {err}").format(err=str(e)),
                         "error")

# ---------------- Module Export ----------------
__all__ = ["Losses"]

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Losses Report v1.1")
    Losses(root, None).pack(fill="both", expand=True)
    root.geometry("1580x820")
    root.mainloop()