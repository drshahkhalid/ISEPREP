"""
stock_summary.py
Stock Summary Dashboard (Scenario -> Kit Number -> Module Number)

HOTFIX:
  - Fixed TypeError during sorting: mixed float/string comparisons for kit_number/module_number.
  - Replaced previous numeric-aware tuple approach with a stable string normalizer (norm_sort)
    that produces a purely string-based sortable key (zero‑padded for numeric values).
  - No other logic changed.

Existing Features:
  - Scenario value normalization (handles scenario names or scenario_id stored as text in stock_data).
  - Grouping: Scenario / Kit Number / Module Number -> Code.
  - Filters: Scenario, Management Mode, Kit Number, Module Number, Item Code, Type, Expiry Period.
  - Metrics and Excel export.
  - Old SQLite 3.0 compatible (no CTE/window functions/IF EXISTS/RETURNING).

If you already integrated a prior version, you may replace only the load_data() function
and the helper norm_sort below. For completeness, the full file is supplied.
"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog
import sqlite3
import datetime
from calendar import monthrange
import re
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

from db import connect_db
from language_manager import lang
try:
    from popup_utils import custom_popup
except ImportError:
    from tkinter import messagebox
    def custom_popup(parent, title, msg, kind="info"):
        if kind == "error":
            messagebox.showerror(title, msg, parent=parent)
        elif kind == "warning":
            messagebox.showwarning(title, msg, parent=parent)
        else:
            messagebox.showinfo(title, msg, parent=parent)

# ----------------------------- DB Helpers -----------------------------
def _fetchall(sql, params=()):
    conn = connect_db()
    if conn is None:
        return []
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute(sql, params)
        return cur.fetchall()
    except sqlite3.Error:
        return []
    finally:
        cur.close(); conn.close()

def _fetchone(sql, params=()):
    rows = _fetchall(sql, params)
    return rows[0] if rows else None

# ------------------------- Expiry Period Logic ------------------------
def compute_recommended_months():
    cols = {r["name"].lower(): True for r in _fetchall("PRAGMA table_info(project_details)")}
    for need in ("lead_time_months","cover_period_months","buffer_months"):
        if need not in cols:
            return 0
    row = _fetchone("""
        SELECT
          COALESCE(lead_time_months,0) lt,
          COALESCE(cover_period_months,0) cp,
          COALESCE(buffer_months,0) bf
        FROM project_details
        ORDER BY rowid ASC
        LIMIT 1
    """)
    if not row:
        return 0
    return int(row["lt"]) + int(row["cp"]) + int(row["bf"])

def cutoff_date_from_months(months: int | None):
    if not months or months <= 0:
        return None
    today = datetime.date.today()
    year = today.year
    month = today.month + months
    while month > 12:
        month -= 12
        year += 1
    last_day = monthrange(year, month)[1]
    return datetime.date(year, month, last_day).isoformat()

# --------------------------- Scenario Mapping -------------------------
def load_scenario_maps():
    rows = _fetchall("SELECT scenario_id, name FROM scenarios")
    id_to_name = {}
    name_set = set()
    for r in rows:
        sid = r["scenario_id"]
        nm = r["name"]
        if nm:
            name_set.add(nm)
        id_to_name[str(sid)] = nm
    return id_to_name, name_set

def normalize_scenario(raw_val, id_to_name, name_set):
    if raw_val is None:
        return ""
    if raw_val in name_set:
        return raw_val
    txt = str(raw_val)
    if txt in id_to_name:
        return id_to_name[txt] or txt
    return txt

# --------------------------- Supporting Loads -------------------------
def load_treecodes():
    mapping = {}
    rows = _fetchall("""
        SELECT code, treecode FROM kit_items
        WHERE treecode IS NOT NULL AND treecode <> ''
        UNION ALL
        SELECT item AS code, treecode FROM kit_items
        WHERE item IS NOT NULL AND item <> '' AND treecode IS NOT NULL AND treecode <> ''
    """)
    for r in rows:
        c = r["code"]; t = r["treecode"]
        if not c or not t:
            continue
        if c not in mapping or t < mapping[c]:
            mapping[c] = t
    return mapping

def load_std_quantities(codes):
    if not codes:
        return {}
    std_map = {}
    conn = connect_db()
    if conn is None:
        return std_map
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        chunk = 200
        clist = list(codes)
        for i in range(0, len(clist), chunk):
            part = clist[i:i+chunk]
            ph = ",".join(["?"] * len(part))
            cur.execute(f"""
                SELECT code, std_qty_collective
                FROM std_list_combined
                WHERE code IN ({ph})
            """, part)
            for r in cur.fetchall():
                std_map[r["code"]] = r["std_qty_collective"] if r["std_qty_collective"] is not None else 0
    finally:
        cur.close(); conn.close()
    return std_map

def load_item_metadata(codes):
    if not codes:
        return {}
    meta = {}
    conn = connect_db()
    if conn is None:
        return meta
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        chunk = 200
        clist = list(codes)
        for i in range(0, len(clist), chunk):
            part = clist[i:i+chunk]
            ph = ",".join(["?"] * len(part))
            cur.execute(f"""
                SELECT code, designation, type, remarks
                FROM items_list
                WHERE code IN ({ph})
            """, part)
            for r in cur.fetchall():
                meta[r["code"]] = {
                    "description": r["designation"] or "",
                    "type": r["type"] or "",
                    "remarks": r["remarks"] or ""
                }
    finally:
        cur.close(); conn.close()
    return meta

# ---------------------------- Aggregation -----------------------------
def build_filters_clause(filters):
    where = ["1=1"]
    params = []
    scen = filters["scenario"]
    if scen:
        where.append("(scenario = ? OR scenario = (SELECT CAST(scenario_id AS TEXT) FROM scenarios WHERE name=? LIMIT 1))")
        params.extend([scen, scen])
    mm = filters["management_mode"].lower()
    if mm == "on-shelf":
        where.append("LOWER(management_mode) IN ('on_shelf','on-shelf','onshelf')")
    elif mm == "in-box":
        where.append("LOWER(management_mode) IN ('in_box','in-box','inbox')")
    if filters["kit_number"]:
        where.append("kit_number = ?")
        params.append(filters["kit_number"])
    if filters["module_number"]:
        where.append("module_number = ?")
        params.append(filters["module_number"])
    if filters["item_code"]:
        where.append("item = ?")
        params.append(filters["item_code"])
    return " AND ".join(where), params

def aggregate_stock(filters, cutoff_iso, id_to_name, name_set):
    where_clause, params = build_filters_clause(filters)
    base_sql = f"""
        SELECT
           scenario AS raw_scenario,
           kit_number,
           module_number,
           management_mode,
           COALESCE(item, module, kit) AS code,
           SUM(final_qty) AS current_stock,
           MIN(exp_date) AS earliest_expiry
        FROM stock_data
        WHERE {where_clause}
          AND final_qty IS NOT NULL
        GROUP BY raw_scenario, kit_number, module_number, management_mode, COALESCE(item,module,kit)
    """
    rows = _fetchall(base_sql, tuple(params))
    merged = {}
    for r in rows:
        norm_scen = normalize_scenario(r["raw_scenario"], id_to_name, name_set)
        code = r["code"]
        if not code:
            continue
        key = (norm_scen, r["kit_number"], r["module_number"], code)
        entry = merged.setdefault(key, {
            "current_stock": 0,
            "earliest_expiry": None,
            "expiring_qty": 0,
            "management_modes": set()
        })
        entry["current_stock"] += r["current_stock"] if r["current_stock"] is not None else 0
        ee = r["earliest_expiry"]
        if ee and ((entry["earliest_expiry"] is None) or ee < entry["earliest_expiry"]):
            entry["earliest_expiry"] = ee
        mm = r["management_mode"] or ""
        if mm:
            entry["management_modes"].add(mm)

    if cutoff_iso:
        exp_sql = f"""
            SELECT
              scenario AS raw_scenario,
              kit_number,
              module_number,
              management_mode,
              COALESCE(item,module,kit) AS code,
              SUM(final_qty) AS expiring_sum
            FROM stock_data
            WHERE {where_clause}
              AND exp_date IS NOT NULL
              AND exp_date <= ?
            GROUP BY raw_scenario, kit_number, module_number, management_mode, COALESCE(item,module,kit)
        """
        exp_params = params + [cutoff_iso]
        exp_rows = _fetchall(exp_sql, tuple(exp_params))
        for er in exp_rows:
            norm_scen = normalize_scenario(er["raw_scenario"], id_to_name, name_set)
            key = (norm_scen, er["kit_number"], er["module_number"], er["code"])
            if key in merged:
                merged[key]["expiring_qty"] += er["expiring_sum"] if er["expiring_sum"] is not None else 0
    return merged

# ------------------------- Distinct Values ---------------------------
def distinct_kit_numbers(scenario=None):
    if scenario:
        rows = _fetchall("""
            SELECT DISTINCT kit_number FROM stock_data
            WHERE kit_number IS NOT NULL
              AND (scenario = ? OR scenario = (SELECT CAST(scenario_id AS TEXT) FROM scenarios WHERE name=? LIMIT 1))
            ORDER BY kit_number
        """, (scenario, scenario))
    else:
        rows = _fetchall("SELECT DISTINCT kit_number FROM stock_data WHERE kit_number IS NOT NULL ORDER BY kit_number")
    return [r["kit_number"] for r in rows if r["kit_number"]]

def distinct_module_numbers(scenario=None, kit_number=None):
    where = ["module_number IS NOT NULL"]
    params = []
    if scenario:
        where.append("(scenario = ? OR scenario = (SELECT CAST(scenario_id AS TEXT) FROM scenarios WHERE name=? LIMIT 1))")
        params.extend([scenario, scenario])
    if kit_number:
        where.append("kit_number = ?"); params.append(kit_number)
    sql = "SELECT DISTINCT module_number FROM stock_data WHERE " + " AND ".join(where) + " ORDER BY module_number"
    rows = _fetchall(sql, tuple(params))
    return [r["module_number"] for r in rows if r["module_number"]]

# ---------------------------- UI Window ------------------------------
class StockSummaryWindow(tk.Toplevel):
    def __init__(self, parent=None, role="user"):
        super().__init__(parent)
        self.role = role or "user"
        self.title(lang.t("stock_summary.title","Stock Summary"))
        self.geometry("1520x880")
        self.configure(bg="#F0F4F8")
        self._all_rows = []
        self._build_ui()
        self.populate_scenarios()
        self.populate_kit_module_lists()
        self.refresh_button.focus_set()
        self.status_var.set(lang.t("reports.ready","Ready (role={role})", role=self.role))

    def _build_ui(self):
        tk.Label(self,
                 text=lang.t("stock_summary.header","Stock Summary Dashboard"),
                 font=("Helvetica",20,"bold"), bg="#F0F4F8").pack(pady=10)

        f = tk.Frame(self, bg="#F0F4F8")
        f.pack(fill="x", padx=10, pady=4)

        tk.Label(f, text=lang.t("reports.scenario","Scenario:"), bg="#F0F4F8").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.scenario_var = tk.StringVar()
        self.scenario_cb = ttk.Combobox(f, textvariable=self.scenario_var, width=28, state="readonly")
        self.scenario_cb.grid(row=0, column=1, padx=4, pady=2, sticky="w")
        self.scenario_cb.bind("<<ComboboxSelected>>", lambda e: self.populate_kit_module_lists())

        tk.Label(f, text=lang.t("stock_summary.management_mode","Management Mode:"), bg="#F0F4F8").grid(row=0, column=2, sticky="w", padx=4, pady=2)
        self.management_var = tk.StringVar(value=lang.t("stock_summary.management_all","All"))
        self.management_cb = ttk.Combobox(
            f,
            textvariable=self.management_var,
            values=[
                lang.t("stock_summary.management_all","All"),
                lang.t("stock_summary.management_on_shelf","On-Shelf"),
                lang.t("stock_summary.management_in_box","In-Box")
            ],
            width=14,
            state="readonly"
        )
        self.management_cb.grid(row=0, column=3, padx=4, pady=2, sticky="w")
        self.management_cb.current(0)

        tk.Label(f, text=lang.t("reports.type","Type:"), bg="#F0F4F8").grid(row=0, column=4, sticky="w", padx=4, pady=2)
        self.type_var = tk.StringVar(value=lang.t("reports.type_all","All"))
        self.type_cb = ttk.Combobox(
            f,
            textvariable=self.type_var,
            values=[
                lang.t("reports.type_all","All"),
                "KIT",
                "MODULE",
                "ITEM"
            ],
            width=10,
            state="readonly"
        )
        self.type_cb.grid(row=0, column=5, padx=4, pady=2, sticky="w")
        self.type_cb.current(0)

        tk.Label(f, text=lang.t("reports.expiry_period","Expiry Period (months):"), bg="#F0F4F8").grid(row=0, column=6, sticky="w", padx=4, pady=2)
        self.period_override_var = tk.StringVar()
        self.period_entry = tk.Entry(f, textvariable=self.period_override_var, width=7)
        self.period_entry.grid(row=0, column=7, padx=4, pady=2, sticky="w")

        tk.Label(f, text=lang.t("stock_summary.kit_number","Kit Number:"), bg="#F0F4F8").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.kit_number_var = tk.StringVar()
        self.kit_number_cb = ttk.Combobox(f, textvariable=self.kit_number_var, width=20, state="readonly")
        self.kit_number_cb.grid(row=1, column=1, padx=4, pady=2, sticky="w")
        self.kit_number_cb.bind("<<ComboboxSelected>>", lambda e: self.populate_module_numbers())

        tk.Label(f, text=lang.t("stock_summary.module_number","Module Number:"), bg="#F0F4F8").grid(row=1, column=2, sticky="w", padx=4, pady=2)
        self.module_number_var = tk.StringVar()
        self.module_number_cb = ttk.Combobox(f, textvariable=self.module_number_var, width=20, state="readonly")
        self.module_number_cb.grid(row=1, column=3, padx=4, pady=2, sticky="w")

        tk.Label(f, text=lang.t("reports.item","Item:"), bg="#F0F4F8").grid(row=1, column=4, sticky="w", padx=4, pady=2)
        self.item_var = tk.StringVar()
        self.item_entry = tk.Entry(f, textvariable=self.item_var, width=18)
        self.item_entry.grid(row=1, column=5, padx=4, pady=2, sticky="w")

        self.recommended_label = tk.Label(self, text="", font=("Helvetica",8),
                                          fg="#555555", bg="#F0F4F8")
        self.recommended_label.pack(anchor="w", padx=14, pady=(0,2))

        sframe = tk.Frame(self, bg="#F0F4F8")
        sframe.pack(fill="x", padx=10, pady=(2,4))
        tk.Label(sframe, text=lang.t("reports.search","Search:"), bg="#F0F4F8").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)
        self.search_entry = tk.Entry(sframe, textvariable=self.search_var, width=40)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<Return>", lambda e: self.load_data())

        bframe = tk.Frame(self, bg="#F0F4F8")
        bframe.pack(fill="x", padx=10, pady=4)
        self.refresh_button = tk.Button(bframe, text=lang.t("reports.load_refresh","Load / Refresh"),
                                        bg="#27AE60", fg="white", command=self.load_data)
        self.refresh_button.pack(side="left", padx=4)
        tk.Button(bframe, text=lang.t("reports.clear_filters","Clear Filters"),
                  bg="#7F8C8D", fg="white", command=self.clear_filters).pack(side="left", padx=4)
        tk.Button(bframe, text=lang.t("reports.export","Export"),
                  bg="#2980B9", fg="white", command=self.export_excel).pack(side="left", padx=4)

        self.status_var = tk.StringVar(value=lang.t("reports.ready","Ready (role={role})", role=self.role))
        tk.Label(self, textvariable=self.status_var, anchor="w",
                 bg="#E0E4E8").pack(fill="x", padx=10, pady=(0,4))

        metrics_frame = tk.Frame(self, bg="#F0F4F8")
        metrics_frame.pack(fill="x", padx=10, pady=(0,4))
        self.metric_labels = {}
        for key in ["total_std","total_stock","total_shortage","total_overstock","coverage_pct","total_expiring"]:
            lbl = tk.Label(metrics_frame, text=f"{key}: 0", bg="#F0F4F8", font=("Helvetica",10,"bold"))
            lbl.pack(side="left", padx=8)
            self.metric_labels[key] = lbl

        self.columns = [
            "scenario","kit_number","module_number","management_modes","treecode","code",
            "description","type","standard_qty","current_stock","coverage_pct",
            "qty_expiring","earliest_expiry","over_stock","missing_qty","remarks"
        ]
        headers = {
            "scenario": lang.t("reports.scenario","Scenario"),
            "kit_number": lang.t("stock_summary.kit_number","Kit Number"),
            "module_number": lang.t("stock_summary.module_number","Module Number"),
            "management_modes": lang.t("stock_summary.management_mode","Management Mode(s)"),
            "treecode": lang.t("stock_summary.treecode","Treecode"),
            "code": lang.t("reports.code","Code"),
            "description": lang.t("reports.description","Description"),
            "type": lang.t("reports.type_header","Type"),
            "standard_qty": lang.t("reports.standard_qty","Standard Quantity"),
            "current_stock": lang.t("reports.current_stock","Current Stock"),
            "coverage_pct": lang.t("stock_summary.coverage_pct","Coverage %"),
            "qty_expiring": lang.t("stock_summary.qty_expiring","Qty Expiring (Period)"),
            "earliest_expiry": lang.t("stock_summary.earliest_expiry","Earliest Expiry"),
            "over_stock": lang.t("reports.over_stock","Over Stock"),
            "missing_qty": lang.t("reports.missing_qty","Missing Quantity"),
            "remarks": lang.t("reports.remarks","Remarks")
        }
        widths = {
            "scenario":140,"kit_number":110,"module_number":110,"management_modes":130,"treecode":90,"code":140,
            "description":280,"type":70,"standard_qty":120,"current_stock":110,"coverage_pct":100,"qty_expiring":150,
            "earliest_expiry":120,"over_stock":90,"missing_qty":120,"remarks":220
        }
        self.tree = ttk.Treeview(self, columns=self.columns, show="headings", height=27)
        for c in self.columns:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=widths.get(c,100), anchor="w")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill="both", expand=True, padx=10, pady=(0,6))
        vsb.place(relx=1.0, rely=0.365, relheight=0.585, anchor="ne")
        hsb.pack(fill="x", padx=10, pady=(0,6))

        self.tree.tag_configure("missing", background="#FFCCCC")
        self.tree.tag_configure("overstock", background="#CCFFCC")
        self.tree.tag_configure("group", background="#D9E1EC", font=("Helvetica",10,"bold"))

    # ---------------- Scenario & Distinct lists ----------------
    def populate_scenarios(self):
        rows = _fetchall("SELECT name FROM scenarios ORDER BY name")
        vals = [r["name"] for r in rows]
        self.scenario_cb["values"] = [""] + vals
        if vals:
            self.scenario_cb.current(0)
        self._update_recommended_label()

    def populate_kit_module_lists(self):
        scen = self.scenario_var.get() or None
        kits = distinct_kit_numbers(scen)
        self.kit_number_cb["values"] = [""] + kits
        self.kit_number_var.set("")
        self.populate_module_numbers()

    def populate_module_numbers(self):
        scen = self.scenario_var.get() or None
        kitn = self.kit_number_var.get() or None
        modules = distinct_module_numbers(scen, kitn)
        self.module_number_cb["values"] = [""] + modules
        self.module_number_var.set("")

    # ---------------- Filters / Helpers ----------------
    def gather_filters(self):
        # Normalize management mode
        msel = self.management_var.get()
        if msel in ("", lang.t("stock_summary.management_all","All")):
            management_mode = "All"
        elif msel == lang.t("stock_summary.management_on_shelf","On-Shelf"):
            management_mode = "On-Shelf"
        elif msel == lang.t("stock_summary.management_in_box","In-Box"):
            management_mode = "In-Box"
        else:
            management_mode = msel

        # Normalize type filter
        tsel = self.type_var.get()
        if tsel in ("", lang.t("reports.type_all","All")):
            type_filter = "All"
        else:
            type_filter = tsel

        return {
            "scenario": self.scenario_var.get() or None,
            "management_mode": management_mode,
            "kit_number": self.kit_number_var.get() or None,
            "module_number": self.module_number_var.get() or None,
            "item_code": self.item_var.get().strip() or None,
            "type_filter": type_filter
        }

    def clear_filters(self):
        self.scenario_var.set("")
        self.management_var.set("All")
        self.kit_number_var.set("")
        self.module_number_var.set("")
        self.item_var.set("")
        self.type_var.set("All")
        self.period_override_var.set("")
        self.search_var.set("")
        self.tree.delete(*self.tree.get_children())
        self._all_rows.clear()
        self._update_recommended_label()
        self._update_metrics({})
        self.status_var.set(lang.t("reports.filters_cleared","Filters cleared. Ready (role={role}).", role=self.role))

    def _update_recommended_label(self):
        rec = compute_recommended_months()
        self.recommended_label.config(
            text=lang.t("stock_summary.recommended",
                        "Recommended expiry period: {months} month(s) (lead+cover+buffer)",
                        months=rec)
        )

    # ---------------- Data Loading ----------------
    def load_data(self):
        self.tree.delete(*self.tree.get_children())
        self._all_rows.clear()
        filters = self.gather_filters()

        recommended = compute_recommended_months()
        raw = self.period_override_var.get().strip()
        if raw == "":
            months = recommended
        else:
            try:
                v = int(raw); months = v if v >= 0 else recommended
            except:
                months = recommended
        cutoff_iso = cutoff_date_from_months(months)
        exp_header = lang.t("stock_summary.qty_expiring_header",
                             "Quantity Expiring in next {months} months",
                             months=months)
        self.tree.heading("qty_expiring", text=exp_header)

        self.status_var.set(lang.t("reports.computing","Computing aggregates..."))
        self.update_idletasks()

        id_to_name, name_set = load_scenario_maps()
        stock_map = aggregate_stock(filters, cutoff_iso, id_to_name, name_set)
        if not stock_map:
            self.status_var.set(lang.t("stock_summary.no_data","No data for selected filters (check scenario id/name mismatch or empty stock)."))

        codes = set(k[3] for k in stock_map.keys())
        std_map = load_std_quantities(codes)
        meta = load_item_metadata(codes)
        treecode_map = load_treecodes()
        type_filter = filters["type_filter"].upper()

        for key, data in stock_map.items():
            scenario, kit_number, module_number, code = key
            std_qty = std_map.get(code, 0) or 0
            m = meta.get(code, {"description":"","type":"","remarks":""})
            ctype = m["type"]
            if type_filter in ("KIT","MODULE","ITEM") and ctype.upper() != type_filter:
                continue
            current_stock = data["current_stock"] or 0
            expiring_qty = data["expiring_qty"] or 0
            over_stock = max(0, current_stock - std_qty)
            missing_qty = std_qty - current_stock + (expiring_qty if expiring_qty < std_qty else std_qty)
            if missing_qty < 0: missing_qty = 0
            coverage_pct = ""
            if std_qty > 0:
                coverage_pct = round((current_stock * 100.0) / std_qty, 1)
            earliest = data["earliest_expiry"] or ""
            if earliest and not re.match(r"^\d{4}-\d{2}-\d{2}$", earliest):
                earliest = ""
            modes = ", ".join(sorted(list(data["management_modes"]))) if data["management_modes"] else ""
            self._all_rows.append({
                "scenario": scenario or "",
                "kit_number": kit_number or "",
                "module_number": module_number or "",
                "management_modes": modes,
                "treecode": treecode_map.get(code, ""),
                "code": code,
                "description": m["description"],
                "type": ctype,
                "standard_qty": std_qty,
                "current_stock": current_stock,
                "coverage_pct": coverage_pct,
                "qty_expiring": expiring_qty,
                "earliest_expiry": earliest,
                "over_stock": over_stock,
                "missing_qty": missing_qty,
                "remarks": m["remarks"]
            })

        # ---- Sorting (stable & safe) ----
        def norm_sort(val):
            """
            Returns a string usable for sorting:
              - Empty -> high sentinel
              - Pure numeric (int or float) -> zero-padded numeric string
              - Else -> original string
            """
            if val is None or val == "":
                return "~" * 20  # pushes empties to bottom
            s = str(val).strip()
            # numeric int?
            if re.fullmatch(r"\d+", s):
                return f"INT_{int(s):010d}"
            # numeric float?
            if re.fullmatch(r"\d*\.\d+", s):
                try:
                    f = float(s)
                    return f"FLT_{f:020.6f}"
                except:
                    pass
            return f"TXT_{s}"

        for row in self._all_rows:
            # Precompute sort helpers
            row["_sort_kit"] = norm_sort(row["kit_number"])
            row["_sort_mod"] = norm_sort(row["module_number"])
            ee = row["earliest_expiry"] if row["earliest_expiry"] else "9999-99-99"
            row["_sort_expiry"] = ee

        self._all_rows.sort(key=lambda r: (
            r["scenario"],
            r["_sort_kit"],
            r["_sort_mod"],
            r["_sort_expiry"],
            r["code"]
        ))

        self._render_rows(self._all_rows)
        metrics = self._compute_metrics(self._all_rows)
        self._update_metrics(metrics)
        self.status_var.set(
            lang.t("stock_summary.loaded_status",
                   "Loaded {rows} rows. Expiry Period: {m} month(s) (cutoff {cutoff}) role={role}",
                   rows=len(self._all_rows), m=months, cutoff=cutoff_iso or "N/A", role=self.role)
        )
        self._update_recommended_label()

    def _render_rows(self, rows):
        self.tree.delete(*self.tree.get_children())
        last_group = (None, None, None)
        for r in rows:
            group = (r["scenario"], r["kit_number"], r["module_number"])
            if group != last_group:
                header_text = f"-- {r['scenario'] or 'N/A'} / {r['kit_number'] or 'NoKit'} / {r['module_number'] or 'NoModule'} --"
                self.tree.insert("", "end", values=(
                    r["scenario"], r["kit_number"], r["module_number"],
                    "", "", header_text, "", "", "", "", "", "", "", "", "", ""
                ), tags=("group",))
                last_group = group
            tags = []
            if r["missing_qty"] > 0:
                tags.append("missing")
            elif r["over_stock"] > 0:
                tags.append("overstock")
            self.tree.insert("", "end", values=(
                r["scenario"], r["kit_number"], r["module_number"], r["management_modes"],
                r["treecode"], r["code"], r["description"], r["type"], r["standard_qty"],
                r["current_stock"], r["coverage_pct"], r["qty_expiring"], r["earliest_expiry"],
                r["over_stock"], r["missing_qty"], r["remarks"]
            ), tags=tuple(tags))

    # ---------------- Metrics ----------------
    def _compute_metrics(self, rows):
        total_std = total_stock = total_short = total_over = total_expiring = 0
        for r in rows:
            if r["code"].startswith("-- "):
                continue
            total_std += r["standard_qty"] or 0
            total_stock += r["current_stock"] or 0
            total_short += r["missing_qty"] or 0
            total_over += r["over_stock"] or 0
            total_expiring += r["qty_expiring"] or 0
        coverage = ""
        if total_std > 0:
            coverage = f"{round((total_stock * 100.0)/ total_std,1)}"
        return {
            "total_std": total_std,
            "total_stock": total_stock,
            "total_shortage": total_short,
            "total_overstock": total_over,
            "coverage_pct": coverage,
            "total_expiring": total_expiring
        }

    def _update_metrics(self, m):
        if not m:
            for k in self.metric_labels:
                self.metric_labels[k].config(text=f"{k}: 0")
            return
        self.metric_labels["total_std"].config(text=f"{lang.t('stock_summary.total_std','Total Std')}: {m['total_std']}")
        self.metric_labels["total_stock"].config(text=f"{lang.t('stock_summary.total_stock','Total Stock')}: {m['total_stock']}")
        self.metric_labels["total_shortage"].config(text=f"{lang.t('stock_summary.total_shortage','Total Shortage')}: {m['total_shortage']}")
        self.metric_labels["total_overstock"].config(text=f"{lang.t('stock_summary.total_overstock','Total Over Stock')}: {m['total_overstock']}")
        self.metric_labels["coverage_pct"].config(text=f"{lang.t('stock_summary.coverage','Coverage %')}: {m['coverage_pct']}")
        self.metric_labels["total_expiring"].config(text=f"{lang.t('stock_summary.total_expiring','Total Expiring')}: {m['total_expiring']}")

    # ---------------- Search Filter ----------------
    def _on_search_change(self, *_):
        q = self.search_var.get().strip().lower()
        if not q:
            self._render_rows(self._all_rows)
            metrics = self._compute_metrics(self._all_rows)
            self._update_metrics(metrics)
            return
        filtered = []
        for r in self._all_rows:
            fields = [
                r["scenario"], r["kit_number"], r["module_number"], r["management_modes"],
                r["treecode"], r["code"], r["description"], r["type"], r["remarks"]
            ]
            combined = " ".join([str(x).lower() for x in fields if x])
            if q in combined:
                filtered.append(r)
        self._render_rows(filtered)
        metrics = self._compute_metrics(filtered)
        self._update_metrics(metrics)

    # ---------------- Export ----------------
    def export_excel(self):
        items = self.tree.get_children("")
        if not items:
            custom_popup(self, lang.t("reports.info","Info"), lang.t("reports.nothing_to_export","Nothing to export."), "info")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel","*.xlsx")],
                                            initialfile="stock_summary_grouped.xlsx")
        if not path:
            self.status_var.set(lang.t("reports.export_cancelled","Export cancelled"))
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Stock_Summary"
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws["A1"] = f"{lang.t('reports.generated','Generated')}: {now} (role={self.role})"
            ws["A1"].font = Font(size=10)
            ws.append([])
            headers = [self.tree.heading(c)["text"] for c in self.columns]
            ws.append(headers)
            missing_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            over_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
            group_fill = PatternFill(start_color="D9E1EC", end_color="D9E1EC", fill_type="solid")

            for iid in items:
                vals = self.tree.item(iid, "values")
                tags = set(self.tree.item(iid, "tags") or [])
                ws.append(list(vals))
                rcells = ws[ws.max_row]
                if "group" in tags:
                    for c in rcells:
                        c.fill = group_fill
                        c.font = Font(bold=True)
                elif "missing" in tags:
                    for c in rcells:
                        c.fill = missing_fill
                elif "overstock" in tags:
                    for c in rcells:
                        c.fill = over_fill
            for col in ws.columns:
                max_len = 0
                letter = get_column_letter(col[0].column)
                for cell in col:
                    v = "" if cell.value is None else str(cell.value)
                    if len(v) > max_len:
                        max_len = len(v)
                ws.column_dimensions[letter].width = min(max_len + 2, 55)
            ws.freeze_panes = "A3"
            wb.save(path)
            custom_popup(self, lang.t("reports.success","Success"),
                         lang.t("reports.export_success","Export successful."), "success")
            self.status_var.set(lang.t("reports.export_success","Export successful."))
        except Exception as e:
            custom_popup(self, lang.t("reports.error","Error"),
                         lang.t("reports.export_failed","Export failed: {err}", err=str(e)), "error")
            self.status_var.set(lang.t("reports.export_failed","Export failed: {err}", err=str(e)))

# ----------------------------- Launcher ------------------------------
def open_stock_summary(parent=None, role="user"):
    win = StockSummaryWindow(parent=parent, role=role)
    win.focus()
    return win

# ------------------------- Standalone Run ----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Stock Summary")
    root.geometry("1540x900")
    open_stock_summary(root, role="admin")
    root.mainloop()