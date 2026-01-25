"""
order.py  v1.1

Adjustments from v1.0:
  * FIX: Loan / Borrow balance always zero when no kit/module filters.
        (SQL 'WHERE' + 'AND' issue). Rewritten query builder so loan filter
        works with or without other filters.
  * FIX: Remarks column now reliably accepts & preserves free text (any string).
        (Editing logic refactored; italic tag applied when non‑empty).
  * CHANGE: qty_needed is now clamped to a minimum of 0 (cannot go negative).
  * IMPROVED: ComboBox (kit/module/type) & item search Enter now auto refresh.
  * ADDED: Auto refresh also on leaving (FocusOut) of lead/cover/buffer months
           and immediate refresh on <<ComboboxSelected>> (instead of needing
           manual Refresh).
  * EDIT: Editing while in Simple mode: only columns present are editable. 
          Remarks only editable in Detailed mode (column not shown in Simple).
  * SAFETY: qty_to_order when blank defaults to qty_needed (already) AFTER clamp.
  * TAG: Rows with remarks use italic font (tag 'remarks_italic').
  * CODE: Minor refactors & inline comments for new/changed sections.
"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog
import sqlite3
import math
from datetime import date, datetime
from calendar import monthrange
import re
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font
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
# BTN_EXPORT     = "#2980B9"
# BTN_REFRESH    = "#2563EB"
# BTN_CLEAR      = "#7F8C8D"
# BTN_TOGGLE     = "#8E44AD"

# ============================================================
# KEPT VISUALIZATION-SPECIFIC COLORS (Excel fill specific)
# ============================================================
KIT_FILL_COLOR    = "228B22"    # Excel fill color (no #)
MODULE_FILL_COLOR = "ADD8E6"    # Excel fill color (no #)

LOAN_OUT_TYPES = {"Loan","Return of Borrowing"}
LOAN_IN_TYPES  = {"In Borrowing","In Return of Loan"}

INT_ENTRY_MIN = 0
INT_ENTRY_MAX = 24

def add_months(base: date, months: int) -> date:
    if months <= 0:
        return base
    y = base.year + (base.month - 1 + months) // 12
    m = (base.month - 1 + months) % 12 + 1
    d = base.day
    last = monthrange(y,m)[1]
    if d > last: d = last
    return date(y,m,d)

def fetch_project_settings():
    conn = connect_db()
    if conn is None:
        return 0,0,0
    cur = conn.cursor()
    lead=cover=buf=0
    try:
        cur.execute("PRAGMA table_info(project_details)")
        cols = {r[1].lower(): r[1] for r in cur.fetchall()}
        needed = []
        if "lead_time_months" in cols: needed.append("lead_time_months")
        if "cover_period_months" in cols: needed.append("cover_period_months")
        if "buffer_months" in cols: needed.append("buffer_months")
        if not needed:
            return 0,0,0
        cur.execute(f"SELECT {', '.join(needed)} FROM project_details LIMIT 1")
        row = cur.fetchone()
        if row:
            idx=0
            if "lead_time_months" in needed:
                lead=int(row[idx] or 0); idx+=1
            if "cover_period_months" in needed:
                cover=int(row[idx] or 0); idx+=1
            if "buffer_months" in needed:
                buf=int(row[idx] or 0)
    except Exception:
        pass
    finally:
        cur.close(); conn.close()
    return tuple(max(0,min(24,x)) for x in (lead,cover,buf))

# ---------- Data Aggregator ----------
class OrderData:
    def __init__(self, kit_filter, module_filter, type_filter,
                 item_search, lead, cover, buffer):
        self.kit_filter = kit_filter
        self.module_filter = module_filter
        self.type_filter = type_filter
        self.item_search = (item_search or "").strip()
        self.lead = lead
        self.cover = cover
        self.buffer = buffer
        self.horizon_months = (lead or 0) + (cover or 0) + (buffer or 0)
        self.horizon_end = add_months(date.today(), self.horizon_months)

    def fetch(self):
        std_map = self._fetch_standard_qty()
        stock_map = self._fetch_current_stock()
        exp_map = self._fetch_expiring_qty()
        loan_map = self._fetch_loan_balance()
        commercial = self._fetch_commercial_data(set().union(std_map, stock_map, exp_map, loan_map))

        codes = set(std_map) | set(stock_map) | set(exp_map) | set(loan_map)
        rows=[]
        for code in sorted(codes):
            desc = get_item_description(code)
            dtype = detect_type(code, desc)
            # Type filter
            if self.type_filter and self.type_filter.lower() != "all":
                if dtype.lower() != self.type_filter.lower():
                    continue
            # Item search only on items
            if self.item_search:
                if dtype.lower() != "item":
                    continue
                if self.item_search.lower() not in code.lower() and \
                   self.item_search.lower() not in desc.lower():
                    continue
            cdat = commercial.get(code, {})
            row = {
                "code": code,
                "description": desc,
                "type": dtype,
                "standard_qty": std_map.get(code,0),
                "current_stock": stock_map.get(code,0),
                "qty_expiring": exp_map.get(code,0),
                "back_orders": 0,
                "loan_balance": loan_map.get(code,0),
                "planned_dons_give": 0,
                "dons_receive": 0,
                "pack_size": cdat.get("pack_size",0),
                "qty_needed": 0,
                "qty_to_order": "",
                "qty_to_order_rounded": 0,
                "price_per_pack": cdat.get("price",0.0),
                "weight_per_pack": cdat.get("weight",0.0),
                "volume_per_pack_dm3": cdat.get("volume",0.0),
                "amount": 0.0,
                "weight_kg": 0.0,
                "volume_m3": 0.0,
                "account_code": cdat.get("account",""),
                "remarks": ""
            }
            rows.append(row)
        return rows

    def _fetch_standard_qty(self):
        res=defaultdict(int)
        conn=connect_db()
        if not conn: return res
        cur=conn.cursor()
        try:
            cur.execute("PRAGMA table_info(std_qty_helper)")
            cols={r[1].lower(): r[1] for r in cur.fetchall()}
            if "code" not in cols or "std_qty" not in cols:
                return res
            filters=[]
            params=[]
            if self.kit_filter and self.kit_filter.lower()!="all" and "kit" in cols:
                filters.append("kit=?"); params.append(self.kit_filter)
            if self.module_filter and self.module_filter.lower()!="all" and "module" in cols:
                filters.append("module=?"); params.append(self.module_filter)
            where = "WHERE "+ " AND ".join(filters) if filters else ""
            cur.execute(f"SELECT code,std_qty FROM std_qty_helper {where}", params)
            for c,q in cur.fetchall():
                if c:
                    try: res[c]+=int(q or 0)
                    except: pass
        finally:
            cur.close(); conn.close()
        return res

    def _fetch_current_stock(self):
        res=defaultdict(int)
        conn=connect_db()
        if not conn: return res
        cur=conn.cursor()
        try:
            cur.execute("PRAGMA table_info(stock_data)")
            cols={r[1].lower(): r[1] for r in cur.fetchall()}
            if "final_qty" not in cols:
                return res
            has_code = "code" in cols
            code_col = "code" if has_code else "unique_id"
            filters=["final_qty IS NOT NULL"]
            params=[]
            if self.kit_filter and self.kit_filter.lower()!="all" and "kit_number" in cols:
                filters.append("kit_number=?"); params.append(self.kit_filter)
            if self.module_filter and self.module_filter.lower()!="all" and "module_number" in cols:
                filters.append("module_number=?"); params.append(self.module_filter)
            where="WHERE "+ " AND ".join(filters)
            cur.execute(f"SELECT {code_col}, final_qty FROM stock_data {where}", params)
            for c,q in cur.fetchall():
                code = c if has_code else self._extract_code_from_unique_id(c)
                if code:
                    try: res[code]+=int(q or 0)
                    except: pass
        finally:
            cur.close(); conn.close()
        return res

    def _fetch_expiring_qty(self):
        res=defaultdict(int)
        if self.horizon_months<=0: return res
        conn=connect_db()
        if not conn: return res
        cur=conn.cursor()
        try:
            cur.execute("PRAGMA table_info(stock_data)")
            cols={r[1].lower(): r[1] for r in cur.fetchall()}
            if not {"final_qty","exp_date"} <= set(cols):
                return res
            has_code="code" in cols
            code_col="code" if has_code else "unique_id"
            horizon_end = self.horizon_end.strftime("%Y-%m-%d")
            filters=["final_qty IS NOT NULL","exp_date IS NOT NULL","exp_date!=''","exp_date <= ?"]
            params=[horizon_end]
            if self.kit_filter and self.kit_filter.lower()!="all" and "kit_number" in cols:
                filters.append("kit_number=?"); params.append(self.kit_filter)
            if self.module_filter and self.module_filter.lower()!="all" and "module_number" in cols:
                filters.append("module_number=?"); params.append(self.module_filter)
            where="WHERE "+ " AND ".join(filters)
            cur.execute(f"SELECT {code_col}, final_qty FROM stock_data {where}", params)
            for c,q in cur.fetchall():
                code = c if has_code else self._extract_code_from_unique_id(c)
                if code:
                    try: res[code]+=int(q or 0)
                    except: pass
        finally:
            cur.close(); conn.close()
        return res

    def _fetch_loan_balance(self):
        """
        Correct WHERE handling even with no kit/module filters.
        """
        res=defaultdict(int)
        conn=connect_db()
        if not conn: return res
        cur=conn.cursor()
        try:
            cur.execute("PRAGMA table_info(stock_transactions)")
            cols={r[1].lower(): r[1] for r in cur.fetchall()}
            if not {"qty_in","qty_out","in_type","out_type","code"} <= set(cols):
                return res
            filters=[]
            params=[]
            if self.kit_filter and self.kit_filter.lower()!="all" and "kit" in cols:
                filters.append("Kit=?"); params.append(self.kit_filter)
            if self.module_filter and self.module_filter.lower()!="all" and "module" in cols:
                filters.append("Module=?"); params.append(self.module_filter)

            loan_condition = f"(Out_Type IN ({','.join('?'*len(LOAN_OUT_TYPES))}) OR IN_Type IN ({','.join('?'*len(LOAN_IN_TYPES))}))"
            params_all = params + list(LOAN_OUT_TYPES) + list(LOAN_IN_TYPES)
            if filters:
                where = "WHERE " + " AND ".join(filters) + f" AND {loan_condition}"
            else:
                where = "WHERE " + loan_condition

            cur.execute(f"""
                SELECT code, Qty_IN, IN_Type, Qty_Out, Out_Type
                FROM stock_transactions
                {where}
            """, params_all)
            for code, q_in, in_type, q_out, out_type in cur.fetchall():
                if not code: continue
                if out_type in LOAN_OUT_TYPES and q_out:
                    try: res[code]+=int(q_out)
                    except: pass
                if in_type in LOAN_IN_TYPES and q_in:
                    try: res[code]-=int(q_in)
                    except: pass
        except Exception:
            pass
        finally:
            cur.close(); conn.close()
        return res

    def _fetch_commercial_data(self, codes):
        res={}
        if not codes: return res
        conn=connect_db()
        if not conn: return res
        cur=conn.cursor()
        try:
            cur.execute("PRAGMA table_info(items_list)")
            cols={r[1].lower(): r[1] for r in cur.fetchall()}
            if "code" not in cols:
                return res
            field_map={
                "pack":"pack_size",
                "price_per_pack_euros":"price",
                "weight_per_pack_kg":"weight",
                "volume_per_pack_dm3":"volume",
                "account_code":"account"
            }
            sel=["code"]+[c for c in field_map if c in cols]
            placeholders=",".join("?" for _ in codes)
            cur.execute(f"SELECT {', '.join(sel)} FROM items_list WHERE code IN ({placeholders})",
                        list(codes))
            idx={name:i for i,name in enumerate(sel)}
            for row in cur.fetchall():
                code=row[idx["code"]]
                data={}
                for src,dst in field_map.items():
                    if src in idx:
                        val=row[idx[src]]
                        try:
                            if dst in ("pack_size",):
                                data[dst]=int(val or 0)
                            elif dst in ("price","weight","volume"):
                                data[dst]=float(val or 0)
                            else:
                                data[dst]=val or ""
                        except:
                            data[dst]=0 if dst!="account" else ""
                res[code]=data
        finally:
            cur.close(); conn.close()
        return res

    @staticmethod
    def _extract_code_from_unique_id(u):
        # scenario/kit/module/item/std/exp/kit_no/module_no
        try:
            p=u.split("/")
            if len(p)>=4 and p[3]!="None": return p[3]
            if len(p)>=3 and p[2]!="None": return p[2]
            if len(p)>=2 and p[1]!="None": return p[1]
            return u
        except:
            return u

# ---------- UI (UPDATED: All color references use AppTheme) ----------
class OrderNeeds(tk.Frame):
    SIMPLE_COLS = ["code","description","standard_qty","current_stock",
                   "qty_needed","qty_to_order_rounded","amount","weight_kg","volume_m3"]
    DETAIL_COLS = [
        "code","description","type","standard_qty","current_stock","qty_expiring",
        "back_orders","loan_balance","planned_dons_give","dons_receive",
        "pack_size","qty_needed","qty_to_order","qty_to_order_rounded",
        "price_per_pack","weight_per_pack","volume_per_pack_dm3",
        "amount","weight_kg","volume_m3","account_code","remarks"
    ]
    EDITABLE_COLS = {"back_orders","planned_dons_give","dons_receive","qty_to_order","remarks"}

    def __init__(self, parent, app, *args, **kwargs):
        super().__init__(parent, bg=AppTheme.BG_MAIN, *args, **kwargs)  # UPDATED: Use AppTheme
        self.app = app
        self.rows=[]
        self.simple_mode=False

        self.kit_var = tk.StringVar(value="All")
        self.module_var = tk.StringVar(value="All")
        self.type_var = tk.StringVar(value="All")
        self.item_search_var = tk.StringVar()
        self.lead_var = tk.StringVar()
        self.cover_var = tk.StringVar()
        self.buffer_var = tk.StringVar()

        self.total_amount_var = tk.StringVar(value="0.00")
        self.total_weight_var = tk.StringVar(value="0.00")
        self.total_volume_var = tk.StringVar(value="0.000")
        self.missing_price_var = tk.StringVar(value="")

        self.tree=None
        self.edit_entry=None
        self.status_var = tk.StringVar(value=lang.t("order_needs.ready","Ready"))

        self._build_ui()
        self._load_initial_params()
        self.populate_dropdowns()
        self.refresh()

    # ----- UI Build (UPDATED: All color references use AppTheme) -----
    def _build_ui(self):
        header=tk.Frame(self, bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        header.pack(fill="x",padx=12,pady=(12,6))
        tk.Label(header,text=lang.t("menu.reports.order_needs","Order / Needs"),
                 font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HUGE, "bold"),  # UPDATED: Use AppTheme
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY)\
            .pack(side="left")  # UPDATED: Use AppTheme
        self.toggle_btn=tk.Button(header,text=lang.t("order_needs.detailed","Detailed"),
                                  bg=AppTheme.BTN_TOGGLE,  # UPDATED: Use AppTheme
                                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                                  relief="flat",
                                  padx=14,pady=6,command=self.toggle_mode)
        self.toggle_btn.pack(side="right",padx=(6,0))
        tk.Button(header,text=lang.t("generic.export","Export"),
                  bg=AppTheme.BTN_EXPORT,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14,pady=6,command=self.export_excel)\
            .pack(side="right",padx=(6,0))
        tk.Button(header,text=lang.t("generic.clear","Clear"),
                  bg=AppTheme.BTN_NEUTRAL,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14,pady=6,command=self.clear_all)\
            .pack(side="right",padx=(6,0))
        tk.Button(header,text=lang.t("generic.refresh","Refresh"),
                  bg=AppTheme.BTN_REFRESH,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14,pady=6,command=self.refresh)\
            .pack(side="right",padx=(6,0))

        # Filters
        filt=tk.Frame(self, bg=AppTheme.BG_MAIN); filt.pack(fill="x",padx=12,pady=(0,6))  # UPDATED: Use AppTheme
        r1=tk.Frame(filt, bg=AppTheme.BG_MAIN); r1.pack(fill="x",pady=2)  # UPDATED: Use AppTheme
        tk.Label(r1,text=lang.t("generic.kit_number","Kit Number"),bg=AppTheme.BG_MAIN).grid(row=0,column=0,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.kit_cb=ttk.Combobox(r1,textvariable=self.kit_var,state="readonly",width=18)
        self.kit_cb.grid(row=0,column=1,padx=(0,12))
        self.kit_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        tk.Label(r1,text=lang.t("generic.module_number","Module Number"),bg=AppTheme.BG_MAIN).grid(row=0,column=2,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.module_cb=ttk.Combobox(r1,textvariable=self.module_var,state="readonly",width=18)
        self.module_cb.grid(row=0,column=3,padx=(0,12))
        self.module_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        tk.Label(r1,text=lang.t("generic.type","Type"),bg=AppTheme.BG_MAIN).grid(row=0,column=4,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.type_cb=ttk.Combobox(r1,textvariable=self.type_var,state="readonly",width=12,
                                  values=["All","Kit","Module","Item"])
        self.type_cb.grid(row=0,column=5,padx=(0,12))
        self.type_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        tk.Label(r1,text=lang.t("order_needs.item_search","Item Search"),bg=AppTheme.BG_MAIN).grid(row=0,column=6,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.item_entry=tk.Entry(r1,textvariable=self.item_search_var,width=20)
        self.item_entry.grid(row=0,column=7,padx=(0,12))
        self.item_entry.bind("<Return>", lambda e: self.refresh())
        self.item_entry.bind("<Escape>", lambda e: self._clear_field(self.item_search_var))

        # Parameters
        r2=tk.Frame(filt, bg=AppTheme.BG_MAIN); r2.pack(fill="x",pady=2)  # UPDATED: Use AppTheme
        tk.Label(r2,text=lang.t("order_needs.lead_time","Lead Time (months)"),bg=AppTheme.BG_MAIN).grid(row=0,column=0,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.lead_entry=tk.Entry(r2,textvariable=self.lead_var,width=6)
        self.lead_entry.grid(row=0,column=1,padx=(0,12))
        self.lead_entry.bind("<FocusOut>", lambda e: self._param_focus_refresh())
        self.lead_entry.bind("<Return>", lambda e: self._param_focus_refresh())

        tk.Label(r2,text=lang.t("order_needs.cover_period","Cover Period (months)"),bg=AppTheme.BG_MAIN).grid(row=0,column=2,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.cover_entry=tk.Entry(r2,textvariable=self.cover_var,width=6)
        self.cover_entry.grid(row=0,column=3,padx=(0,12))
        self.cover_entry.bind("<FocusOut>", lambda e: self._param_focus_refresh())
        self.cover_entry.bind("<Return>", lambda e: self._param_focus_refresh())

        tk.Label(r2,text=lang.t("order_needs.buffer","Security Stock (buffer, months)"),bg=AppTheme.BG_MAIN).grid(row=0,column=4,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        self.buffer_entry=tk.Entry(r2,textvariable=self.buffer_var,width=6)
        self.buffer_entry.grid(row=0,column=5,padx=(0,12))
        self.buffer_entry.bind("<FocusOut>", lambda e: self._param_focus_refresh())
        self.buffer_entry.bind("<Return>", lambda e: self._param_focus_refresh())

        # Totals Row (UPDATED: Colors use AppTheme)
        info=tk.Frame(self, bg=AppTheme.BG_MAIN); info.pack(fill="x",padx=12,pady=(2,4))  # UPDATED: Use AppTheme
        tk.Label(info,text=lang.t("order_needs.total_amount","Total Amount (€):"),
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                 font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"))\
            .grid(row=0,column=0,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        tk.Label(info,textvariable=self.total_amount_var,
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY)\
            .grid(row=0,column=1,padx=(0,14),sticky="w")  # UPDATED: Use AppTheme

        tk.Label(info,text=lang.t("order_needs.total_weight","Total Weight (kg):"),
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                 font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"))\
            .grid(row=0,column=2,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        tk.Label(info,textvariable=self.total_weight_var,
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY)\
            .grid(row=0,column=3,padx=(0,14),sticky="w")  # UPDATED: Use AppTheme

        tk.Label(info,text=lang.t("order_needs.total_volume","Total Volume (m3):"),
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                 font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"))\
            .grid(row=0,column=4,sticky="w",padx=(0,4))  # UPDATED: Use AppTheme
        tk.Label(info,textvariable=self.total_volume_var,
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY)\
            .grid(row=0,column=5,padx=(0,14),sticky="w")  # UPDATED: Use AppTheme

        tk.Label(info,textvariable=self.missing_price_var,
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg="#B45309",font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_SMALL, "italic")).grid(row=0,column=6,sticky="w")  # UPDATED: Use AppTheme

        tk.Label(self,textvariable=self.status_var,anchor="w",
                 bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                 fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                 relief="sunken")\
            .pack(fill="x",padx=12,pady=(0,8))

        # Table (UPDATED: Colors use AppTheme)
        frame = tk.Frame(self, bg=AppTheme.COLOR_BORDER, bd=1, relief="solid")  # UPDATED: Use AppTheme
        frame.pack(fill="both",expand=True,padx=12,pady=(0,12))
        self.tree=ttk.Treeview(frame,columns=(),show="headings",height=24)
        self.tree.pack(side="left",fill="both",expand=True)
        vsb=ttk.Scrollbar(frame,orient="vertical",command=self.tree.yview)
        vsb.pack(side="right",fill="y")
        hsb=ttk.Scrollbar(self,orient="horizontal",command=self.tree.xview)
        hsb.pack(fill="x",padx=12,pady=(0,12))
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)

        # Style configuration (UPDATED: Removed theme_use, using AppTheme)
        style=ttk.Style()
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
        self.tree.tag_configure("remarks_italic", font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "italic"))  # UPDATED: Use AppTheme

        self.tree.bind("<Double-1>", self._start_edit)
        self.tree.bind("<Button-1>", self._click_cancel_edit)

        self.bind_all("<Escape>", self._global_esc)

    # ----- Parameter change helper -----
    def _param_focus_refresh(self):
        self._validate_int_field(self.lead_var)
        self._validate_int_field(self.cover_var)
        self._validate_int_field(self.buffer_var)
        self.refresh()

    # ----- Initial params -----
    def _load_initial_params(self):
        lead,cover,buffer = fetch_project_settings()
        self.lead_var.set(str(lead))
        self.cover_var.set(str(cover))
        self.buffer_var.set(str(buffer))

    # ----- Dropdown population -----
    def populate_dropdowns(self):
        self.kit_cb['values']=["All"]+self._distinct("stock_data","kit_number")
        self.module_cb['values']=["All"]+self._distinct("stock_data","module_number")

    def _distinct(self,table,column):
        conn=connect_db()
        if not conn: return []
        cur=conn.cursor()
        try:
            cur.execute(f"PRAGMA table_info({table})")
            cols={r[1].lower(): r[1] for r in cur.fetchall()}
            if column.lower() not in cols: return []
            cur.execute(f"""
                SELECT DISTINCT {column} FROM {table}
                WHERE {column} IS NOT NULL AND {column}!='' AND {column}!='None'
                ORDER BY {column}
            """)
            return [r[0] for r in cur.fetchall()]
        except:
            return []
        finally:
            cur.close(); conn.close()

    # ----- Refresh -----
    def refresh(self):
        od = OrderData(
            kit_filter=self.kit_var.get(),
            module_filter=self.module_var.get(),
            type_filter=self.type_var.get(),
            item_search=self.item_search_var.get(),
            lead=self._safe_int(self.lead_var.get()),
            cover=self._safe_int(self.cover_var.get()),
            buffer=self._safe_int(self.buffer_var.get())
        )
        self.rows = od.fetch()
        for r in self.rows:
            self._recompute_row(r)
        self._populate_tree()
        self._recompute_totals()
        self.status_var.set(
            lang.t("order_needs.loaded","Loaded {n} rows").format(n=len(self.rows))
        )

    def clear_all(self):
        self.kit_var.set("All")
        self.module_var.set("All")
        self.type_var.set("All")
        self.item_search_var.set("")
        self._load_initial_params()
        self.refresh()

    # ----- Mode -----
    def toggle_mode(self):
        self.simple_mode = not self.simple_mode
        self.toggle_btn.config(
            text=lang.t("order_needs.simple","Simple") if self.simple_mode
            else lang.t("order_needs.detailed","Detailed")
        )
        self._populate_tree()
        self._recompute_totals()

    # ----- Columns / Populate -----
    def _current_columns(self):
        return self.SIMPLE_COLS if self.simple_mode else self.DETAIL_COLS

    def _populate_tree(self):
        cols=self._current_columns()
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"]=cols

        headers={
            "code": lang.t("generic.code","Code"),
            "description": lang.t("generic.description","Description"),
            "type": lang.t("generic.type","Type"),
            "standard_qty": lang.t("order_needs.standard_qty","Standard Qty"),
            "current_stock": lang.t("order_needs.current_stock","Current Stock"),
            "qty_expiring": lang.t("order_needs.qty_expiring","Qty Expiring"),
            "back_orders": lang.t("order_needs.back_orders","Back Orders"),
            "loan_balance": lang.t("order_needs.loan_balance","Loan/Borrow Balance"),
            "planned_dons_give": lang.t("order_needs.planned_dons_give","Planned Donations (Give)"),
            "dons_receive": lang.t("order_needs.dons_receive","Donations To Receive"),
            "pack_size": lang.t("order_needs.pack_size","Pack Size"),
            "qty_needed": lang.t("order_needs.qty_needed","Qty Needed"),
            "qty_to_order": lang.t("order_needs.qty_to_order","Qty To Order"),
            "qty_to_order_rounded": lang.t("order_needs.qty_to_order_rounded","Order Rounded"),
            "price_per_pack": lang.t("order_needs.price_per_pack","Price/Pack (€)"),
            "weight_per_pack": lang.t("order_needs.weight_per_pack","Weight/Pack (kg)"),
            "volume_per_pack_dm3": lang.t("order_needs.volume_per_pack","Volume/Pack (dm3)"),
            "amount": lang.t("order_needs.amount","Amount (€)"),
            "weight_kg": lang.t("order_needs.weight","Weight (kg)"),
            "volume_m3": lang.t("order_needs.volume_m3","Volume (m3)"),
            "account_code": lang.t("order_needs.account_code","Account Code"),
            "remarks": lang.t("order_needs.remarks","Remarks")
        }
        widths={
            "code":170,"description":340,"type":90,"standard_qty":110,"current_stock":110,
            "qty_expiring":110,"back_orders":110,"loan_balance":140,"planned_dons_give":160,
            "dons_receive":140,"pack_size":90,"qty_needed":110,"qty_to_order":120,
            "qty_to_order_rounded":140,"price_per_pack":110,"weight_per_pack":130,
            "volume_per_pack_dm3":140,"amount":110,"weight_kg":110,"volume_m3":110,
            "account_code":130,"remarks":260
        }
        for c in cols:
            self.tree.heading(c,text=headers.get(c,c))
            self.tree.column(c,width=widths.get(c,120),anchor="w",stretch=True)

        for idx,r in enumerate(self.rows):
            vals=[self._format_cell(r,c) for c in cols]
            tag="alt" if idx%2 else ""
            dtype=r.get("type","").upper()
            if dtype=="KIT": tag="kitrow"
            elif dtype=="MODULE": tag="modrow"
            if r.get("remarks") and "remarks" in cols:
                # Add italic tag in addition to base
                self.tree.insert("", "end", values=vals, tags=(tag,"remarks_italic"))
            else:
                self.tree.insert("", "end", values=vals, tags=(tag,))

    def _format_cell(self,row,col):
        v=row.get(col,"")
        if col in ("amount","price_per_pack"):
            return f"{float(v):.2f}"
        if col in ("weight_kg","weight_per_pack","volume_per_pack_dm3"):
            return f"{float(v):.3f}"
        if col=="volume_m3":
            return f"{float(v):.4f}"
        if col=="qty_to_order" and v=="":
            return ""
        return v

    # ----- Editing -----
    def _click_cancel_edit(self, event):
        if self.edit_entry:
            try: self.edit_entry.destroy()
            except: pass
            self.edit_entry=None

    def _start_edit(self,event):
        if self.edit_entry:
            try: self.edit_entry.destroy()
            except: pass
            self.edit_entry=None
        region=self.tree.identify("region",event.x,event.y)
        if region!="cell": return
        row_id=self.tree.identify_row(event.y)
        col_id=self.tree.identify_column(event.x)
        if not row_id or not col_id: return
        col_index=int(col_id.replace("#",""))-1
        cols=self._current_columns()
        if col_index<0 or col_index>=len(cols): return
        col_name=cols[col_index]
        if col_name not in self.EDITABLE_COLS:
            return
        # Remarks only present in detailed mode; safe.
        bbox=self.tree.bbox(row_id,col_id)
        if not bbox: return
        x,y,w,h=bbox
        value=self.tree.set(row_id,col_name)
        entry=tk.Entry(self.tree,font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "italic" if col_name=="remarks" else "normal"))  # UPDATED: Use AppTheme
        entry.place(x=x,y=y,width=w,height=h)
        entry.insert(0,value)
        entry.focus()
        self.edit_entry=entry
        def save(_=None):
            nv=entry.get()
            self._apply_edit(row_id,col_name,nv)
            try: entry.destroy()
            except: pass
            self.edit_entry=None
        def cancel(_=None):
            try: entry.destroy()
            except: pass
            self.edit_entry=None
        entry.bind("<Return>",save)
        entry.bind("<Escape>",cancel)
        entry.bind("<FocusOut>",save)

    def _apply_edit(self,iid,col_name,new_val):
        code_index=self._current_columns().index("code")
        code=self.tree.item(iid,"values")[code_index]
        row=next((r for r in self.rows if r["code"]==code),None)
        if not row: return
        if col_name in {"back_orders","planned_dons_give","dons_receive"}:
            if new_val.strip()=="":
                row[col_name]=0
            elif re.match(r"^-?\d+$",new_val.strip()):
                row[col_name]=int(new_val.strip())
            else:
                custom_popup(self, lang.t("generic.error","Error"),
                             lang.t("order_needs.invalid_int","Enter whole integer."),"error")
                return
        elif col_name=="qty_to_order":
            if new_val.strip()=="":
                row[col_name]=""
            elif re.match(r"^-?\d+$",new_val.strip()):
                row[col_name]=int(new_val.strip())
            else:
                custom_popup(self, lang.t("generic.error","Error"),
                             lang.t("order_needs.invalid_int","Enter whole integer."),"error")
                return
        elif col_name=="remarks":
            row[col_name]=new_val
        self._recompute_row(row)
        self._populate_tree()
        self._recompute_totals()

    # ----- Recompute Row -----
    def _recompute_row(self,r):
        std=r.get("standard_qty",0)
        cur=r.get("current_stock",0)
        exp=r.get("qty_expiring",0)
        back=r.get("back_orders",0)
        loan=r.get("loan_balance",0)
        give=r.get("planned_dons_give",0)
        rec=r.get("dons_receive",0)
        qty_needed = std - cur + exp - back - loan + give - rec
        if qty_needed < 0:
            qty_needed = 0  # clamp as requested
        r["qty_needed"]=qty_needed
        q_to=r.get("qty_to_order","")
        if q_to=="":
            q_to=qty_needed
        r["qty_to_order"]=q_to
        pack=r.get("pack_size",0)
        if pack and pack>0:
            rrounded=math.ceil(q_to/pack)*pack
        else:
            rrounded=q_to
        r["qty_to_order_rounded"]=rrounded
        pack_div=pack if pack and pack>0 else None
        price=float(r.get("price_per_pack") or 0)
        wt=float(r.get("weight_per_pack") or 0)
        vol_dm3=float(r.get("volume_per_pack_dm3") or 0)
        if pack_div:
            packs=rrounded/pack_div
            amount=packs*price
            weight=packs*wt
            volume_m3=(packs*vol_dm3)/1000
        else:
            amount=0.0; weight=0.0; volume_m3=0.0
        r["amount"]=amount
        r["weight_kg"]=weight
        r["volume_m3"]=volume_m3

    # ----- Totals -----
    def _recompute_totals(self):
        total_amount=sum(r.get("amount",0) for r in self.rows)
        total_weight=sum(r.get("weight_kg",0) for r in self.rows)
        total_volume=sum(r.get("volume_m3",0) for r in self.rows)
        missing=sum(1 for r in self.rows if (r.get("price_per_pack") or 0)==0)
        self.total_amount_var.set(f"{total_amount:,.2f}")
        self.total_weight_var.set(f"{total_weight:,.2f}")
        self.total_volume_var.set(f"{total_volume:,.3f}")
        if missing:
            self.missing_price_var.set(
                lang.t("order_needs.missing_prices",
                       fallback="{n} items have missing price (0€ used)").format(n=missing)
            )
        else:
            self.missing_price_var.set("")

    # ----- Utilities -----
    def _validate_int_field(self,var):
        val=var.get().strip()
        if val=="" or not val.isdigit():
            var.set("0"); return
        iv=int(val)
        if iv<INT_ENTRY_MIN: iv=INT_ENTRY_MIN
        if iv>INT_ENTRY_MAX: iv=INT_ENTRY_MAX
        var.set(str(iv))

    def _clear_field(self,var):
        var.set("")
        self.refresh()

    def _global_esc(self,event):
        if isinstance(event.widget, tk.Entry) and event.widget in (
            self.item_entry,self.lead_entry,self.cover_entry,self.buffer_entry
        ):
            return
        if self.edit_entry:
            try: self.edit_entry.destroy()
            except: pass
            self.edit_entry=None
        else:
            self.clear_all()

    @staticmethod
    def _safe_int(txt,default=0):
        try: return int(txt)
        except: return default

    # ----- Export -----
    def export_excel(self):
        if not self.rows:
            custom_popup(self, lang.t("generic.info","Info"),
                         lang.t("order_needs.no_data_export","Nothing to export."),"warning")
            return
        path=filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files","*.xlsx")],
            title=lang.t("order_needs.export_title","Save Order/Needs Report"),
            initialfile=f"OrderNeeds_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not path: return
        try:
            wb=openpyxl.Workbook()
            ws=wb.active
            ws.title="OrderNeeds"
            now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws.append([lang.t("generic.generated","Generated"), now])
            ws.append([lang.t("generic.filters_used","Filters Used")])
            ws.append(["Kit", self.kit_var.get(), "Module", self.module_var.get(), "Type", self.type_var.get()])
            ws.append(["Item Search", self.item_search_var.get(),
                       "Lead", self.lead_var.get(),
                       "Cover", self.cover_var.get()])
            ws.append(["Buffer", self.buffer_var.get(),
                       "Mode", "Simple" if self.simple_mode else "Detailed"])
            ws.append([])
            ws.append(["Total Amount (€)", self.total_amount_var.get(),
                       "Total Weight (kg)", self.total_weight_var.get(),
                       "Total Volume (m3)", self.total_volume_var.get(),
                       "Missing Price Rows", self.missing_price_var.get()])
            ws.append([])

            cols=self._current_columns()
            ws.append([c.replace("_"," ").title() for c in cols])

            kit_fill=PatternFill(start_color=KIT_FILL_COLOR,end_color=KIT_FILL_COLOR,fill_type="solid")
            module_fill=PatternFill(start_color=MODULE_FILL_COLOR,end_color=MODULE_FILL_COLOR,fill_type="solid")

            for r in self.rows:
                row_out=[]
                for c in cols:
                    val=r.get(c,"")
                    if c in ("amount","price_per_pack","weight_kg","weight_per_pack",
                             "volume_per_pack_dm3","volume_m3"):
                        val=float(val)
                    row_out.append(val)
                ws.append(row_out)
                dtype=r.get("type","").upper()
                if dtype=="KIT":
                    for cell in ws[ws.max_row]: cell.fill=kit_fill
                elif dtype=="MODULE":
                    for cell in ws[ws.max_row]: cell.fill=module_fill
                if "remarks" in cols:
                    try:
                        idx=cols.index("remarks")
                        ws.cell(row=ws.max_row,column=idx+1).font=Font(italic=True)
                    except: pass

            # Autosize
            for col in ws.columns:
                length=0
                letter=get_column_letter(col[0].column)
                for cell in col:
                    val="" if cell.value is None else str(cell.value)
                    if len(val)>length: length=len(val)
                ws.column_dimensions[letter].width=min(length+2,60)
            ws.freeze_panes="A10"
            wb.save(path)
            custom_popup(self, lang.t("generic.success","Success"),
                         lang.t("order_needs.export_success","Export completed: {f}").format(f=path),
                         "info")
        except Exception as e:
            custom_popup(self, lang.t("generic.error","Error"),
                         lang.t("order_needs.export_fail","Export failed: {err}").format(err=str(e)),
                         "error")

# ---------- Module Export ----------
__all__=["OrderNeeds"]

if __name__=="__main__":
    root=tk.Tk()
    root.title("Order / Needs v1.1")
    OrderNeeds(root,None).pack(fill="both",expand=True)
    root.geometry("1850x880")
    root.mainloop()