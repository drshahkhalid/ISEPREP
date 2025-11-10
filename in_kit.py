import tkinter as tk
from tkinter import ttk, filedialog, simpledialog
from datetime import datetime
import sqlite3
import re
from calendar import monthrange
from collections import defaultdict
import logging
import openpyxl
from openpyxl.utils import get_column_letter
from tkinter import messagebox as mb
from popup_utils import custom_popup, custom_askyesno, custom_dialog
import os

from db import connect_db
from manage_items import get_item_description, detect_type
from language_manager import lang
from popup_utils import custom_popup

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

# ---------------------------------------------------------------
# Dialog centering helper
# ---------------------------------------------------------------
def _center_child(win: tk.Toplevel, parent: tk.Widget):
    win.update_idletasks()
    try:
        if parent and parent.winfo_exists():
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            w, h = win.winfo_width(), win.winfo_height()
            x = px + (pw // 2) - (w // 2)
            y = py + (ph // 2) - (h // 2)
        else:
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            w, h = win.winfo_width(), win.winfo_height()
            x = (sw // 2) - (w // 2)
            y = (sh // 2) - (h // 2)
        x = max(x, 0)
        y = max(y, 0)
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass

# Center tkinter.simpledialog (monkey patch once)
import tkinter.simpledialog as _sd
if not hasattr(_sd, "_CenteredQueryString"):
    _OrigQS = _sd._QueryString
    class _CenteredQueryString(_OrigQS):
        def body(self, master):
            res = super().body(master)
            _center_child(self, self.parent)
            return res
    _sd._QueryString = _CenteredQueryString

# ---------------------------------------------------------------
# Parsing / DB helper functions
# ---------------------------------------------------------------
def parse_expiry(date_str: str) -> str:
    if not date_str:
        return None
    ds = str(date_str).strip()
    if ds.lower() == "none":
        return None
    ds = re.sub(r'[-.\s]+', '/', ds)
    patterns = [
        r'^(\d{1,2})/(\d{1,2})/(\d{4})$',
        r'^(\d{4})/(\d{1,2})/(\d{1,2})$',
        r'^(\d{1,2})/(\d{4})$',
        r'^(\d{4})/(\d{1,2})$'
    ]
    for idx, pat in enumerate(patterns):
        m = re.match(pat, ds)
        if not m:
            continue
        g = m.groups()
        try:
            if idx == 0:
                d, mth, y = map(int, g)
                return datetime(y, mth, d).strftime("%Y-%m-%d")
            elif idx == 1:
                y, mth, d = map(int, g)
                return datetime(y, mth, d).strftime("%Y-%m-%d")
            elif idx == 2:
                mth, y = map(int, g)
                _, last = monthrange(y, mth)
                return datetime(y, mth, last).strftime("%Y-%m-%d")
            elif idx == 3:
                y, mth = map(int, g)
                _, last = monthrange(y, mth)
                return datetime(y, mth, last).strftime("%Y-%m-%d")
        except ValueError as e:
            logging.error(f"[parse_expiry] invalid date: {ds} -> {e}")
            return None
    return None

def check_expiry_required(code: str) -> bool:
    if not code:
        return False
    conn = connect_db()
    if conn is None:
        return False
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA table_info(items_management)")
        cols = {r[1].lower() for r in cur.fetchall()}
        if "remarks" not in cols or "code" not in cols:
            return False
        cur.execute("SELECT remarks FROM items_management WHERE code=?",(code,))
        row = cur.fetchone()
        if not row or not row["remarks"]:
            return False
        return "exp" in row["remarks"].lower()
    except sqlite3.Error as e:
        logging.error(f"[check_expiry_required] {e}")
        return False
    finally:
        cur.close()
        conn.close()

def fetch_project_details():
    conn = connect_db()
    if conn is None:
        return ("Unknown Project","PRJ")
    cur = conn.cursor()
    try:
        cur.execute("SELECT project_name, project_code FROM project_details LIMIT 1")
        r = cur.fetchone()
        return (r[0] if r and r[0] else "Unknown Project",
                r[1] if r and r[1] else "PRJ")
    except sqlite3.Error as e:
        logging.error(f"[fetch_project_details] {e}")
        return ("Unknown Project","PRJ")
    finally:
        cur.close()
        conn.close()

def fetch_third_parties():
    conn = connect_db()
    if conn is None: return []
    cur = conn.cursor()
    try:
        cur.execute("SELECT name FROM third_parties ORDER BY name")
        return [r[0] for r in cur.fetchall()]
    except sqlite3.Error as e:
        logging.error(f"[fetch_third_parties] {e}")
        return []
    finally:
        cur.close()
        conn.close()

def fetch_end_users():
    conn = connect_db()
    if conn is None: return []
    cur = conn.cursor()
    try:
        cur.execute("SELECT name FROM end_users ORDER BY name")
        return [r[0] for r in cur.fetchall()]
    except sqlite3.Error as e:
        logging.error(f"[fetch_end_users] {e}")
        return []
    finally:
        cur.close()
        conn.close()

# ---------------------------------------------------------------
# stock_data manipulator
# ---------------------------------------------------------------
class StockData:
    @staticmethod
    def add_or_update(unique_id, scenario, qty_in=0, qty_out=0,
                      exp_date=None, kit_number=None, module_number=None):
        conn = connect_db()
        if conn is None:
            raise ValueError("DB connection failed")
        cur = conn.cursor()
        try:
            cur.execute("PRAGMA table_info(stock_data)")
            cols = {row[1].lower() for row in cur.fetchall()}
            has_scenario = "scenario" in cols
            cur.execute("SELECT qty_in, qty_out FROM stock_data WHERE unique_id=?",(unique_id,))
            row = cur.fetchone()
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if row:
                new_in = (row[0] or 0) + qty_in
                new_out = (row[1] or 0) + qty_out
                if has_scenario:
                    cur.execute("""
                        UPDATE stock_data
                           SET qty_in=?, qty_out=?, exp_date=?, kit_number=?, module_number=?, scenario=?, updated_at=?
                         WHERE unique_id=?""",
                        (new_in, new_out, exp_date, kit_number, module_number, scenario, now, unique_id))
                else:
                    cur.execute("""
                        UPDATE stock_data
                           SET qty_in=?, qty_out=?, exp_date=?, kit_number=?, module_number=?, updated_at=?
                         WHERE unique_id=?""",
                        (new_in, new_out, exp_date, kit_number, module_number, now, unique_id))
            else:
                if has_scenario:
                    cur.execute("""
                        INSERT INTO stock_data
                            (unique_id, qty_in, qty_out, exp_date, kit_number, module_number, scenario, updated_at)
                        VALUES (?,?,?,?,?,?,?,?)""",
                        (unique_id, qty_in, qty_out, exp_date, kit_number, module_number, scenario, now))
                else:
                    cur.execute("""
                        INSERT INTO stock_data
                            (unique_id, qty_in, qty_out, exp_date, kit_number, module_number, updated_at)
                        VALUES (?,?,?,?,?,?,?)""",
                        (unique_id, qty_in, qty_out, exp_date, kit_number, module_number, now))
            conn.commit()
        except sqlite3.Error as e:
            conn.rollback()
            logging.error(f"[StockData.add_or_update] {e}")
            raise
        finally:
            cur.close()
            conn.close()

    @staticmethod
    def consume_by_line_id(line_id: int, qty_out_add: int):
        """
        Increase qty_out for existing stock_data line identified by line_id.
        """
        if qty_out_add <= 0:
            return
        conn = connect_db()
        if conn is None:
            raise ValueError("DB connection failed")
        cur = conn.cursor()
        try:
            cur.execute("SELECT qty_in, qty_out FROM stock_data WHERE line_id=?", (line_id,))
            row = cur.fetchone()
            if not row:
                logging.warning(f"[consume_by_line_id] line_id {line_id} not found")
                return
            new_out = (row[1] or 0) + qty_out_add
            cur.execute("""
                UPDATE stock_data
                   SET qty_out = ?, updated_at = CURRENT_TIMESTAMP
                 WHERE line_id = ?
            """, (new_out, line_id))
            conn.commit()
        except sqlite3.Error as e:
            conn.rollback()
            logging.error(f"[consume_by_line_id] {e}")
        finally:
            cur.close()
            conn.close()

# ---------------------------------------------------------------
# Main UI class
# ---------------------------------------------------------------
class StockInKit(tk.Frame):
    FIXED_IN_TYPE = "Internal move from on-shelf items"

    def __init__(self, parent, app, role: str = "supervisor"):
        super().__init__(parent)
        self.parent = parent
        self.app = app
        self.role = role.lower()

        def _sv(default=""):
            try:
                return tk.StringVar(value=default)
            except Exception:
                class DummyVar:
                    def __init__(self, v): self._v = v
                    def set(self, v): self._v = v
                    def get(self): return self._v
                return DummyVar(default)

        self.trans_type_var    = _sv(self.FIXED_IN_TYPE)
        self.scenario_var      = _sv("")
        self.mode_var          = _sv("")
        self.kit_var           = _sv("")
        self.kit_number_var    = _sv("")
        self.module_var        = _sv("")
        self.module_number_var = _sv("")
        self.end_user_var      = _sv("")
        self.third_party_var   = _sv("")
        self.status_var        = _sv(lang.t("receive_kit.ready","Ready"))

        self.tree = None
        self.search_var = None
        self.search_listbox = None
        self.scenario_cb = None
        self.mode_cb = None
        self.kit_cb = None
        self.kit_number_cb = None
        self.module_cb = None
        self.module_number_cb = None
        self.end_user_cb = None
        self.third_party_cb = None
        self.remarks_entry = None

        self.row_data = {}
        self.code_to_iid = {}
        self.mode_definitions = []
        self.mode_label_to_key = {}
        self.selected_scenario_id = None
        self.selected_scenario_name = None
        self.suggested_usage = defaultdict(int)
        self.editing_cell = None

        self.scenario_map = self.fetch_scenario_map()
        self.reverse_scenario_map = {v: k for k, v in self.scenario_map.items()}

        logging.info("StockInKit initialized")

        if self.parent is not None and self.parent.winfo_exists():
            self.pack(fill="both", expand=True)
            self.after(50, self.initialize_ui)
        else:
            logging.error("Parent window missing at initialization")

    # ------------- Ensure Vars ----------------
    def ensure_vars_ready(self):
        needed = {
            "trans_type_var": self.FIXED_IN_TYPE,
            "scenario_var": "",
            "mode_var": "",
            "kit_var": "",
            "kit_number_var": "",
            "module_var": "",
            "module_number_var": "",
            "end_user_var": "",
            "third_party_var": "",
            "status_var": lang.t("receive_kit.ready","Ready")
        }
        for name, default in needed.items():
            cur = getattr(self, name, None)
            if cur is None or not hasattr(cur, "set"):
                try:
                    setattr(self, name, tk.StringVar(value=default))
                except Exception:
                    class DummyVar:
                        def __init__(self,v): self._v=v
                        def set(self,v): self._v=v
                        def get(self): return self._v
                    setattr(self, name, DummyVar(default))

    # ------------- Scenario helpers ------------
    def fetch_scenario_map(self):
        conn = connect_db()
        if conn is None:
            return {}
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("SELECT scenario_id, name FROM scenarios ORDER BY name")
            return {str(r['scenario_id']): r['name'] for r in cur.fetchall()}
        except sqlite3.Error as e:
            logging.error(f"[fetch_scenario_map] {e}")
            return {}
        finally:
            cur.close()
            conn.close()

    def fetch_scenarios(self):
        conn = connect_db()
        if conn is None: return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("SELECT scenario_id, name FROM scenarios ORDER BY name")
            return [{"id": r['scenario_id'], "name": r['name']} for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[fetch_scenarios] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    # ------------- Modes -----------------------
    def build_mode_definitions(self):
        scenario = self.selected_scenario_name or ""
        self.mode_definitions = [
            ("receive_kit",        lang.t("receive_kit.mode_generate_kit", "Generate Kit")),
            ("add_standalone",     lang.t("receive_kit.mode_add_standalone",     "Add standalone item/s in {scenario}", scenario=scenario)),
            ("add_module_scenario",lang.t("receive_kit.mode_add_module_scenario","Add module to {scenario}", scenario=scenario)),
            ("add_module_kit",     lang.t("receive_kit.mode_add_module_kit",     "Add module to a kit")),
            ("add_items_kit",      lang.t("receive_kit.mode_add_items_kit",      "Add items to a kit")),
            ("add_items_module",   lang.t("receive_kit.mode_add_items_module",   "Add items to a module"))
        ]
        self.mode_label_to_key = {lbl: key for key, lbl in self.mode_definitions}

    def current_mode_key(self):
        label = self.mode_var.get()
        return self.mode_label_to_key.get(label)

    def ensure_mode_ready(self):
        if not self.mode_definitions:
            self.build_mode_definitions()
        mk = self.current_mode_key()
        if not mk and self.mode_definitions:
            self.mode_var.set(self.mode_definitions[0][1])
            mk = self.mode_definitions[0][0]
        return mk

    # ------------- UI Init --------------------
    def initialize_ui(self):
        if not (self.parent and self.parent.winfo_exists()):
            return
        try:
            self.render_ui()
        except tk.TclError as e:
            logging.error(f"[initialize_ui] {e}")
            custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                         f"Failed to render UI: {e}", "error")

    def render_ui(self):
        self.ensure_vars_ready()
        for w in list(self.parent.winfo_children()):
            try: w.destroy()
            except: pass

        title_lbl = tk.Label(self.parent,
                             text=lang.t("receive_kit.title", "Receive Kit-Module"),
                             font=("Helvetica", 20, "bold"),
                             bg="#F0F4F8")
        title_lbl.pack(pady=10, fill="x")

        main = tk.Frame(self.parent, bg="#F0F4F8")
        main.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(main, text=lang.t("receive_kit.scenario","Scenario:"), bg="#F0F4F8")\
            .grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.scenario_cb = ttk.Combobox(main, textvariable=self.scenario_var,
                                        state="readonly", width=40)
        self.scenario_cb.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        self.scenario_cb.bind("<<ComboboxSelected>>", self.on_scenario_selected)

        tk.Label(main, text=lang.t("receive_kit.movement_type","Movement Type:"), bg="#F0F4F8")\
            .grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.mode_cb = ttk.Combobox(main, textvariable=self.mode_var,
                                    state="readonly", width=40)
        self.mode_cb.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        self.mode_cb.bind("<<ComboboxSelected>>", self.update_mode)

        tk.Label(main, text=lang.t("receive_kit.select_kit","Select Kit:"), bg="#F0F4F8")\
            .grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.kit_cb = ttk.Combobox(main, textvariable=self.kit_var, state="disabled", width=40)
        self.kit_cb.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.kit_cb.bind("<<ComboboxSelected>>", self.on_kit_selected)

        tk.Label(main, text=lang.t("receive_kit.select_kit_number","Select Kit Number:"), bg="#F0F4F8")\
            .grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.kit_number_cb = ttk.Combobox(main, textvariable=self.kit_number_var, state="disabled", width=20)
        self.kit_number_cb.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.kit_number_cb.bind("<<ComboboxSelected>>", self.on_kit_number_selected)

        tk.Label(main, text=lang.t("receive_kit.select_module","Select Module:"), bg="#F0F4F8")\
            .grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.module_cb = ttk.Combobox(main, textvariable=self.module_var, state="disabled", width=40)
        self.module_cb.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.module_cb.bind("<<ComboboxSelected>>", self.on_module_selected)

        tk.Label(main, text=lang.t("receive_kit.select_module_number","Select Module Number:"), bg="#F0F4F8")\
            .grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.module_number_cb = ttk.Combobox(main, textvariable=self.module_number_var, state="disabled", width=20)
        self.module_number_cb.grid(row=3, column=3, padx=5, pady=5, sticky="w")
        self.module_number_cb.bind("<<ComboboxSelected>>", self.on_module_number_selected)

        type_frame = tk.Frame(main, bg="#F0F4F8")
        type_frame.grid(row=4, column=0, columnspan=4, pady=5, sticky="w")
        tk.Label(type_frame, text=lang.t("receive_kit.in_type","IN Type:"), bg="#F0F4F8")\
            .grid(row=0, column=0, padx=5, sticky="w")
        tk.Label(type_frame, textvariable=self.trans_type_var,
                 bg="#E0E0E0", fg="#000", relief="sunken", width=30, anchor="w")\
            .grid(row=0, column=1, padx=5, pady=5, sticky="w")

        tk.Label(type_frame, text=lang.t("receive_kit.end_user","End User:"), bg="#F0F4F8")\
            .grid(row=0, column=2, padx=5, sticky="w")
        self.end_user_cb = ttk.Combobox(type_frame, textvariable=self.end_user_var,
                                        state="disabled", width=30)
        self.end_user_cb['values'] = fetch_end_users()
        self.end_user_cb.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(type_frame, text=lang.t("receive_kit.third_party","Third Party:"), bg="#F0F4F8")\
            .grid(row=0, column=4, padx=5, sticky="w")
        self.third_party_cb = ttk.Combobox(type_frame, textvariable=self.third_party_var,
                                           state="disabled", width=30)
        self.third_party_cb['values'] = fetch_third_parties()
        self.third_party_cb.grid(row=0, column=5, padx=5, pady=5)

        tk.Label(type_frame, text=lang.t("receive_kit.remarks","Remarks:"), bg="#F0F4F8")\
            .grid(row=0, column=6, padx=5, sticky="w")
        self.remarks_entry = tk.Entry(type_frame, width=40, state="disabled")
        self.remarks_entry.grid(row=0, column=7, padx=5, pady=5)

        tk.Label(main, text=lang.t("receive_kit.search_item","Search Kit/Module/Item:"), bg="#F0F4F8")\
            .grid(row=5, column=0, padx=5, pady=(10,0), sticky="w")
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(main, textvariable=self.search_var, width=50)
        search_entry.grid(row=5, column=1, columnspan=3, padx=5, pady=(10,0), sticky="w")
        search_entry.bind("<KeyRelease>", self.search_items)
        search_entry.bind("<Return>", self.select_first_result)

        self.search_listbox = tk.Listbox(main, height=5, width=60)
        self.search_listbox.grid(row=6, column=0, columnspan=4, padx=5, pady=5, sticky="we")
        self.search_listbox.bind("<<ListboxSelect>>", self.fill_from_search)

        # Columns + hidden line_id & qty_out (last two)
        cols = ("code","description","type","kit","module",
                "std_qty","qty_to_receive","expiry_date","batch_no","line_id","qty_out")
        self.tree = ttk.Treeview(main, columns=cols, show="headings", height=22)
        self.tree.tag_configure("light_red", background="#FF9999")
        self.tree.tag_configure("kit", background="#228B22", foreground="white")
        self.tree.tag_configure("module", background="#ADD8E6")

        headers = {
            "code": lang.t("receive_kit.code","Code"),
            "description": lang.t("receive_kit.description","Description"),
            "type": lang.t("receive_kit.type","Type"),
            "kit": lang.t("receive_kit.kit","Kit"),
            "module": lang.t("receive_kit.module","Module"),
            "std_qty": lang.t("receive_kit.std_qty","Std Qty"),
            "qty_to_receive": lang.t("receive_kit.qty_to_receive","Qty to Receive"),
            "expiry_date": lang.t("receive_kit.expiry_date","Expiry Date"),
            "batch_no": lang.t("receive_kit.batch_no","Batch No"),
            "line_id": "line_id (hidden)",
            "qty_out": "qty_out (hidden)"
        }
        widths = {
            "code":130,"description":380,"type":110,"kit":120,
            "module":120,"std_qty":90,"qty_to_receive":120,
            "expiry_date":120,"batch_no":120,"line_id":1,"qty_out":1
        }
        aligns = {
            "code":"w","description":"w","type":"w","kit":"w","module":"w",
            "std_qty":"e","qty_to_receive":"e","expiry_date":"w",
            "batch_no":"w","line_id":"w","qty_out":"e"
        }
        for c in cols:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=widths[c], anchor=aligns[c], stretch=False if c in ("line_id","qty_out") else True)
        vsb = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(main, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=7, column=0, columnspan=4, pady=10, sticky="nsew")
        vsb.grid(row=7, column=4, sticky="ns")
        hsb.grid(row=8, column=0, columnspan=4, sticky="ew")
        main.grid_rowconfigure(7, weight=1)
        main.grid_columnconfigure(1, weight=1)

        for ev in ("<Double-1>", "<KeyPress-Return>", "<KeyPress-Tab>"):
            self.tree.bind(ev, self.start_edit)
        for ev in ("<KeyPress-Up>", "<KeyPress-Down>", "<KeyPress-Left>", "<KeyPress-Right>"):
            self.tree.bind(ev, self.navigate_tree)
        self.tree.bind("<Button-3>", self.show_context_menu)

        btn_frame = tk.Frame(main, bg="#F0F4F8")
        btn_frame.grid(row=9, column=0, columnspan=4, pady=5)
        tk.Button(btn_frame, text=lang.t("receive_kit.add_missing","Add Missing Item"),
                  bg="#FFA500", fg="white", command=self.add_missing_item).pack(side="left", padx=5)
        tk.Button(btn_frame, text=lang.t("receive_kit.save","Save"),
                  bg="#27AE60", fg="white",
                  state="normal" if self.role in ["admin","manager"] else "disabled",
                  command=self.save_all).pack(side="left", padx=5)
        tk.Button(btn_frame, text=lang.t("receive_kit.clear","Clear"),
                  bg="#7F8C8D", fg="white", command=self.clear_form).pack(side="left", padx=5)
        tk.Button(btn_frame, text=lang.t("receive_kit.export","Export"),
                  bg="#2980B9", fg="white", command=self.export_data).pack(side="left", padx=5)

        tk.Label(main, textvariable=self.status_var, relief="sunken",
                 anchor="w", bg="#F0F4F8").grid(row=10, column=0, columnspan=4, sticky="ew")

        self.load_scenarios()

    # ------------- Scenario & mode logic -------------
    def load_scenarios(self):
        self.ensure_vars_ready()
        scenarios = self.fetch_scenarios()
        names = [f"{s['id']} - {s['name']}" for s in scenarios]
        self.scenario_cb['values'] = names
        if names:
            self.scenario_cb.current(0)
            self.on_scenario_selected()

    def on_scenario_selected(self, event=None):
        scen = self.scenario_var.get()
        if not scen:
            self.selected_scenario_id = None
            self.selected_scenario_name = None
            self.clear_search()
            return
        self.selected_scenario_id, self.selected_scenario_name = scen.split(" - ", 1)
        self.build_mode_definitions()
        self.mode_cb['values'] = [lbl for _, lbl in self.mode_definitions]
        if self.mode_definitions:
            self.mode_var.set(self.mode_definitions[0][1])
        self.update_mode()
        self.clear_search()

    def update_mode(self, event=None):
        if self.tree:
            self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.code_to_iid.clear()
        self.suggested_usage.clear()
        if self.search_listbox:
            self.search_listbox.delete(0, tk.END)
        if self.search_var:
            self.search_var.set("")
        mk = self.current_mode_key()
        for cb in (self.kit_cb, self.kit_number_cb, self.module_cb, self.module_number_cb):
            cb.config(state="disabled")
        if mk in ["add_module_kit","add_items_kit","add_items_module"]:
            self.kit_cb.config(state="readonly")
            self.kit_cb['values'] = self.fetch_kits(self.selected_scenario_id)
            self.kit_number_cb.config(state="readonly")
            self.kit_number_cb['values'] = self.fetch_available_kit_numbers(self.selected_scenario_id)
            if mk == "add_items_module":
                self.module_cb.config(state="readonly")
                self.module_cb['values'] = self.fetch_all_modules(self.selected_scenario_id)
                self.module_number_cb.config(state="readonly")
                self.module_number_cb['values'] = self.fetch_module_numbers(self.selected_scenario_id)
        results = self.fetch_search_results("", self.selected_scenario_id, mk)
        for r in results:
            self.search_listbox.insert(tk.END, f"{r['code']} - {r['description']}")
        self.status_var.set(lang.t("receive_kit.found_items", f"Found {self.search_listbox.size()} items"))

    # ------------- Kit / Module selection handlers -------------
    def on_kit_selected(self, event=None):
        kit_code = self.kit_var.get() or ""
        self.kit_number_var.set("")
        self.module_var.set("")
        self.module_number_var.set("")
        mk = self.current_mode_key()
        if mk in ["add_module_kit", "add_items_kit", "add_items_module"]:
            self.kit_number_cb.config(state="readonly")
            self.kit_number_cb['values'] = self.fetch_available_kit_numbers(self.selected_scenario_id, kit_code or None)
        else:
            self.kit_number_cb.config(state="disabled")
        if mk == "add_items_module":
            if kit_code:
                self.module_cb.config(state="readonly")
                self.module_cb['values'] = self.fetch_modules_for_kit(self.selected_scenario_id, kit_code)
            else:
                self.module_cb.config(state="disabled")
                self.module_cb['values'] = []
            self.module_number_cb.config(state="disabled")
            self.module_number_cb['values'] = []
            self.module_number_var.set("")
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.code_to_iid.clear()
        self.suggested_usage.clear()
        self.search_listbox.delete(0, tk.END)
        res = self.fetch_search_results("", self.selected_scenario_id, mk)
        for r in res:
            self.search_listbox.insert(tk.END, f"{r['code']} - {r['description']}")
        self.status_var.set(lang.t("receive_kit.found_items", f"Found {self.search_listbox.size()} items"))

    def on_kit_number_selected(self, event=None):
        mk = self.current_mode_key()
        if mk not in ["add_module_kit", "add_items_kit", "add_items_module"]:
            return
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.code_to_iid.clear()
        self.suggested_usage.clear()
        self.status_var.set(lang.t("receive_kit.ready", "Ready"))
        self.search_listbox.delete(0, tk.END)
        res = self.fetch_search_results("", self.selected_scenario_id, mk)
        for r in res:
            self.search_listbox.insert(tk.END, f"{r['code']} - {r['description']}")
        self.status_var.set(lang.t("receive_kit.found_items", f"Found {self.search_listbox.size()} items"))

    def on_module_selected(self, event=None):
        mk = self.current_mode_key()
        if mk != "add_items_module":
            return
        kit_code = self.kit_var.get()
        module_code = self.module_var.get()
        self.module_number_var.set("")
        self.module_number_cb['values'] = self.fetch_module_numbers(
            self.selected_scenario_id,
            kit_code=kit_code or None,
            module_code=module_code or None
        )
        self.module_number_cb.config(
            state="readonly" if self.module_number_cb['values'] else "disabled"
        )
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.code_to_iid.clear()
        self.suggested_usage.clear()
        self.status_var.set(lang.t("receive_kit.ready","Ready"))

    def on_module_number_selected(self, event=None):
        mk = self.current_mode_key()
        if mk != "add_items_module":
            return
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.code_to_iid.clear()
        self.suggested_usage.clear()
        self.status_var.set(lang.t("receive_kit.ready","Ready"))

    # ------------- Clear & Search -------------
    def clear_search(self):
        if self.search_var:
            self.search_var.set("")
        if self.search_listbox:
            self.search_listbox.delete(0, tk.END)
        if self.tree:
            self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.code_to_iid.clear()
        self.suggested_usage.clear()
        self.status_var.set(lang.t("receive_kit.ready","Ready"))

    def clear_form(self):
        self.ensure_vars_ready()
        self.clear_search()
        self.trans_type_var.set(self.FIXED_IN_TYPE)
        self.scenario_var.set("")
        self.mode_var.set("")
        self.kit_var.set("")
        self.kit_number_var.set("")
        self.module_var.set("")
        self.module_number_var.set("")
        self.end_user_var.set("")
        self.third_party_var.set("")
        if self.remarks_entry:
            self.remarks_entry.config(state="normal")
            self.remarks_entry.delete(0, tk.END)
            self.remarks_entry.config(state="disabled")
        if self.end_user_cb:
            self.end_user_cb.config(state="disabled")
        if self.third_party_cb:
            self.third_party_cb.config(state="disabled")
        self.load_scenarios()

    def search_items(self, event=None):
        if not self.selected_scenario_id:
            return
        q = self.search_var.get().strip() if self.search_var else ""
        self.search_listbox.delete(0, tk.END)
        mk = self.current_mode_key()
        results = self.fetch_search_results(q, self.selected_scenario_id, mk)
        for r in results:
            self.search_listbox.insert(tk.END, f"{r['code']} - {r['description']}")
        self.status_var.set(lang.t("receive_kit.found_items", f"Found {self.search_listbox.size()} items"))

    def select_first_result(self, event=None):
        if self.search_listbox.size() > 0:
            self.search_listbox.selection_set(0)
            self.fill_from_search()

    def fill_from_search(self, event=None):
        sel = self.search_listbox.curselection()
        if not sel:
            return
        line = self.search_listbox.get(sel[0])
        code = line.split(" - ")[0]
        self.search_listbox.selection_clear(0, tk.END)
        mk = self.current_mode_key()
        if mk == "receive_kit":
            self.load_hierarchy(code)
        else:
            self.add_to_tree(code)

    # ------------- Fetch domain data -------------
    def fetch_kits(self, scenario_id):
        if not scenario_id:
            return []
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT code
                  FROM kit_items
                 WHERE scenario_id=? AND level='primary'
                 ORDER BY code
            """,(scenario_id,))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[fetch_kits] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_all_modules(self, scenario_id):
        if not scenario_id: return []
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT code
                  FROM kit_items
                 WHERE scenario_id=? AND level='secondary'
                 ORDER BY code
            """,(scenario_id,))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[fetch_all_modules] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_modules_for_kit(self, scenario_id, kit_code):
        if not scenario_id or not kit_code: return []
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT code
                  FROM kit_items
                 WHERE scenario_id=? AND kit=? AND level='secondary'
                 ORDER BY code
            """,(scenario_id, kit_code))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[fetch_modules_for_kit] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_available_kit_numbers(self, scenario_id, kit_code=None):
        if not scenario_id: return []
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            if kit_code:
                cur.execute("""
                    SELECT DISTINCT kit_number
                      FROM stock_data
                     WHERE kit_number IS NOT NULL
                       AND kit_number!='None'
                       AND unique_id LIKE ?
                     ORDER BY kit_number
                """,(f"{scenario_id}/{kit_code}/%",))
            else:
                cur.execute("""
                    SELECT DISTINCT kit_number
                      FROM stock_data
                     WHERE kit_number IS NOT NULL
                       AND kit_number!='None'
                       AND unique_id LIKE ?
                     ORDER BY kit_number
                """,(f"{scenario_id}/%",))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[fetch_available_kit_numbers] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_module_numbers(self, scenario_id, kit_code=None, module_code=None):
        if not scenario_id: return []
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            where = ["module_number IS NOT NULL","module_number!='None'","unique_id LIKE ?"]
            params = [f"{scenario_id}/%"]
            if kit_code:
                where.append("unique_id LIKE ?")
                params.append(f"{scenario_id}/{kit_code}/%")
            if module_code:
                where.append("unique_id LIKE ?")
                params.append(f"{scenario_id}/%/{module_code}/%")
            sql = f"""
                SELECT DISTINCT module_number
                  FROM stock_data
                 WHERE {' AND '.join(where)}
                 ORDER BY module_number
            """
            cur.execute(sql, params)
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[fetch_module_numbers] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_search_results(self, query, scenario_id, mode_key):
        if not scenario_id:
            return []
        if not mode_key:
            mode_key = self.ensure_mode_ready()
        if mode_key in (self.mode_label_to_key or {}):
            mode_key = self.mode_label_to_key[mode_key]
        mk = (mode_key or "").lower()
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            q = (query or "").lower()
            def common_params():
                return (f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%")
            if mk == "receive_kit":
                sql = """
                   SELECT DISTINCT ki.code, ki.level
                     FROM kit_items ki
                LEFT JOIN items_list il ON il.code=ki.code
                    WHERE ki.scenario_id=?
                      AND LOWER(ki.level)='primary'
                      AND (
                        UPPER(ki.code) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_en,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_fr,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_sp,'')) LIKE UPPER(?)
                      )
                 ORDER BY ki.code
                """
                params = (scenario_id, *common_params())
            elif mk == "add_standalone":
                sql = """
                   SELECT DISTINCT ki.code, ki.level
                     FROM kit_items ki
                LEFT JOIN items_list il ON il.code=ki.code
                    WHERE ki.scenario_id=?
                      AND LOWER(ki.level)='tertiary'
                      AND (
                        UPPER(ki.code) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_en,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_fr,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_sp,'')) LIKE UPPER(?)
                      )
                 ORDER BY ki.code
                """
                params = (scenario_id, *common_params())
            elif mk == "add_module_scenario":
                sql = """
                   SELECT DISTINCT ki.code, ki.level
                     FROM kit_items ki
                LEFT JOIN items_list il ON il.code=ki.code
                    WHERE ki.scenario_id=?
                      AND LOWER(ki.level)='secondary'
                      AND (
                        UPPER(ki.code) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_en,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_fr,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_sp,'')) LIKE UPPER(?)
                      )
                 ORDER BY ki.code
                """
                params = (scenario_id, *common_params())
            elif mk == "add_module_kit":
                kit_code = self.kit_var.get()
                sql = """
                   SELECT DISTINCT ki.code, ki.level
                     FROM kit_items ki
                LEFT JOIN items_list il ON il.code=ki.code
                    WHERE ki.scenario_id=?
                      AND ki.kit=?
                      AND LOWER(ki.level)='secondary'
                      AND (
                        UPPER(ki.code) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_en,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_fr,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_sp,'')) LIKE UPPER(?)
                      )
                 ORDER BY ki.code
                """
                params = (scenario_id, kit_code, *common_params())
            elif mk == "add_items_kit":
                kit_code = self.kit_var.get()
                sql = """
                   SELECT DISTINCT ki.code, ki.level
                     FROM kit_items ki
                LEFT JOIN items_list il ON il.code=ki.code
                    WHERE ki.scenario_id=?
                      AND ki.kit=?
                      AND LOWER(ki.level)='tertiary'
                      AND (
                        UPPER(ki.code) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_en,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_fr,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_sp,'')) LIKE UPPER(?)
                      )
                 ORDER BY ki.code
                """
                params = (scenario_id, kit_code, *common_params())
            elif mk == "add_items_module":
                kit_code = self.kit_var.get()
                module_code = self.module_var.get()
                sql = """
                   SELECT DISTINCT ki.code, ki.level
                     FROM kit_items ki
                LEFT JOIN items_list il ON il.code=ki.code
                    WHERE ki.scenario_id=?
                      AND ki.kit=?
                      AND ki.module=?
                      AND LOWER(ki.level)='tertiary'
                      AND (
                        UPPER(ki.code) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_en,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_fr,'')) LIKE UPPER(?)
                        OR UPPER(COALESCE(il.designation_sp,'')) LIKE UPPER(?)
                      )
                 ORDER BY ki.code
                """
                params = (scenario_id, kit_code, module_code, *common_params())
            else:
                return []
            cur.execute(sql, params)
            rows = cur.fetchall()
            out = []
            for r in rows:
                code = r['code']
                desc = get_item_description(code)
                t = detect_type(code, desc)
                out.append({
                    'code': code,
                    'description': desc,
                    'level': r['level'],
                    'type': t
                })
            return out
        except sqlite3.Error as e:
            logging.error(f"[fetch_search_results] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_kit_items(self, scenario_id, code):
        if not scenario_id or not code:
            return []
        conn = connect_db()
        if conn is None: return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT code, level, kit, module, item,
                       COALESCE(std_qty,0) AS std_qty,
                       treecode
                  FROM kit_items
                 WHERE scenario_id=?
                   AND (code=? OR kit=? OR module=? OR item=?)
                 ORDER BY treecode
            """,(scenario_id, code, code, code, code))
            comps = []
            for r in cur.fetchall():
                dsc = get_item_description(r['code'])
                comps.append({
                    'code': r['code'],
                    'description': dsc,
                    'type': detect_type(r['code'], dsc),
                    'kit': r['kit'] or "-----",
                    'module': r['module'] or "-----",
                    'item': r['item'] or "-----",
                    'std_qty': r['std_qty'],
                    'treecode': r['treecode']
                })
            return comps
        except sqlite3.Error as e:
            logging.error(f"[fetch_kit_items] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    # ------------- On-shelf batch extraction (per line) -------------
    def fetch_on_shelf_batches(self, code: str):
        """
        Return individual stock_data lines (not aggregated) for the given item code
        under the selected scenario (management_mode on-shelf if present).
        Each entry: { code, expiry, management_mode, final_qty, line_id }
        final_qty = qty_in - qty_out (computed).
        """
        if not code:
            return []
        requested = code.strip().upper()
        scenario_name = (self.selected_scenario_name or "").strip()
        like_param = f"{scenario_name}/%" if scenario_name else "%"
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        batches = []
        try:
            cur.execute("PRAGMA table_info(stock_data)")
            cols = {c[1].lower(): c[1] for c in cur.fetchall()}
            has_mgmt = 'management_mode' in cols
            has_qin = 'qty_in' in cols
            has_qout = 'qty_out' in cols

            # Pull rows for scenario prefix
            cur.execute("SELECT * FROM stock_data WHERE unique_id LIKE ?", (like_param,))
            for r in cur.fetchall():
                u = r['unique_id']
                parts = u.split('/')
                if len(parts) < 6:
                    continue
                scen_part, kit_part, module_part, item_part, std_part, exp_part = parts[:6]
                if scenario_name and scen_part.lower() != scenario_name.lower():
                    continue
                item_canon = item_part.strip().upper()
                if item_canon != requested:
                    continue
                if has_mgmt:
                    mm = (r['management_mode'] or "")
                    norm = re.sub(r'[\s_-]+','', mm.lower())
                    if norm != 'onshelf':
                        continue
                qin = r['qty_in'] if has_qin and r['qty_in'] else 0
                qout = r['qty_out'] if has_qout and r['qty_out'] else 0
                final_qty = (qin or 0) - (qout or 0)
                if final_qty <= 0:
                    continue
                exp_raw = exp_part if exp_part != "None" else ""
                parsed = parse_expiry(exp_raw) if exp_raw else ""
                expiry_norm = parsed if parsed else exp_raw
                batches.append({
                    'code': code,
                    'expiry': expiry_norm,
                    'management_mode': 'on_shelf',
                    'final_qty': final_qty,
                    'line_id': r['line_id']
                })
            logging.info(f"[fetch_on_shelf_batches] code={code} lines={len(batches)}")
            return batches
        except sqlite3.Error as e:
            logging.error(f"[fetch_on_shelf_batches] {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def remaining_available(self, code, expiry, management_mode, total_final):
        key = (code, expiry or "", management_mode)
        used = self.suggested_usage.get(key, 0)
        return max(total_final - used, 0)

    def record_suggested(self, code, expiry, management_mode, qty):
        key = (code, expiry or "", management_mode)
        self.suggested_usage[key] += qty

    def _remaining_std_allowance(self, code: str) -> int:
        std_val = None
        assigned = 0
        for top in self.tree.get_children():
            stack = [top]
            while stack:
                iid = stack.pop()
                vals = self.tree.item(iid, "values")
                if not vals: continue
                c, t, s, q = vals[0], vals[2], vals[5], vals[6]
                if c == code and t.upper() == "ITEM":
                    if std_val is None and str(s).isdigit():
                        std_val = int(s)
                    if q and str(q).isdigit():
                        assigned += int(q)
                stack.extend(self.tree.get_children(iid))
        if not std_val or std_val <= 0:
            return 10**12
        rem = std_val - assigned
        return rem if rem > 0 else 0

    # ------------- Module number helper -------------
    def ask_module_number(self, kit_number: str, module_code: str) -> str:
        base = kit_number if kit_number else module_code
        suggestion = f"{base}-M" if base else "M1"
        while True:
            entered = simpledialog.askstring(
                lang.t("receive_kit.module_number","Module Number"),
                f"Enter Module Number (suggested: {suggestion})",
                parent=self.parent
            )
            if entered is None:
                return None
            entered = entered.strip()
            if not entered:
                continue
            if self.is_module_number_unique(kit_number, entered):
                return entered
            custom_popup(self.parent,
                         lang.t("dialog_titles.error","Error"),
                         lang.t("receive_kit.duplicate_module_number",
                                f"Module Number '{entered}' already exists."), "error")

    # ------------- Hierarchy load -------------
    def load_hierarchy(self, kit_code):
        if not self.selected_scenario_id or not kit_code:
            return
        self.tree.delete(*self.tree.get_children())
        self.row_data.clear()
        self.code_to_iid.clear()
        self.suggested_usage.clear()
        comps = self.fetch_kit_items(self.selected_scenario_id, kit_code)
        if not comps:
            self.status_var.set(lang.t("receive_kit.no_items", f"No items for {kit_code}"))
            return
        kit_number = self.kit_number_var.get().strip() if self.kit_number_var.get() else None
        if not kit_number:
            while True:
                kn = simpledialog.askstring(
                    lang.t("receive_kit.kit_number","Kit Number"),
                    f"Enter Kit Number for {kit_code}",
                    parent=self.parent
                )
                if kn is None:
                    self.status_var.set(lang.t("receive_kit.module_number_cancelled","Cancelled"))
                    return
                kn = kn.strip()
                if not kn:
                    continue
                if self.is_kit_number_unique(kn):
                    kit_number = kn
                    self.kit_number_var.set(kn)
                    break
                custom_popup(self.parent,
                             lang.t("dialog_titles.error","Error"),
                             lang.t("receive_kit.duplicate_kit_number",
                                    f"Kit Number '{kn}' already exists globally."), "error")
        treecode_to_iid = {}
        module_number_map = {}
        for comp in sorted(comps, key=lambda x: x['treecode']):
            tc = comp['treecode']
            parent_tc = tc[:-3] if len(tc) >= 3 else ""
            parent_iid = treecode_to_iid.get(parent_tc, "")
            ctype = comp['type'].upper()
            std_qty = comp['std_qty']
            kit_disp = comp['kit'] or "-----"
            module_disp = comp['module'] or "-----"
            module_number = None
            if ctype == "MODULE":
                module_number = self.ask_module_number(kit_number, comp['code'])
                if module_number is None:
                    self.status_var.set(lang.t("receive_kit.module_number_cancelled","Cancelled"))
                    return
                module_number_map[comp['code']] = module_number
            else:
                if comp['module'] and comp['module'] in module_number_map:
                    module_number = module_number_map[comp['module']]
            row_iid = self.tree.insert(parent_iid, "end", values=(
                comp['code'], comp['description'], comp['type'],
                kit_disp, module_disp,
                std_qty, (std_qty if ctype!="ITEM" else 0), "", "",
                "", ""  # hidden line_id, qty_out
            ))
            treecode_to_iid[tc] = row_iid
            self.code_to_iid[comp['code']] = row_iid
            if ctype == "KIT":
                self.tree.item(row_iid, tags=("kit",))
            elif ctype == "MODULE":
                self.tree.item(row_iid, tags=("module",))
            self.row_data[row_iid] = {
                'kit_number': kit_number,
                'module_number': module_number,
                'treecode': tc
            }
            if ctype == "ITEM":
                self.tree.delete(row_iid)
                del self.row_data[row_iid]
                if comp['code'] in self.code_to_iid:
                    del self.code_to_iid[comp['code']]
                self.insert_item_batches(
                    code=comp['code'],
                    parent_iid=parent_iid,
                    kit_number=kit_number,
                    module_number=module_number,
                    structural_kit=kit_disp if kit_disp != "-----" else None,
                    structural_module=module_disp if module_disp != "-----" else None,
                    std_qty=std_qty
                )
        self.update_child_quantities()
        self.update_parent_expiry()
        self.status_var.set(lang.t("receive_kit.loaded_records", f"Loaded hierarchy for {kit_code}"))

    # ------------- Insert batch rows (with line_id & hidden qty_out) -------------
    def insert_item_batches(self, code, parent_iid, kit_number, module_number,
                            structural_kit, structural_module, std_qty):
        batches = self.fetch_on_shelf_batches(code)
        desc = get_item_description(code)
        def sort_key(b):
            raw = b['expiry']
            if not raw:
                return (2, "")
            parsed = parse_expiry(raw)
            if parsed:
                return (0, parsed)
            return (1, raw)
        batches_sorted = sorted(batches, key=sort_key, reverse=True)
        def remaining_std_allowance_for_insertion():
            total_assigned = 0
            current_std = std_qty if isinstance(std_qty, int) else (int(std_qty) if str(std_qty).isdigit() else 0)
            for top in self.tree.get_children():
                stack = [top]
                while stack:
                    iid = stack.pop()
                    vals = self.tree.item(iid, "values")
                    if vals and vals[0] == code and vals[2].upper()=="ITEM":
                        q = vals[6]
                        if q and str(q).isdigit():
                            total_assigned += int(q)
                    stack.extend(self.tree.get_children(iid))
            if current_std <= 0:
                return 10**12
            rem = current_std - total_assigned
            return rem if rem > 0 else 0
        if not batches_sorted:
            iid = self.tree.insert(parent_iid, "end", values=(
                code, desc, "ITEM",
                structural_kit or "-----",
                structural_module or "-----",
                std_qty, 0, "", "",
                "", "0"
            ))
            self.row_data[iid] = {
                'kit_number': kit_number,
                'module_number': module_number,
                'max_qty': 0,
                'management_mode': 'on_shelf',
                'expiry_key': ""
            }
            return
        for b in batches_sorted:
            expiry = b['expiry'] or ""
            final_qty = b['final_qty']
            line_id = b['line_id']
            remain_physical = self.remaining_available(code, expiry, b['management_mode'], final_qty)
            if remain_physical <= 0:
                continue
            remain_std = remaining_std_allowance_for_insertion()
            if remain_std <= 0:
                break
            allocate = min(remain_physical, remain_std)
            if allocate <= 0:
                continue
            iid = self.tree.insert(parent_iid, "end", values=(
                code, desc, "ITEM",
                structural_kit or "-----",
                structural_module or "-----",
                std_qty, allocate, expiry, "",
                str(line_id), str(allocate)  # hidden columns
            ))
            self.row_data[iid] = {
                'kit_number': kit_number,
                'module_number': module_number,
                'max_qty': final_qty,
                'management_mode': b['management_mode'],
                'expiry_key': expiry,
                'line_id': line_id
            }
            self.record_suggested(code, expiry, b['management_mode'], allocate)
        found = False
        for top in self.tree.get_children():
            stack = [top]
            while stack:
                iid = stack.pop()
                vals = self.tree.item(iid, "values")
                if vals and vals[0] == code and vals[2].upper()=="ITEM":
                    found = True
                stack.extend(self.tree.get_children(iid))
        if not found:
            iid = self.tree.insert(parent_iid, "end", values=(
                code, desc, "ITEM",
                structural_kit or "-----",
                structural_module or "-----",
                std_qty, 0, "", "",
                "", "0"
            ))
            self.row_data[iid] = {
                'kit_number': kit_number,
                'module_number': module_number,
                'max_qty': 0,
                'management_mode': 'on_shelf',
                'expiry_key': ""
            }

    # ------------- Add single code -------------
    def add_to_tree(self, code):
        mk = self.current_mode_key()
        desc = get_item_description(code)
        item_type = detect_type(code, desc).upper()
        kit_code = self.kit_var.get() if mk in ["add_module_kit","add_items_kit","add_items_module"] else None
        module_code = self.module_var.get() if mk == "add_items_module" else None
        kit_number = self.kit_number_var.get().strip() if self.kit_number_var.get() else None
        module_number = self.module_number_var.get().strip() if self.module_number_var.get() else None
        if mk == "add_module_kit":
            if not kit_code or not kit_number:
                custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                             lang.t("receive_kit.no_kit_number","Please select a Kit & Kit Number"),"error")
                return
        elif mk == "add_items_kit":
            if not kit_number:
                custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                             lang.t("receive_kit.no_kit_number","Please select a Kit Number"),"error")
                return
        elif mk == "add_items_module":
            if not module_code or not module_number:
                custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                             lang.t("receive_kit.no_module_number","Please select a Module & Module Number"),"error")
                return
        scenario_module_mode = (mk == "add_module_scenario")
        if item_type in ["KIT","MODULE"]:
            if item_type == "MODULE" and mk in ["add_module_kit","add_module_scenario"]:
                mn = self.ask_module_number(kit_number, code)
                if mn is None:
                    self.status_var.set(lang.t("receive_kit.module_number_cancelled","Cancelled"))
                    return
                module_number = mn
            iid = self.tree.insert("", "end", values=(
                code, desc, item_type,
                kit_code or "-----",
                module_code or "-----",
                0, 1, "", "",
                "", ""  # hidden
            ))
            if item_type == "KIT":
                self.tree.item(iid, tags=("kit",))
            else:
                self.tree.item(iid, tags=("module",))
            self.row_data[iid] = {
                'kit_number': kit_number if not scenario_module_mode else None,
                'module_number': module_number,
                'treecode': None
            }
            self.status_var.set(lang.t("receive_kit.added_item", f"Added {item_type} {code}"))
            return
        structural_kit = None if scenario_module_mode else kit_code
        structural_module = module_code
        self.insert_item_batches(
            code=code,
            parent_iid="",
            kit_number=kit_number,
            module_number=module_number,
            structural_kit=structural_kit,
            structural_module=structural_module,
            std_qty=0
        )
        self.status_var.set(lang.t("receive_kit.added_item", f"Added ITEM {code} (batches)"))

    # ------------- Add missing item dialog -------------
    def add_missing_item(self):
        if not self.selected_scenario_id:
            custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                         lang.t("receive_kit.no_scenario","Please select a scenario"),"error")
            return
        dlg = tk.Toplevel(self.parent)
        dlg.title(lang.t("receive_kit.add_missing","Add Missing Item"))
        dlg.geometry("560x360")
        dlg.transient(self.parent)
        dlg.grab_set()
        tk.Label(dlg, text=lang.t("receive_kit.search_item","Search Kit/Module/Item"),
                 font=("Helvetica",10,"bold")).pack(pady=6)
        sv = tk.StringVar()
        entry = tk.Entry(dlg, textvariable=sv, font=("Helvetica",10), width=60)
        entry.pack(padx=10, pady=4, fill=tk.X)
        lb = tk.Listbox(dlg, font=("Helvetica",10), height=14,
                        selectbackground="#2563EB", selectforeground="white")
        lb.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        def refresh(*_):
            lb.delete(0, tk.END)
            mk = self.current_mode_key()
            res = self.fetch_search_results(sv.get().strip(), self.selected_scenario_id, mk)
            for r in res:
                lb.insert(tk.END, f"{r['code']} - {r['description']}")
            if not res:
                lb.insert(tk.END, lang.t("receive_kit.no_items_found","No items found"))
        sv.trace_add("write", refresh)
        refresh()
        def choose(_=None):
            idxs = lb.curselection()
            if idxs:
                line = lb.get(idxs[0])
                c = line.split(" - ")[0]
                dlg.destroy()
                mk = self.current_mode_key()
                if mk == "receive_kit":
                    self.load_hierarchy(c)
                else:
                    self.add_to_tree(c)
            else:
                dlg.destroy()
        lb.bind("<Double-Button-1>", choose)
        entry.bind("<Return>", choose)
        entry.focus()
        _center_child(dlg, self.parent)
        dlg.wait_window()

    # ------------- Editing -------------
    def start_edit(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_index = int(col_id.replace("#","")) - 1
        # qty_to_receive(6), expiry_date(7), batch_no(8) only
        if col_index not in [6,7,8]:
            return
        self._start_edit_cell(row_id, col_index)

    def _start_edit_cell(self, row_id, col_index):
        try:
            bbox = self.tree.bbox(row_id, f"#{col_index+1}")
            if not bbox:
                return
            x,y,w,h = bbox
            orig_val = self.tree.set(row_id, self.tree["columns"][col_index])
            entry = tk.Entry(self.tree, font=("Helvetica",10))
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, orig_val)
            entry.focus()
            self.editing_cell = entry
            def finalize(_=None):
                new_val = entry.get().strip()
                if col_index == 6:
                    new_qty = int(new_val) if new_val.isdigit() else 0
                    t = self.tree.set(row_id, "type").upper()
                    if t == "ITEM":
                        code = self.tree.set(row_id, "code")
                        expiry = self.tree.set(row_id, "expiry_date") or ""
                        meta = self.row_data.get(row_id, {})
                        mgmt_mode = meta.get('management_mode')
                        if mgmt_mode == 'on_shelf':
                            max_qty = meta.get('max_qty')
                            key = (code, expiry, mgmt_mode)
                            prev_used = self.suggested_usage.get(key, 0)
                            old_qty = int(orig_val) if orig_val.isdigit() else 0
                            used_minus_self = prev_used - old_qty
                            physical_remaining = (max_qty if max_qty is not None else new_qty) - used_minus_self
                            physical_remaining = max(physical_remaining, 0)
                            if new_qty > physical_remaining:
                                new_qty = physical_remaining
                            std_allow = self._remaining_std_allowance(code) + old_qty
                            if new_qty > std_allow:
                                new_qty = std_allow
                            if new_qty < 0: new_qty = 0
                            self.suggested_usage[key] = used_minus_self + new_qty
                    self.tree.set(row_id, "qty_to_receive", str(new_qty))
                    # Mirror into hidden qty_out
                    self.tree.set(row_id, "qty_out", str(new_qty))
                    if t in ["KIT","MODULE"]:
                        self.update_child_quantities()
                elif col_index == 7:
                    if new_val:
                        parsed = parse_expiry(new_val)
                        self.tree.set(row_id, "expiry_date", parsed if parsed else "")
                    else:
                        self.tree.set(row_id, "expiry_date", "")
                elif col_index == 8:
                    if len(new_val) > 30:
                        new_val = new_val[:30]
                    self.tree.set(row_id, "batch_no", new_val)
                entry.destroy()
                self.editing_cell = None
                self.update_parent_expiry()
            entry.bind("<Return>", finalize)
            entry.bind("<Tab>", finalize)
            entry.bind("<FocusOut>", finalize)
            entry.bind("<Escape>", lambda e: (entry.destroy(), setattr(self, "editing_cell", None)))
        except Exception as e:
            logging.error(f"[start_edit_cell] {e}")

    # ------------- Quantity propagation -------------
    def update_child_quantities(self):
        try:
            for iid in self.tree.get_children():
                self._update_subtree_quantities(iid)
        except Exception as e:
            logging.error(f"[update_child_quantities] {e}")

    def _update_subtree_quantities(self, iid):
        vals = self.tree.item(iid, "values")
        if not vals: return
        t = vals[2].upper()
        if t in ["KIT","MODULE"]:
            for child in self.tree.get_children(iid):
                self._update_subtree_quantities(child)

    # ------------- Expiry propagation -------------
    def update_parent_expiry(self):
        try:
            for iid in self.tree.get_children():
                self._update_parent_expiry_recursive(iid)
        except Exception as e:
            logging.error(f"[update_parent_expiry] {e}")

    def _update_parent_expiry_recursive(self, iid):
        vals = self.tree.item(iid, "values")
        if not vals: return
        t = vals[2].upper()
        child_exp = []
        for c in self.tree.get_children(iid):
            self._update_parent_expiry_recursive(c)
            cv = self.tree.item(c,"values")
            if not cv: continue
            if cv[2].upper()=="ITEM":
                p = parse_expiry(cv[7])
                if p: child_exp.append(p)
            else:
                cp = parse_expiry(cv[7])
                if cp: child_exp.append(cp)
        if t in ["KIT","MODULE"] and child_exp:
            self.tree.set(iid, "expiry_date", min(child_exp))
        code = vals[0]
        qty = vals[6]
        expv = self.tree.set(iid, "expiry_date")
        if qty and str(qty).isdigit() and int(qty)>0 and check_expiry_required(code) and not parse_expiry(expv):
            self.tree.item(iid, tags=("light_red",))
        else:
            if t == "KIT":
                self.tree.item(iid, tags=("kit",))
            elif t == "MODULE":
                self.tree.item(iid, tags=("module",))

    # ------------- Navigation & Context Menu -------------
    def navigate_tree(self, event):
        if self.editing_cell:
            return
        all_rows = []
        def collect(p=""):
            for r in self.tree.get_children(p):
                all_rows.append(r)
                collect(r)
        collect()
        if not all_rows:
            return
        sel = self.tree.selection()
        current = sel[0] if sel else all_rows[0]
        if current not in all_rows:
            current = all_rows[0]
        idx = all_rows.index(current)
        if event.keysym == "Up" and idx > 0:
            self.tree.selection_set(all_rows[idx-1]); self.tree.focus(all_rows[idx-1])
        elif event.keysym == "Down" and idx < len(all_rows)-1:
            self.tree.selection_set(all_rows[idx+1]); self.tree.focus(all_rows[idx+1])

    def show_context_menu(self, event):
        menu = tk.Menu(self.tree, tearoff=0)
        menu.add_command(label=lang.t("receive_kit.add_new_row","Add New Row"),
                         command=lambda: None)
        menu.post(event.x_root, event.y_root)

    # ------------- Uniqueness helpers -------------
    def is_kit_number_unique(self, kit_number: str) -> bool:
        if not kit_number or kit_number.lower()=="none":
            return False
        conn = connect_db()
        if conn is None: return False
        cur = conn.cursor()
        try:
            cur.execute("""SELECT COUNT(*) FROM stock_data
                           WHERE kit_number=? AND kit_number!='None'""",(kit_number.strip(),))
            return cur.fetchone()[0] == 0
        finally:
            cur.close()
            conn.close()

    def is_module_number_unique(self, kit_number: str, module_number: str) -> bool:
        if not module_number or module_number.lower()=="none":
            return False
        conn = connect_db()
        if conn is None: return False
        cur = conn.cursor()
        try:
            cur.execute("""SELECT COUNT(*) FROM stock_data
                           WHERE module_number=? AND module_number!='None'""",(module_number.strip(),))
            return cur.fetchone()[0] == 0
        finally:
            cur.close()
            conn.close()

    def generate_unique_id(self, scenario_id, kit, module, item, std_qty, exp_date, kit_number, module_number):
        return f"{scenario_id}/{kit or 'None'}/{module or 'None'}/{item or 'None'}/{std_qty}/{exp_date or 'None'}/{kit_number or 'None'}/{module_number or 'None'}"

    # ------------- Document & Logging -------------
    def generate_document_number(self, in_type_text: str) -> str:
        project_name, project_code = fetch_project_details()
        project_code = (project_code or "PRJ").upper()
        base_map = {
            "In MSF":"IMSF","In Local Purchase":"ILP","In from Quarantine":"IFQ",
            "In Donation":"IDN","Return from End User":"IREU","In Supply Non-MSF":"ISNM",
            "In Borrowing":"IBR","In Return of Loan":"IRL","In Correction of Previous Transaction":"ICOR",
            "In from on-shelf items":"INTERNAL"
        }
        raw = (in_type_text or "").strip()
        abbr = None
        norm_raw = re.sub(r'[^a-z0-9]+','', raw.lower())
        for k,v in base_map.items():
            if re.sub(r'[^a-z0-9]+','', k.lower()) == norm_raw:
                abbr = v; break
        if not abbr:
            tokens = re.split(r'\s+', raw.upper())
            stop = {"OF","FROM","THE","AND","DE","DU","DES","LA","LE","LES"}
            letters=[]
            for t in tokens:
                if not t or t in stop: continue
                if t=="MSF": letters.append("MSF")
                else: letters.append(t[0])
            abbr = "".join(letters) or (raw[:4].upper() or "DOC")
            abbr = abbr[:8]
        now = datetime.now()
        prefix = f"{now.year:04d}/{now.month:02d}/{project_code}/{abbr}"
        conn = connect_db()
        serial = 1
        if conn:
            cur = conn.cursor()
            try:
                cur.execute("""
                    SELECT document_number FROM stock_transactions
                     WHERE document_number LIKE ?
                     ORDER BY document_number DESC LIMIT 1
                """,(prefix+"/%",))
                r = cur.fetchone()
                if r and r[0]:
                    last = r[0].rsplit('/',1)[-1]
                    if last.isdigit():
                        serial = int(last)+1
            finally:
                cur.close()
                conn.close()
        doc = f"{prefix}/{serial:04d}"
        self.current_document_number = doc
        return doc

    def log_transaction(self,
                        unique_id: str,
                        code: str,
                        description: str,
                        expiry_date: str,
                        batch_number: str,
                        scenario: str,
                        kit: str,
                        module: str,
                        qty_in,
                        in_type: str,
                        qty_out,
                        out_type: str,
                        third_party: str,
                        end_user: str,
                        remarks: str,
                        movement_type: str,
                        document_number: str):
        """
        Insert a single stock transaction row.

        Note: Per your new requirement, save_all() now calls this twice
        for each ITEM row:
          1) Incoming (qty_in > 0, qty_out = None, in_type set, out_type None)
          2) Outgoing mirror (qty_in = None, qty_out > 0, in_type None, out_type = original in_type)

        Parameters may be None; they are written as NULL in SQLite.
        """
        conn = connect_db()
        if conn is None:
            return
        cur = conn.cursor()
        try:
            cur.execute("""
                INSERT INTO stock_transactions
                (Date, Time, unique_id, code, Description, Expiry_date, Batch_Number,
                 Scenario, Kit, Module, Qty_IN, IN_Type, Qty_Out, Out_Type,
                 Third_Party, End_User, Remarks, Movement_Type, document_number)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,(
                datetime.today().strftime('%Y-%m-%d'),
                datetime.now().strftime('%H:%M:%S'),
                unique_id,
                code,
                description,
                expiry_date,
                batch_number,
                scenario,
                kit,
                module,
                qty_in,
                in_type,
                qty_out,
                out_type,
                third_party,
                end_user,
                remarks,
                movement_type,
                document_number
            ))
            conn.commit()
        except sqlite3.Error as e:
            conn.rollback()
            logging.error(f"[log_transaction] {e}")
        finally:
            cur.close()
            conn.close()


    # ------------- Uniqueness enforcement -------------
    def _enforce_unique_numbers_before_save(self):
        kits = {}
        modules = {}
        for iid, meta in self.row_data.items():
            vals = self.tree.item(iid, "values")
            if not vals: continue
            t = (vals[2] or "").upper()
            kn = meta.get('kit_number')
            mn = meta.get('module_number')
            if t == "KIT" and kn and kn.lower() != 'none':
                kits.setdefault(kn, []).append(iid)
            if t == "MODULE" and mn and mn.lower() != 'none':
                modules.setdefault(mn, []).append(iid)
        dup_kits = {k: v for k,v in kits.items() if len(v) > 1}
        dup_modules = {m: v for m,v in modules.items() if len(v) > 1}
        if not dup_kits and not dup_modules:
            return True
        for kit_number, iids in dup_kits.items():
            for iid in iids[1:]:
                while True:
                    new_kn = simpledialog.askstring(
                        lang.t("receive_kit.kit_number","Kit Number"),
                        f"Duplicate Kit Number '{kit_number}'. Enter a new Kit Number:",
                        parent=self.parent
                    )
                    if new_kn is None:
                        custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                                     f"Save aborted (duplicate Kit Number {kit_number}).","error")
                        return False
                    new_kn = new_kn.strip()
                    if not new_kn:
                        continue
                    if self.is_kit_number_unique(new_kn):
                        self.row_data[iid]['kit_number'] = new_kn
                        break
                    custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                                 f"Kit Number '{new_kn}' already exists globally.","error")
        for mod_number, iids in dup_modules.items():
            for iid in iids[1:]:
                while True:
                    new_mn = simpledialog.askstring(
                        lang.t("receive_kit.module_number","Module Number"),
                        f"Duplicate Module Number '{mod_number}'. Enter a new Module Number:",
                        parent=self.parent
                    )
                    if new_mn is None:
                        custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                                     f"Save aborted (duplicate Module Number {mod_number}).","error")
                        return False
                    new_mn = new_mn.strip()
                    if not new_mn:
                        continue
                    if self.is_module_number_unique(None, new_mn):
                        self.row_data[iid]['module_number'] = new_mn
                        break
                    custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                                 f"Module Number '{new_mn}' already exists globally.","error")
        return True

    # ------------- 70% rule -------------
    def _validate_global_70_rule(self):
        per_code = {}
        for top in self.tree.get_children():
            stack=[top]
            while stack:
                iid=stack.pop()
                vals=self.tree.item(iid,"values")
                if vals:
                    code, _, t, _, _, std_qty, qty, _, _, _, _ = vals
                    if (t or "").upper()=="ITEM":
                        std = int(std_qty) if str(std_qty).isdigit() else 0
                        issued = int(qty) if str(qty).isdigit() else 0
                        if code not in per_code:
                            per_code[code]={'std':std,'issued':issued}
                        else:
                            per_code[code]['issued'] += issued
                            if std > per_code[code]['std']:
                                per_code[code]['std'] = std
                stack.extend(self.tree.get_children(iid))
        if not per_code:
            return True
        total = len(per_code)
        covered = 0
        uncovered=[]
        for c,st in per_code.items():
            if st['std']<=0 or st['issued']>=st['std']:
                covered+=1
            else:
                uncovered.append((c, st['std'], st['issued']))
        coverage = covered/total if total else 1
        if coverage >= 0.70:
            return True
        uncovered.sort(key=lambda x: (x[1]-x[2]), reverse=True)
        sample = "\n".join([f"{c}: {issd}/{std} ({(issd/std*100):.1f}%)" for c,std,issd in uncovered[:15]])
        more = f"\n... {len(uncovered)-15} more." if len(uncovered)>15 else ""
        mode_key = self.current_mode_key()
        context = {
            "receive_kit":"kit",
            "add_module_kit":"module",
            "add_module_scenario":"module",
            "add_items_kit":"kit",
            "add_items_module":"module",
            "add_standalone":"operation"
        }.get(mode_key,"operation")
        msg = (f"At least 70% of items must have enough stock. Current: {coverage*100:.2f}% "
               f"({covered}/{total}). The {context} cannot be created.\n\nBelow standard (sample):\n"
               f"{sample}{more}")
        custom_popup(self.parent, lang.t("dialog_titles.error","Error"), msg, "error")
        return False

    # ------------- Save logic -------------
    def save_all(self):
        """
        Save current tree (kit/module/items) into stock_data and log transactions.

        Updated behavior:
          - Still consumes existing on-shelf batches (qty_out on original lines).
          - Adds new stock_data lines for built kit/module/item composition (qty_in).
          - Logs TWO transaction rows per ITEM with qty_to_receive > 0:
              1) Incoming:  Qty_IN = qty_to_receive, IN_Type = trans_type_var value.
              2) Outgoing mirror: Qty_Out = hidden 'qty_out' column value,
                 Out_Type = original IN_Type value, Qty_IN = NULL.
        """
        self.ensure_vars_ready()
        if self.role not in ["admin","manager"]:
            custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                         lang.t("receive_kit.no_permission","Only admin or manager roles can save changes."),"error")
            return
        if not self.tree.get_children():
            custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                         lang.t("receive_kit.no_rows","No rows to save."),"error")
            return

        in_type = self.trans_type_var.get()
        if not in_type:
            custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                         "Internal error: IN Type missing.","error")
            return

        # Enforce uniqueness of kit/module numbers (interactive resolution)
        if not self._enforce_unique_numbers_before_save():
            return

        # 70% coverage rule
        if not self._validate_global_70_rule():
            return

        doc = self.generate_document_number(in_type)
        scenario_name = self.scenario_map.get(self.selected_scenario_id, "Unknown")

        saved = 0
        errors = []
        all_iids = []

        def collect(root=""):
            for r in self.tree.get_children(root):
                all_iids.append(r)
                collect(r)
        collect()

        # 1) Consume existing stock lines (qty_out) by line_id
        consumed_lines = set()
        for iid in all_iids:
            vals = self.tree.item(iid,"values")
            if not vals:
                continue
            (code, description, t, kit_col, module_col, std_qty,
             qty_to_receive, expiry_date, batch_no, line_id, qty_out_hidden) = vals
            if (t or "").upper() != "ITEM":
                continue
            if not line_id or not line_id.isdigit():
                continue
            if not qty_out_hidden or not qty_out_hidden.isdigit():
                continue
            consume_qty = int(qty_out_hidden)
            if consume_qty <= 0:
                continue
            if line_id in consumed_lines:
                continue
            # Update existing stock_data line's qty_out
            StockData.consume_by_line_id(int(line_id), consume_qty)
            consumed_lines.add(line_id)

        # 2) Insert new lines (composition) & log transactions twice per ITEM
        for iid in all_iids:
            vals = self.tree.item(iid,"values")
            if not vals:
                continue
            (code, description, t, kit_col, module_col, std_qty,
             qty_to_receive, expiry_date, batch_no, line_id, qty_out_hidden) = vals

            if not qty_to_receive or not str(qty_to_receive).isdigit():
                continue
            qty_to_receive = int(qty_to_receive)
            if qty_to_receive <= 0:
                continue

            parsed_exp = parse_expiry(expiry_date) if expiry_date else None

            # If item requires an expiry, enforce
            if (t or "").upper() == "ITEM" and check_expiry_required(code) and not parsed_exp:
                self.tree.item(iid, tags=("light_red",))
                errors.append(code)
                continue

            meta = self.row_data.get(iid, {})
            kit_number = meta.get('kit_number')
            module_number = meta.get('module_number')

            kit_struct = kit_col if kit_col and kit_col != "-----" else None
            module_struct = module_col if module_col and module_col != "-----" else None
            item_part = code if (t or "").upper() == "ITEM" else None
            std_numeric = int(std_qty) if str(std_qty).isdigit() else 0

            unique_id = self.generate_unique_id(
                self.selected_scenario_id,
                kit_struct,
                module_struct,
                item_part,
                std_numeric,
                parsed_exp,
                kit_number,
                module_number
            )

            # Add/update the composed line in stock_data (qty_in)
            StockData.add_or_update(unique_id,
                                    scenario=scenario_name,
                                    qty_in=qty_to_receive,
                                    exp_date=parsed_exp,
                                    kit_number=kit_number,
                                    module_number=module_number)

            # First transaction (incoming)
            self.log_transaction(
                unique_id=unique_id,
                code=code,
                description=description,
                expiry_date=parsed_exp,
                batch_number=batch_no or None,
                scenario=scenario_name,
                kit=kit_number,
                module=module_number,
                qty_in=qty_to_receive,
                in_type=in_type,
                qty_out=None,
                out_type=None,
                third_party=self.third_party_var.get() or None,
                end_user=self.end_user_var.get() or None,
                remarks=self.remarks_entry.get().strip() if self.remarks_entry else None,
                movement_type=self.mode_var.get() or "stock_in",
                document_number=doc
            )

            # Second transaction (outgoing mirror) if hidden qty_out is valid
            if qty_out_hidden and str(qty_out_hidden).isdigit():
                out_q = int(qty_out_hidden)
                if out_q > 0:
                    self.log_transaction(
                        unique_id=unique_id,
                        code=code,
                        description=description,
                        expiry_date=parsed_exp,
                        batch_number=batch_no or None,
                        scenario=scenario_name,
                        kit=kit_number,
                        module=module_number,
                        qty_in=None,          # no incoming qty this row
                        in_type=None,         # IN_Type left NULL
                        qty_out=out_q,        # outgoing quantity
                        out_type=in_type,     # mirror original IN type here
                        third_party=self.third_party_var.get() or None,
                        end_user=self.end_user_var.get() or None,
                        remarks=self.remarks_entry.get().strip() if self.remarks_entry else None,
                        movement_type=self.mode_var.get() or "stock_in",
                        document_number=doc
                    )

            saved += 1

        if errors:
            custom_popup(self.parent, lang.t("dialog_titles.error","Error"),
                         lang.t("receive_kit.invalid_expiry",
                                f"Valid expiry required for: {', '.join(errors)}"),"error")
            return

        custom_popup(self.parent, lang.t("dialog_titles.success","Success"),
                     lang.t("receive_kit.save_success",
                            f"Kit received successfully. Logged {saved} item rows (doubled to {saved*2} transactions)."),"info")

        self.status_var.set(
            lang.t("receive_kit.document_number_generated",
                   f"Saved {saved} rows (transactions: {saved*2}). Document Number: {doc}")
        )
        self.clear_form()
    # ------------- Export -------------
    def export_data(self, rows_to_export=None):
        if not self.tree.get_children():
            self.status_var.set(lang.t("receive_kit.export_cancelled","Export cancelled"))
            return
        default_dir = "D:/ISEPREP"
        if not os.path.exists(default_dir):
            try: os.makedirs(default_dir)
            except: pass
        current_time = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        in_type_raw = self.trans_type_var.get() or lang.t("receive_kit.unknown","Unknown")
        movement_type_raw = self.mode_var.get() or lang.t("receive_kit.unknown","Unknown")
        doc_number = getattr(self, "current_document_number", "")
        safe = lambda s: (re.sub(r'[^A-Za-z0-9]+','_', s or '') or "Unknown").strip('_')
        file_name = f"in_kit_{safe(movement_type_raw)}_{safe(in_type_raw)}_{current_time}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile=file_name,
            initialdir=default_dir
        )
        if not path:
            self.status_var.set(lang.t("receive_kit.export_cancelled","Export cancelled"))
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws_title = lang.t("receive_kit.title","Receive Kit-Module")
        ws.title = ws_title[:31]
        project_name, project_code = fetch_project_details()
        ws['A1'] = f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}    Document Number: {doc_number}"
        ws['A2'] = f"{ws_title} – Movement: {movement_type_raw}"
        ws['A3'] = f"{project_name} - {project_code}"
        ws['A4'] = f"{lang.t('receive_kit.in_type','IN Type')}: {in_type_raw}"
        ws['A5'] = f"{lang.t('receive_kit.movement_type','Movement Type')}: {movement_type_raw}"
        ws.append([])
        headers = [
            lang.t("receive_kit.code","Code"),
            lang.t("receive_kit.description","Description"),
            lang.t("receive_kit.type","Type"),
            lang.t("receive_kit.kit","Kit"),
            lang.t("receive_kit.module","Module"),
            lang.t("receive_kit.std_qty","Std Qty"),
            lang.t("receive_kit.qty_to_receive","Qty to Receive"),
            lang.t("receive_kit.expiry_date","Expiry Date"),
            lang.t("receive_kit.batch_no","Batch No"),
            "line_id",
            "qty_out_consumed"
        ]
        ws.append(headers)
        def append_rows_recursive(iid):
            vals = self.tree.item(iid,"values")
            if vals:
                qty = vals[6]
                if qty and str(qty).isdigit() and int(qty)>0 and vals[2].upper()=="ITEM":
                    ws.append(list(vals))
            for c in self.tree.get_children(iid):
                append_rows_recursive(c)
        for iid in self.tree.get_children():
            append_rows_recursive(iid)
        for col in ws.columns:
            max_len = 0
            letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[letter].width = min(max_len+2, 55)
        wb.save(path)
        custom_popup(self.parent, lang.t("dialog_titles.success","Success"),
                     lang.t("receive_kit.export_success", f"Export successful: {path}"), "info")
        self.status_var.set(lang.t("receive_kit.export_success", f"Export successful: {path}"))

# ------------- Standalone harness -------------
if __name__ == "__main__":
    root = tk.Tk()
    root.title("In Kit (On-Shelf Batches)")
    class DummyApp: pass
    DummyApp.role = "admin"
    app = DummyApp()
    StockInKit(root, app, role="admin")
    root.geometry("1400x850")
    root.mainloop()
