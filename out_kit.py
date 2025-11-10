# ============================= out_kit.py (Part 1/6) =============================
# Break Kit / Module
# Derived and extended from dispatch_kit.py to implement "Break" logic:
#  - Movement primary mode renamed "Break Kit" (internal key keeps dispatch_Kit for reuse of logic patterns)
#  - OUT Type fixed: "Internal move from in-box items"
#  - Dual logging: for every OUT transaction, a mirror IN transaction is also recorded (Qty_IN + IN_Type)
#  - Hidden columns: line_id, qty_in_hidden (mirrors user-entered qty_to_issue)
#  - Tree ordering restored using treecode via kit_items index
#  - Color highlighting: distinct styles for headers and Kit/Module data rows
#  - Quantities for kits/modules propagate to sub-levels; star/★ marking editable cells
#  - Document number generation with "BRK" abbreviation
#
# ================================================================================
import tkinter as tk
from tkinter import ttk, filedialog
import sqlite3
import logging
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from datetime import datetime
from popup_utils import custom_popup, custom_askyesno, custom_dialog
import os


from db import connect_db
from manage_items import get_item_description, detect_type
from language_manager import lang

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

OUT_TYPE_FIXED = "Internal move from in-box items"

def fetch_project_details():
    conn = connect_db()
    if conn is None:
        logging.error("[BREAK] DB connection failed (fetch_project_details)")
        return "Unknown Project", "Unknown Code"
    cur = conn.cursor()
    try:
        cur.execute("SELECT project_name, project_code FROM project_details LIMIT 1")
        row = cur.fetchone()
        return (row[0] if row and row[0] else "Unknown Project",
                row[1] if row and row[1] else "Unknown Code")
    except sqlite3.Error as e:
        logging.error(f"[BREAK] fetch_project_details error: {e}")
        return "Unknown Project", "Unknown Code"
    finally:
        cur.close()
        conn.close()

def configure_db_pragmas():
    conn = connect_db()
    if conn is None:
        return
    try:
        cur = conn.cursor()
        cur.execute("PRAGMA journal_mode=WAL;")
        cur.execute("PRAGMA busy_timeout=5000;")
        conn.commit()
    except Exception:
        pass
    finally:
        try: conn.close()
        except: pass


class StockOutKit(tk.Frame):
    """
    Break Kit / Module logic with dual transaction logging.
    Retains dispatch-like structure but adapted for "break" internal movement.
    """

    def __init__(self, parent, app, role="supervisor"):
        super().__init__(parent)
        self.parent = parent
        self.app = app
        self.role = role.lower()

        try:
            configure_db_pragmas()
        except Exception:
            pass

        # UI variable references
        self.scenario_var = tk.StringVar()
        self.mode_var = tk.StringVar()
        self.Kit_var = tk.StringVar()
        self.Kit_number_var = tk.StringVar()
        self.module_var = tk.StringVar()
        self.module_number_var = tk.StringVar()
        self.out_type_var = tk.StringVar(value=OUT_TYPE_FIXED)
        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar(value=lang.t("break_kit.ready", "Ready"))

        # Widgets placeholders
        self.scenario_cb = None
        self.mode_cb = None
        self.Kit_cb = None
        self.Kit_number_cb = None
        self.module_cb = None
        self.module_number_cb = None
        self.tree = None

        # Data state
        self.scenario_map = self.fetch_scenario_map()
        self.selected_scenario_id = None
        self.selected_scenario_name = None
        self.mode_definitions = []
        self.mode_label_to_key = {}
        self.full_items = []      # enriched rows (without headers)
        self.row_data = {}        # iid -> metadata
        self.search_min_chars = 2
        self.editing_cell = None
        self._item_index_cache = {}

        if self.parent and self.parent.winfo_exists():
            self.pack(fill="both", expand=True)
            self.after(50, self.render_ui)

    # -------------------- Scenario / Modes --------------------
    def fetch_scenario_map(self):
        conn = connect_db()
        if conn is None: return {}
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("SELECT scenario_id, name FROM scenarios ORDER BY name")
            rows = cur.fetchall()
            return {str(r['scenario_id']): r['name'] for r in rows}
        except sqlite3.Error as e:
            logging.error(f"[BREAK] fetch_scenario_map error: {e}")
            return {}
        finally:
            cur.close()
            conn.close()

    def build_mode_definitions(self):
        scenario = self.selected_scenario_name or ""
        self.mode_definitions = [
            ("dispatch_Kit",  lang.t("break_kit.mode_break_kit", "Break Kit")),
            ("issue_standalone",
             lang.t("dispatch_Kit.mode_issue_standalone", "Issue standalone item/s from {scenario}", scenario=scenario)),
            ("issue_module_scenario",
             lang.t("dispatch_Kit.mode_issue_module_scenario", "Issue module from {scenario}", scenario=scenario)),
            ("issue_module_Kit",
             lang.t("dispatch_Kit.mode_issue_module_Kit", "Issue module from a Kit")),
            ("issue_items_Kit",
             lang.t("dispatch_Kit.mode_issue_items_Kit", "Issue items from a Kit")),
            ("issue_items_module",
             lang.t("dispatch_Kit.mode_issue_items_module", "Issue items from a module"))
        ]
        self.mode_label_to_key = {lbl: key for key, lbl in self.mode_definitions}

    def current_mode_key(self):
        return self.mode_label_to_key.get(self.mode_var.get())

    def load_scenarios(self):
        vals = [f"{sid} - {nm}" for sid, nm in self.scenario_map.items()]
        self.scenario_cb['values'] = vals
        if vals:
            self.scenario_cb.current(0)
            self.on_scenario_selected()

    def on_scenario_selected(self, event=None):
        sel = self.scenario_var.get()
        if not sel:
            self.selected_scenario_id = None
            self.selected_scenario_name = None
            return
        self.selected_scenario_id = sel.split(" - ")[0]
        self.selected_scenario_name = sel.split(" - ", 1)[1] if " - " in sel else ""
        self.build_mode_definitions()
        self.mode_cb['values'] = [lbl for _, lbl in self.mode_definitions]
        if self.mode_definitions:
            self.mode_var.set(self.mode_definitions[0][1])
        self.on_mode_changed()

    def on_mode_changed(self, event=None):
        mode_key = self.current_mode_key()
        for cb in [self.Kit_cb, self.Kit_number_cb, self.module_cb, self.module_number_cb]:
            cb.config(state="disabled")
        self.Kit_var.set("")
        self.Kit_number_var.set("")
        self.module_var.set("")
        self.module_number_var.set("")
        self.Kit_number_cb['values'] = []
        self.module_number_cb['values'] = []
        self.full_items = []
        self.clear_table_only()
        if not self.selected_scenario_id:
            return
        if mode_key in ("dispatch_Kit", "issue_items_Kit", "issue_module_Kit"):
            self.Kit_cb.config(state="readonly")
            self.Kit_cb['values'] = self.fetch_Kits(self.selected_scenario_id)
        if mode_key in ("issue_items_module", "issue_module_Kit", "issue_module_scenario"):
            self.module_cb.config(state="readonly")
            self.module_cb['values'] = self.fetch_all_modules(self.selected_scenario_id)
        if mode_key == "issue_standalone":
            self.populate_standalone_items()

# ============================= out_kit.py (Part 2/6) =============================
    # -------------------- Index / Treecode --------------------
    def ensure_item_index(self, scenario_id):
        if (self._item_index_cache.get("scenario_id") == scenario_id and
            self._item_index_cache.get("flat_map")):
            return
        self._item_index_cache = {
            "scenario_id": scenario_id,
            "flat_map": {},
            "triple_map": {}
        }
        conn = connect_db()
        if conn is None: return
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT code, kit, module, item, treecode, level
                  FROM kit_items
                 WHERE scenario_id=?
            """,(scenario_id,))
            for r in cur.fetchall():
                code = (r["code"] or "").strip()
                kit = (r["kit"] or "").strip()
                module = (r["module"] or "").strip()
                item = (r["item"] or "").strip()
                treecode = r["treecode"]
                entry = {
                    "code": code,
                    "kit": kit,
                    "module": module,
                    "item": item,
                    "treecode": treecode,
                    "level": (r["level"] or "").lower()
                }
                for tok in (code, kit, module, item):
                    if tok and tok.upper() != "NONE":
                        self._item_index_cache["flat_map"].setdefault(tok, entry)
                self._item_index_cache["triple_map"][(kit or None,
                                                      module or None,
                                                      item or None)] = entry
        except sqlite3.Error as e:
            logging.error(f"[BREAK] ensure_item_index error: {e}")
        finally:
            cur.close()
            conn.close()

    @staticmethod
    def parse_unique_id(unique_id: str):
        kit = module = item = None
        std_qty = 1
        parts = unique_id.split("/") if unique_id else []
        if len(parts) >= 2: kit = parts[1] or None
        if len(parts) >= 3: module = parts[2] or None
        if len(parts) >= 4: item = parts[3] or None
        if len(parts) >= 5:
            try:
                v = int(parts[4])
                if v > 0: std_qty = v
            except:
                std_qty = 1
        return {"kit": kit, "module": module, "item": item, "std_qty": std_qty}

    def enrich_stock_row(self, scenario_id, unique_id, final_qty, exp_date,
                         Kit_number, module_number, line_id):
        self.ensure_item_index(scenario_id)
        parsed = self.parse_unique_id(unique_id)
        kit_code = parsed["kit"]
        module_code = parsed["module"]
        item_code = parsed["item"]
        std_qty = parsed["std_qty"]

        if item_code and item_code.upper() != "NONE":
            display_code = item_code
            forced = "Item"
            triple_key = (kit_code, module_code, item_code)
        elif module_code and module_code.upper() != "NONE":
            display_code = module_code
            forced = "Module"
            triple_key = (kit_code, module_code, None)
        else:
            display_code = kit_code
            forced = "Kit"
            triple_key = (kit_code, None, None)

        idx = self._item_index_cache
        entry = idx.get("triple_map", {}).get(triple_key) or \
                idx.get("flat_map", {}).get(display_code or "")
        treecode = entry.get("treecode") if entry else None

        description = get_item_description(display_code or "")
        detected = detect_type(display_code or "", description) or forced
        m = {"KIT":"Kit","MODULE":"Module","ITEM":"Item"}
        detected_norm = m.get(detected.upper(), forced)
        final_type = forced if forced in ("Module","Item") else ("Kit" if detected_norm not in ("Kit",) else detected_norm)

        return {
            "unique_id": unique_id,
            "code": display_code or "",
            "description": description,
            "type": final_type,
            "Kit": kit_code or "-----",
            "module": module_code or "-----",
            "current_stock": final_qty,
            "expiry_date": exp_date or "",
            "batch_no": "",
            "Kit_number": Kit_number,
            "module_number": module_number,
            "std_qty": std_qty if final_type == "Item" else None,
            "line_id": line_id,
            "treecode": treecode
        }

    # -------------------- Stock Fetch --------------------
    def fetch_stock_data_for_Kit_number(self, scenario_id, Kit_number, Kit_code=None):
        self.ensure_item_index(scenario_id)
        conn = connect_db()
        if conn is None: return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT unique_id, final_qty, exp_date, kit_number, module_number, line_id
                  FROM stock_data
                 WHERE kit_number=? AND unique_id LIKE ? AND final_qty > 0
            """,(Kit_number, f"{scenario_id}/%"))
            rows = cur.fetchall()
            return [self.enrich_stock_row(scenario_id, r["unique_id"], r["final_qty"],
                                          r["exp_date"], r["kit_number"], r["module_number"], r["line_id"])
                    for r in rows]
        except sqlite3.Error as e:
            logging.error(f"[BREAK] fetch_stock_data_for_Kit_number error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_stock_data_for_module_number(self, scenario_id, module_number, Kit_code=None, module_code=None):
        self.ensure_item_index(scenario_id)
        conn = connect_db()
        if conn is None: return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT unique_id, final_qty, exp_date, kit_number, module_number, line_id
                  FROM stock_data
                 WHERE module_number=? AND unique_id LIKE ? AND final_qty > 0
            """,(module_number, f"{scenario_id}/%"))
            rows = cur.fetchall()
            return [self.enrich_stock_row(scenario_id, r["unique_id"], r["final_qty"],
                                          r["exp_date"], r["kit_number"], r["module_number"], r["line_id"])
                    for r in rows]
        except sqlite3.Error as e:
            logging.error(f"[BREAK] fetch_stock_data_for_module_number error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_standalone_stock_items(self, scenario_id):
        self.ensure_item_index(scenario_id)
        conn = connect_db()
        if conn is None: return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT unique_id, final_qty, exp_date, kit_number, module_number, line_id
                  FROM stock_data
                 WHERE final_qty>0
                   AND (kit_number IS NULL OR kit_number='None')
                   AND (module_number IS NULL OR module_number='None')
                   AND unique_id LIKE ?
            """,(f"{scenario_id}/%",))
            rows = cur.fetchall()
            return [self.enrich_stock_row(scenario_id, r["unique_id"], r["final_qty"],
                                          r["exp_date"], r["kit_number"], r["module_number"], r["line_id"])
                    for r in rows]
        except sqlite3.Error as e:
            logging.error(f"[BREAK] fetch_standalone_stock_items error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    # -------------------- Structural fetch helpers --------------------
    def fetch_Kits(self, scenario_id):
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT code FROM kit_items
                 WHERE scenario_id=? AND level='primary'
                 ORDER BY code
            """,(scenario_id,))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[BREAK] fetch_Kits error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_all_modules(self, scenario_id):
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT code FROM kit_items
                 WHERE scenario_id=? AND level='secondary'
                 ORDER BY code
            """,(scenario_id,))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[BREAK] fetch_all_modules error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_available_Kit_numbers(self, scenario_id, Kit_code=None):
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            if Kit_code:
                cur.execute("""
                    SELECT DISTINCT kit_number
                      FROM stock_data
                     WHERE kit_number IS NOT NULL
                       AND kit_number!='None'
                       AND unique_id LIKE ?
                       AND unique_id LIKE ?
                     ORDER BY kit_number
                """,(f"{scenario_id}/%", f"{scenario_id}/{Kit_code}/%"))
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
            logging.error(f"[BREAK] fetch_available_Kit_numbers error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_module_numbers(self, scenario_id, Kit_code=None, module_code=None):
        conn = connect_db()
        if conn is None: return []
        cur = conn.cursor()
        try:
            where = ["module_number IS NOT NULL","module_number!='None'","unique_id LIKE ?"]
            params = [f"{scenario_id}/%"]
            if Kit_code:
                where.append("unique_id LIKE ?")
                params.append(f"{scenario_id}/{Kit_code}/%")
            if module_code:
                where.append("unique_id LIKE ?")
                params.append(f"{scenario_id}/%/{module_code}/%")
            sql = f"""
                SELECT DISTINCT module_number FROM stock_data
                 WHERE {' AND '.join(where)} ORDER BY module_number
            """
            cur.execute(sql, params)
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[BREAK] fetch_module_numbers error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

# ============================= out_kit.py (Part 3/6) =============================
    # -------------------- UI Rendering --------------------
    def render_ui(self):
        if not self.parent: return
        for w in self.parent.winfo_children():
            try: w.destroy()
            except: pass

        title_frame = tk.Frame(self.parent, bg="#F0F4F8")
        title_frame.pack(fill="x")
        tk.Label(title_frame, text=lang.t("break_kit.title","Break Kit/Module"),
                 font=("Helvetica", 20, "bold"), bg="#F0F4F8").pack(pady=(10,0))

        main = tk.Frame(self.parent, bg="#F0F4F8")
        main.pack(fill="both", expand=True, padx=10, pady=10)

        tk.Label(main, text=lang.t("receive_Kit.scenario","Scenario:"), bg="#F0F4F8")\
            .grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.scenario_cb = ttk.Combobox(main, textvariable=self.scenario_var, state="readonly", width=40)
        self.scenario_cb.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        self.scenario_cb.bind("<<ComboboxSelected>>", self.on_scenario_selected)

        tk.Label(main, text=lang.t("receive_Kit.movement_type","Movement Type:"), bg="#F0F4F8")\
            .grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.mode_cb = ttk.Combobox(main, textvariable=self.mode_var, state="readonly", width=40)
        self.mode_cb.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        self.mode_cb.bind("<<ComboboxSelected>>", self.on_mode_changed)

        # Kit / Kit number
        tk.Label(main, text=lang.t("receive_Kit.select_Kit","Select Kit:"), bg="#F0F4F8")\
            .grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.Kit_cb = ttk.Combobox(main, textvariable=self.Kit_var, state="disabled", width=40)
        self.Kit_cb.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.Kit_cb.bind("<<ComboboxSelected>>", self.on_Kit_selected)

        tk.Label(main, text=lang.t("receive_Kit.select_Kit_number","Select Kit Number:"), bg="#F0F4F8")\
            .grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.Kit_number_cb = ttk.Combobox(main, textvariable=self.Kit_number_var, state="disabled", width=20)
        self.Kit_number_cb.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.Kit_number_cb.bind("<<ComboboxSelected>>", self.on_Kit_number_selected)

        # Module / Module number
        tk.Label(main, text=lang.t("receive_Kit.select_module","Select Module:"), bg="#F0F4F8")\
            .grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.module_cb = ttk.Combobox(main, textvariable=self.module_var, state="disabled", width=40)
        self.module_cb.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.module_cb.bind("<<ComboboxSelected>>", self.on_module_selected)

        tk.Label(main, text=lang.t("receive_Kit.select_module_number","Select Module Number:"), bg="#F0F4F8")\
            .grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.module_number_cb = ttk.Combobox(main, textvariable=self.module_number_var, state="disabled", width=20)
        self.module_number_cb.grid(row=3, column=3, padx=5, pady=5, sticky="w")
        self.module_number_cb.bind("<<ComboboxSelected>>", self.on_module_number_selected)

        # Fixed OUT Type label
        type_frame = tk.Frame(main, bg="#F0F4F8")
        type_frame.grid(row=4, column=0, columnspan=4, pady=5, sticky="w")
        tk.Label(type_frame, text=lang.t("dispatch_Kit.out_type","OUT Type:"), bg="#F0F4F8")\
            .grid(row=0, column=0, padx=5, sticky="w")
        tk.Label(type_frame, textvariable=self.out_type_var,
                 bg="#E0E0E0", fg="#000", relief="sunken",
                 width=35, anchor="w").grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Search
        tk.Label(main, text=lang.t("receive_Kit.item","Kit/Module/Item:"), bg="#F0F4F8")\
            .grid(row=5, column=0, padx=5, pady=5, sticky="w")
        search_entry = tk.Entry(main, textvariable=self.search_var, width=40)
        search_entry.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        search_entry.bind("<KeyRelease>", self.search_items)

        tk.Button(main, text=lang.t("receive_Kit.clear_search","Clear Search"),
                  bg="#7F8C8D", fg="white",
                  command=self.clear_search).grid(row=5, column=2, padx=5, pady=5)

        self.search_listbox = tk.Listbox(main, height=5, width=60)
        self.search_listbox.grid(row=6, column=1, columnspan=3, padx=5, pady=5, sticky="we")

        # Tree definition
        cols = ("code","description","type","Kit","module",
                "current_stock","expiry_date","batch_no",
                "qty_to_issue","unique_id","line_id","qty_in_hidden")
        self.tree = ttk.Treeview(main, columns=cols, show="headings", height=18)

        headers = {
            "code":"Code","description":"Description","type":"Type","Kit":"Kit",
            "module":"Module","current_stock":"Current Stock","expiry_date":"Expiry Date",
            "batch_no":"Batch No","qty_to_issue":"Qty to Break",
            "unique_id":"Unique ID","line_id":"line_id (hidden)","qty_in_hidden":"qty_in (hidden)"
        }
        widths = {
            "code":150,"description":360,"type":110,"Kit":120,"module":120,
            "current_stock":110,"expiry_date":140,"batch_no":130,"qty_to_issue":140,
            "unique_id":1,"line_id":1,"qty_in_hidden":1
        }
        aligns = {
            "code":"w","description":"w","type":"w","Kit":"w","module":"w",
            "current_stock":"e","expiry_date":"w","batch_no":"w","qty_to_issue":"e",
            "unique_id":"w","line_id":"w","qty_in_hidden":"e"
        }
        for c in cols:
            self.tree.heading(c, text=headers[c])
            stretch = False if c in ("unique_id","line_id","qty_in_hidden") else True
            self.tree.column(c, width=widths[c], anchor=aligns[c], stretch=stretch, minwidth=1)

        # Tag styles
        self.tree.tag_configure("header_kit", background="#E3F6E1", font=("Helvetica",10,"bold"))
        self.tree.tag_configure("header_module", background="#E1ECFC", font=("Helvetica",10,"bold"))
        self.tree.tag_configure("kit_data", background="#C5EDC1")
        self.tree.tag_configure("module_data", background="#C9E2FA")
        self.tree.tag_configure("editable_row", foreground="#000000")
        self.tree.tag_configure("non_editable", foreground="#666666")
        self.tree.tag_configure("item_row", foreground="#222222")

        vsb = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        vsb.grid(row=7, column=4, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(main, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=8, column=0, columnspan=4, sticky="ew")
        self.tree.configure(xscrollcommand=hsb.set)
        self.tree.grid(row=7, column=0, columnspan=4, pady=10, sticky="nsew")
        main.grid_rowconfigure(7, weight=1)
        main.grid_columnconfigure(1, weight=1)

        # Bindings
        self.tree.bind("<Double-1>", self.start_edit)
        self.tree.bind("<KeyPress-Return>", self.start_edit)
        self.tree.bind("<KeyPress-Tab>", self.start_edit)
        self.tree.bind("<KeyPress-Up>", self.navigate_tree)
        self.tree.bind("<KeyPress-Down>", self.navigate_tree)

        btnf = tk.Frame(main, bg="#F0F4F8")
        btnf.grid(row=9, column=0, columnspan=4, pady=5)
        tk.Button(btnf, text=lang.t("receive_Kit.save","Save"),
                  bg="#27AE60", fg="white",
                  command=self.save_all,
                  state="normal" if self.role in ["admin","manager"] else "disabled").pack(side="left", padx=5)
        tk.Button(btnf, text=lang.t("receive_Kit.clear","Clear"),
                  bg="#7F8C8D", fg="white", command=self.clear_form).pack(side="left", padx=5)
        tk.Button(btnf, text=lang.t("receive_Kit.export","Export"),
                  bg="#2980B9", fg="white", command=self.export_data).pack(side="left", padx=5)

        tk.Label(main, textvariable=self.status_var, relief="sunken",
                 anchor="w", bg="#F0F4F8").grid(row=10, column=0, columnspan=4, sticky="ew")

        self.load_scenarios()

    # -------------------- Row building & sorting (treecode) --------------------
    def _build_with_headers(self, rows):
        def sort_key(it):
            return (
                it.get("Kit_number") or "",
                it.get("module_number") or "",
                it.get("treecode") or "ZZZ",
                it.get("code") or ""
            )
        ordered = sorted(rows, key=sort_key)
        result = []
        seen_kits = set()
        seen_modules = set()
        for it in ordered:
            kit_code = it.get("Kit") if it.get("Kit") and it.get("Kit") != "-----" else None
            module_code = it.get("module") if it.get("module") and it.get("module") != "-----" else None
            kit_number = it.get("Kit_number")
            module_number = it.get("module_number")
            if kit_code and kit_number and (kit_code, kit_number) not in seen_kits:
                result.append({
                    "is_header": True,
                    "type": "Kit",
                    "code": kit_code,
                    "description": get_item_description(kit_code),
                    "Kit": kit_number,
                    "module": "",
                    "current_stock": "",
                    "expiry_date": "",
                    "batch_no": "",
                    "unique_id": "",
                    "Kit_number": kit_number,
                    "module_number": None,
                    "line_id": "",
                    "treecode": ""
                })
                seen_kits.add((kit_code, kit_number))
            if module_code and module_number and (kit_code, module_code, module_number, kit_number) not in seen_modules:
                result.append({
                    "is_header": True,
                    "type": "Module",
                    "code": module_code,
                    "description": get_item_description(module_code),
                    "Kit": kit_number or "",
                    "module": module_number,
                    "current_stock": "",
                    "expiry_date": "",
                    "batch_no": "",
                    "unique_id": "",
                    "Kit_number": kit_number,
                    "module_number": module_number,
                    "line_id": "",
                    "treecode": ""
                })
                seen_modules.add((kit_code, module_code, module_number, kit_number))
            result.append(it)
        return result

    # -------------------- Mode rules / quantity derivation --------------------
    def get_mode_rules(self):
        mode = self.current_mode_key()
        rules = {
            "editable_types": set(),
            "derive_modules_from_Kit": False,
            "derive_items_from_modules": False
        }
        if mode == "dispatch_Kit":
            rules.update({
                "editable_types": {"Kit"},
                "derive_modules_from_Kit": True,
                "derive_items_from_modules": True
            })
        elif mode in ("issue_module_scenario","issue_module_Kit"):
            rules.update({
                "editable_types": {"Module"},
                "derive_items_from_modules": True
            })
        elif mode in ("issue_standalone","issue_items_module","issue_items_Kit"):
            rules.update({"editable_types":{"Item"}})
        return rules

    def initialize_quantities_and_highlight(self):
        rules = self.get_mode_rules()
        mode_key = self.current_mode_key()
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if meta.get("is_header"): continue
            vals = list(self.tree.item(iid, "values"))
            rtype = (vals[2] or "").lower()
            try:
                stock = int(vals[5]) if vals[5] else 0
            except:
                stock = 0
            if rtype == "kit":
                qty = 1 if mode_key == "dispatch_Kit" else 0
            elif rtype == "module":
                qty = 1 if ("module" in {t.lower() for t in rules["editable_types"]} and stock>0) else 0
            else:
                qty = 0
            vals[8] = str(qty)
            vals[11] = str(qty)
            self.tree.item(iid, values=vals)
        if rules.get("derive_modules_from_Kit"):
            self._derive_modules_from_Kits()
        if rules.get("derive_items_from_modules"):
            self._derive_items_from_modules()
        self._reapply_editable_icons(rules)

    def _derive_modules_from_Kits(self):
        kit_qty = {}
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid,{})
            if meta.get("is_header"): continue
            vals = self.tree.item(iid,"values")
            if (vals[2] or "").lower()=="kit":
                raw = vals[8]
                if raw.startswith("★"):
                    raw = raw[2:].strip()
                kit_qty[meta.get("Kit_number")] = int(raw) if raw.isdigit() else 0
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid,{})
            if meta.get("is_header"): continue
            vals = list(self.tree.item(iid,"values"))
            if (vals[2] or "").lower()=="module":
                base = kit_qty.get(meta.get("Kit_number"),0)
                try:
                    stock = int(vals[5]) if vals[5] else 0
                except: stock=0
                if base>stock: base=stock
                vals[8]=str(base)
                vals[11]=str(base)
                self.tree.item(iid, values=vals)

    def _derive_items_from_modules(self):
        module_map = {}
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid,{})
            if meta.get("is_header"): continue
            vals = self.tree.item(iid,"values")
            if (vals[2] or "").lower()=="module":
                raw = vals[8]
                if raw.startswith("★"):
                    raw = raw[2:].strip()
                module_map[meta.get("module_number")] = int(raw) if raw.isdigit() else 0
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid,{})
            if meta.get("is_header"): continue
            if meta.get("row_type") != "Item":
                continue
            vals = list(self.tree.item(iid,"values"))
            std_qty = meta.get("std_qty") or 1
            mod_qty = module_map.get(meta.get("module_number"),0)
            desired = std_qty * mod_qty
            try:
                stock = int(vals[5]) if vals[5] else 0
            except: stock=0
            if desired>stock: desired=stock
            vals[8]=str(desired)
            vals[11]=str(desired)
            self.tree.item(iid, values=vals)

    def _reapply_editable_icons(self, rules):
        editable_lower = {t.lower() for t in rules["editable_types"]}
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid,{})
            vals = list(self.tree.item(iid,"values"))
            rtype = (vals[2] or "").lower()
            tags=[]
            if meta.get("is_header"):
                if rtype=="kit":
                    tags.append("header_kit")
                elif rtype=="module":
                    tags.append("header_module")
            else:
                if rtype=="kit":
                    tags.append("kit_data")
                elif rtype=="module":
                    tags.append("module_data")
                else:
                    tags.append("item_row")
            if not meta.get("is_header") and rtype in editable_lower and meta.get("unique_id"):
                core = vals[8]
                raw = core[2:].strip() if core.startswith("★") else core.strip()
                if raw=="":
                    raw="0"
                vals[8] = f"★ {raw}"
                tags.append("editable_row")
                self.tree.item(iid, values=vals, tags=tuple(tags))
            else:
                if not meta.get("is_header"):
                    tags.append("non_editable")
                self.tree.item(iid, tags=tuple(tags))

# ============================= out_kit.py (Part 4/6) =============================
    # -------------------- Event handlers (selection) --------------------
    def on_Kit_selected(self, event=None):
        Kit_code = (self.Kit_var.get() or "").strip()
        if not Kit_code:
            self.Kit_number_cb.config(state="disabled")
            self.Kit_number_cb['values']=[]
            self.Kit_number_var.set("")
            return
        self.Kit_number_cb.config(state="readonly")
        self.Kit_number_cb['values'] = self.fetch_available_Kit_numbers(self.selected_scenario_id, Kit_code)
        self.Kit_number_var.set("")

    def on_Kit_number_selected(self, event=None):
        kit_no = (self.Kit_number_var.get() or "").strip()
        if not kit_no:
            self.clear_table_only()
            self.full_items=[]
            return
        kit_code = self.Kit_var.get() or None
        items = self.fetch_stock_data_for_Kit_number(self.selected_scenario_id, kit_no, kit_code)
        for it in items: it["row_type"]=it["type"]
        self.full_items = items[:]
        self.populate_rows(self.full_items,
                           f"Loaded {len(self.full_items)} lines for Kit number {kit_no}")

    def on_module_selected(self, event=None):
        module_code = (self.module_var.get() or "").strip()
        kit_code = (self.Kit_var.get() or "").strip() or None
        if not module_code:
            self.module_number_cb.config(state="disabled")
            self.module_number_cb['values']=[]
            self.module_number_var.set("")
            return
        self.module_number_cb.config(state="readonly")
        self.module_number_cb['values'] = self.fetch_module_numbers(self.selected_scenario_id, kit_code, module_code)
        self.module_number_var.set("")

    def on_module_number_selected(self, event=None):
        module_number = (self.module_number_var.get() or "").strip()
        mode_key = self.current_mode_key()
        if mode_key not in ("issue_items_module","issue_module_scenario","issue_module_Kit"):
            return
        if not module_number:
            self.clear_table_only()
            self.full_items=[]
            return
        kit_code = self.Kit_var.get() or None
        module_code = self.module_var.get() or None
        items = self.fetch_stock_data_for_module_number(self.selected_scenario_id, module_number, kit_code, module_code)
        for it in items: it["row_type"]=it["type"]
        self.full_items = items[:]
        self.populate_rows(self.full_items,
                           f"Loaded {len(self.full_items)} lines for module number {module_number}")

    def populate_standalone_items(self):
        if not self.selected_scenario_id: return
        items = self.fetch_standalone_stock_items(self.selected_scenario_id)
        for it in items: it["row_type"]=it["type"]
        self.full_items = items[:]
        self.populate_rows(self.full_items,
                           f"Loaded {len(self.full_items)} standalone rows")

    # -------------------- Table population --------------------
    def clear_table_only(self):
        if self.tree:
            self.tree.delete(*self.tree.get_children())
        self.row_data.clear()

    def populate_rows(self, items=None, status_msg=""):
        if items is None:
            items = self.full_items
        display_rows = self._build_with_headers(items)
        self.clear_table_only()
        for row in display_rows:
            if row.get("is_header"):
                values = (
                    row["code"], row["description"], row["type"],
                    row["Kit"], row["module"],
                    row["current_stock"], row["expiry_date"], row["batch_no"], "",
                    row.get("unique_id",""), row.get("line_id",""), "0"
                )
                iid = self.tree.insert("", "end", values=values)
                self.row_data[iid] = {
                    "is_header": True,
                    "row_type": row["type"],
                    "Kit_number": row.get("Kit_number"),
                    "module_number": row.get("module_number")
                }
            else:
                values = (
                    row["code"], row["description"], row["type"],
                    row["Kit"], row["module"],
                    row["current_stock"], row["expiry_date"], row["batch_no"], "",
                    row["unique_id"], row["line_id"], "0"
                )
                iid = self.tree.insert("", "end", values=values)
                self.row_data[iid] = {
                    "unique_id": row["unique_id"],
                    "Kit_number": row["Kit_number"],
                    "module_number": row["module_number"],
                    "current_stock": row["current_stock"],
                    "is_header": False,
                    "row_type": row["type"],
                    "std_qty": row.get("std_qty"),
                    "line_id": row.get("line_id")
                }
        self.initialize_quantities_and_highlight()
        if status_msg:
            self.status_var.set(status_msg)
        else:
            self.status_var.set(f"Showing {len(display_rows)} rows")

    # -------------------- Editing --------------------
    def start_edit(self, event):
        if event.type == tk.EventType.KeyPress:
            sel = self.tree.selection()
            if not sel: return
            self._start_edit_cell(sel[0], 8)
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id: return
        col_index = int(col_id.replace("#","")) - 1
        if col_index != 8: return
        self._start_edit_cell(row_id, col_index)

    def navigate_tree(self, event):
        if self.editing_cell: return
        rows = list(self.tree.get_children())
        if not rows: return
        sel = self.tree.selection()
        cur = sel[0] if sel else rows[0]
        idx = rows.index(cur)
        if event.keysym=="Up" and idx>0:
            self.tree.selection_set(rows[idx-1]); self.tree.focus(rows[idx-1])
        elif event.keysym=="Down" and idx < len(rows)-1:
            self.tree.selection_set(rows[idx+1]); self.tree.focus(rows[idx+1])

    def _start_edit_cell(self, row_id, col_index):
        if col_index != 8: return
        rules = self.get_mode_rules()
        editable_lower = {t.lower() for t in rules["editable_types"]}
        meta = self.row_data.get(row_id,{})
        if meta.get("is_header") or not meta.get("unique_id"): return
        vals = self.tree.item(row_id,"values")
        rtype = (vals[2] or "").lower()
        if rtype not in editable_lower: return
        bbox = self.tree.bbox(row_id, f"#{col_index+1}")
        if not bbox: return
        x,y,w,h = bbox
        raw_old = vals[8]
        old_clean = raw_old[2:].strip() if raw_old.startswith("★") else raw_old.strip()
        if self.editing_cell:
            try: self.editing_cell.destroy()
            except: pass
        entry = tk.Entry(self.tree, font=("Helvetica",10), background="#FFFBE0")
        entry.place(x=x,y=y,width=w,height=h)
        entry.insert(0, old_clean if old_clean else "")
        entry.focus()
        self.editing_cell = entry

        def save(_=None):
            val = entry.get().strip()
            try:
                stock = int(vals[5]) if vals[5] else 0
            except: stock=0
            if rtype in ("kit","module"):
                if val not in ("0","1"):
                    val = "1" if stock>0 else "0"
                if stock==0 and val=="1":
                    val="0"
            else:
                if not val.isdigit():
                    val="0"
                else:
                    iv=int(val)
                    if iv<0: iv=0
                    if iv>stock: iv=stock
                    val=str(iv)
            vals_list = list(vals)
            vals_list[8] = f"★ {val}"
            vals_list[11] = val
            self.tree.item(row_id, values=vals_list)
            entry.destroy()
            self.editing_cell=None
            if rtype=="kit" and rules.get("derive_modules_from_Kit"):
                self._derive_modules_from_Kits()
                if rules.get("derive_items_from_modules"):
                    self._derive_items_from_modules()
                self._reapply_editable_icons(rules)
            elif rtype=="module" and rules.get("derive_items_from_modules"):
                self._derive_items_from_modules()
                self._reapply_editable_icons(rules)

        entry.bind("<Return>", save)
        entry.bind("<Tab>", save)
        entry.bind("<FocusOut>", save)
        entry.bind("<Escape>", lambda e: (entry.destroy(), setattr(self,"editing_cell",None)))

    # -------------------- Search --------------------
    def search_items(self, event=None):
        q = (self.search_var.get() or "").strip().lower()
        if q == "":
            self.populate_rows(self.full_items,
                               f"Showing {len(self.full_items)} rows (reset)")
            return
        if len(q) < self.search_min_chars:
            self.status_var.set(f"Type at least {self.search_min_chars} chars to search...")
            return
        filtered=[]
        for it in self.full_items:
            if q in (it['code'] or "").lower() or q in (it['description'] or "").lower():
                filtered.append(it)
        self.populate_rows(filtered, f"Found {len(filtered)} matching rows")

# ============================= out_kit.py (Part 5/6) =============================
    # -------------------- Transaction helpers --------------------
    def _insert_transaction_out(self, cur, *, unique_id, code, description,
                                expiry_date, batch_number, scenario, kit_number,
                                module_number, qty_out, out_type,
                                ts_date, ts_time, movement_type, document_number):
        cur.execute("""
            INSERT INTO stock_transactions
            (Date, Time, unique_id, code, Description, Expiry_date, Batch_Number,
             Scenario, Kit, Module, Qty_IN, IN_Type, Qty_Out, Out_Type,
             Third_Party, End_User, Remarks, Movement_Type, document_number)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            ts_date, ts_time, unique_id, code, description,
            expiry_date, batch_number, scenario, kit_number, module_number,
            None, None, qty_out, out_type,
            None, None, None, movement_type, document_number
        ))

    def _insert_transaction_in_mirror(self, cur, *, unique_id, code, description,
                                      expiry_date, batch_number, scenario, kit_number,
                                      module_number, qty_in, out_type_as_in_type,
                                      ts_date, ts_time, movement_type, document_number):
        cur.execute("""
            INSERT INTO stock_transactions
            (Date, Time, unique_id, code, Description, Expiry_date, Batch_Number,
             Scenario, Kit, Module, Qty_IN, IN_Type, Qty_Out, Out_Type,
             Third_Party, End_User, Remarks, Movement_Type, document_number)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            ts_date, ts_time, unique_id, code, description,
            expiry_date, batch_number, scenario, kit_number, module_number,
            qty_in, out_type_as_in_type, None, None,
            None, None, None, movement_type, document_number
        ))

    # -------------------- Document Number --------------------
    def generate_document_number(self, out_type_text: str) -> str:
        project_name, project_code = fetch_project_details()
        project_code = (project_code or "PRJ").upper()
        now = datetime.now()
        prefix = f"{now.year:04d}/{now.month:02d}/{project_code}/BRK"
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
                    tail = r[0].rsplit('/',1)[-1]
                    if tail.isdigit():
                        serial = int(tail)+1
            finally:
                cur.close()
                conn.close()
        doc = f"{prefix}/{serial:04d}"
        self.current_document_number = doc
        return doc

    # -------------------- Save (Break logic) --------------------
    def save_all(self):
        if self.role not in ["admin","manager"]:
            custom_popup(self.parent, "Error",
                         "Only admin or manager roles can save changes.","error")
            return
        out_type = OUT_TYPE_FIXED
        rows=[]
        for iid in self.tree.get_children():
            vals = self.tree.item(iid,"values")
            if not vals: continue
            (code, desc, tfield, kit_col, module_col,
             current_stock, exp_date, batch_no, qty_to_issue,
             unique_id, line_id, qty_in_hidden) = vals
            meta = self.row_data.get(iid,{})
            if meta.get("is_header"): continue
            raw_q = qty_to_issue[2:].strip() if qty_to_issue.startswith("★") else qty_to_issue
            if not raw_q.isdigit(): continue
            q_out = int(raw_q)
            if q_out <= 0: continue
            try:
                stock_int = int(current_stock) if str(current_stock).isdigit() else 0
            except:
                stock_int=0
            if q_out > stock_int and stock_int>0:
                custom_popup(self.parent, "Error",
                             f"Cannot break more than stock for {code}","error")
                return
            q_in = int(qty_in_hidden) if qty_in_hidden.isdigit() else q_out
            rows.append({
                "code": code,
                "desc": desc,
                "type": tfield,
                "stock": stock_int,
                "qty_out": q_out,
                "qty_in": q_in,
                "exp_date": exp_date or None,
                "batch_no": batch_no or None,
                "unique_id": unique_id,
                "line_id": line_id if line_id else None,
                "kit_number": meta.get("Kit_number") or kit_col or None,
                "module_number": meta.get("module_number") or module_col or None
            })
        if not rows:
            custom_popup(self.parent,"Error","No quantities entered.","error")
            return

        scenario_name = self.scenario_map.get(self.selected_scenario_id,"")
        movement_label = self.mode_var.get()
        doc_number = self.generate_document_number(out_type)
        self.status_var.set(f"Processing... Document Number: {doc_number}")

        import time
        max_attempts=4
        for attempt in range(1,max_attempts+1):
            conn = connect_db()
            if conn is None:
                custom_popup(self.parent,"Error","Database connection failed","error")
                return
            try:
                conn.execute("PRAGMA busy_timeout=5000;")
                cur = conn.cursor()
                now_date = datetime.today().strftime('%Y-%m-%d')
                now_time = datetime.now().strftime('%H:%M:%S')
                for r in rows:
                    # Update qty_out
                    cur.execute("SELECT final_qty FROM stock_data WHERE unique_id=?",(r["unique_id"],))
                    row = cur.fetchone()
                    if not row or row[0] is None or row[0] < r["qty_out"]:
                        raise ValueError(f"Insufficient stock or concurrency issue for {r['code']}")
                    cur.execute("""
                        UPDATE stock_data
                           SET qty_out = qty_out + ?,
                               updated_at = ?
                         WHERE unique_id = ?
                           AND (qty_in - qty_out) >= ?
                    """,(r["qty_out"], f"{now_date} {now_time}", r["unique_id"], r["qty_out"]))
                    if cur.rowcount==0:
                        raise ValueError(f"Concurrent change or insufficient stock for {r['code']}")
                    # Mirror +qty_in using same line_id to offset final_qty (neutral net)
                    if r["line_id"] and r["qty_in"]>0:
                        cur.execute("""
                            UPDATE stock_data
                               SET qty_in = qty_in + ?,
                                   updated_at = ?
                             WHERE line_id = ?
                        """,(r["qty_in"], f"{now_date} {now_time}", r["line_id"]))
                    # Log OUT
                    self._insert_transaction_out(
                        cur,
                        unique_id=r["unique_id"],
                        code=r["code"],
                        description=r["desc"],
                        expiry_date=r["exp_date"],
                        batch_number=r["batch_no"],
                        scenario=scenario_name,
                        kit_number=r["kit_number"],
                        module_number=r["module_number"],
                        qty_out=r["qty_out"],
                        out_type=out_type,
                        ts_date=now_date,
                        ts_time=now_time,
                        movement_type=movement_label,
                        document_number=doc_number
                    )
                    # Log mirror IN
                    if r["qty_in"]>0:
                        self._insert_transaction_in_mirror(
                            cur,
                            unique_id=r["unique_id"],
                            code=r["code"],
                            description=r["desc"],
                            expiry_date=r["exp_date"],
                            batch_number=r["batch_no"],
                            scenario=scenario_name,
                            kit_number=r["kit_number"],
                            module_number=r["module_number"],
                            qty_in=r["qty_in"],
                            out_type_as_in_type=out_type,
                            ts_date=now_date,
                            ts_time=now_time,
                            movement_type=movement_label,
                            document_number=doc_number
                        )
                conn.commit()
                custom_popup(self.parent,"Success",
                             f"Break complete. Logged {len(rows)*2} transactions.","info")
                self.status_var.set(f"Break complete. Document Number: {doc_number}")
                if custom_askyesno(self.parent, "Confirm",
                                   "Export the break operation to Excel?") == "yes":
                    self.export_data(rows)
                self.clear_form()
                return
            except sqlite3.OperationalError as e:
                if "locked" in str(e).lower() and attempt < max_attempts:
                    try: conn.rollback()
                    except: pass
                    time.sleep(0.8*attempt)
                    continue
                else:
                    try: conn.rollback()
                    except: pass
                    custom_popup(self.parent,"Error", f"Break failed: {e}","error")
                    return
            except Exception as e:
                try: conn.rollback()
                except: pass
                custom_popup(self.parent,"Error", f"Break failed: {e}","error")
                return
            finally:
                try: cur.close()
                except: pass
                try: conn.close()
                except: pass
        custom_popup(self.parent,"Error","Break failed: database remained locked.","error")

    # -------------------- Utility / Clear / Export --------------------
    def clear_search(self):
        self.search_var.set("")
        self.populate_rows(self.full_items,
                           f"Showing {len(self.full_items)} rows (reset)")

    def clear_form(self):
        self.clear_table_only()
        self.scenario_var.set("")
        self.mode_var.set("")
        self.Kit_var.set("")
        self.Kit_number_var.set("")
        self.module_var.set("")
        self.module_number_var.set("")
        self.out_type_var.set(OUT_TYPE_FIXED)
        self.status_var.set(lang.t("break_kit.ready","Ready"))
        self.scenario_map = self.fetch_scenario_map()
        self.load_scenarios()

# ============================= out_kit.py (Part 6/6) =============================
    def export_data(self, rows_processed=None):
        if self.parent is None or not self.parent.winfo_exists():
            return
        try:
            export_rows = []
            for iid in self.tree.get_children():
                vals = self.tree.item(iid,"values")
                if not vals or len(vals) < 12: continue
                (code, desc, tfield, kit_col, module_col,
                 current_stock, exp_date, batch_no, qty_to_issue,
                 unique_id, line_id, qty_in_hidden) = vals
                raw_q = qty_to_issue[2:].strip() if qty_to_issue.startswith("★") else qty_to_issue
                qty_out = int(raw_q) if raw_q.isdigit() else 0
                qty_in = int(qty_in_hidden) if qty_in_hidden.isdigit() else qty_out
                meta = self.row_data.get(iid,{})
                export_rows.append({
                    "code": code,
                    "description": desc,
                    "type": tfield,
                    "kit_number": meta.get("Kit_number") or kit_col or "",
                    "module_number": meta.get("module_number") or module_col or "",
                    "current_stock": int(current_stock) if str(current_stock).isdigit() else 0,
                    "expiry_date": exp_date or "",
                    "batch_number": batch_no or "",
                    "qty_out": qty_out,
                    "qty_in": qty_in
                })
            if rows_processed:
                # Optionally refine; left as-is (already correct).
                pass
            if not any(r["qty_out"]>0 for r in export_rows):
                custom_popup(self.parent,"Error","No break quantities to export.","error")
                return

            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            movement_type_raw = self.mode_var.get() or "Break Kit"
            scenario_name = self.selected_scenario_name or "N/A"
            doc_number = getattr(self, "current_document_number", None)
            out_type_raw = OUT_TYPE_FIXED

            def sanitize(s):
                import re
                s = re.sub(r'[^A-Za-z0-9]+','_', s)
                s = re.sub(r'_+','_', s)
                return s.strip('_') or "Unknown"

            movement_slug = sanitize(movement_type_raw)
            default_dir = "D:/ISEPREP"
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)
            file_name = f"Break_{movement_slug}_{current_time.replace(':','-')}.xlsx"
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files","*.xlsx")],
                initialfile=file_name,
                initialdir=default_dir
            )
            if not path:
                self.status_var.set("Export cancelled")
                return

            wb = openpyxl.Workbook()
            ws = wb.active
            ws_title_base = "Break Kit-Module"
            ws.title = ws_title_base[:31]

            if doc_number:
                ws['A1'] = f"Date: {current_time}{' '*8}Document Number: {doc_number}"
            else:
                ws['A1'] = f"Date: {current_time}"
            ws['A1'].font = Font(name="Calibri", size=11)
            project_name, project_code = fetch_project_details()
            ws['A2'] = f"{ws_title_base} – Movement: {movement_type_raw}"
            ws['A2'].font = Font(name="Tahoma", size=14, bold=True)
            ws['A2'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A2:J2')
            ws['A3'] = f"{project_name} - {project_code}"
            ws['A3'].font = Font(name="Tahoma", size=14, bold=True)
            ws['A3'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A3:J3')
            ws['A4'] = f"OUT Type: {out_type_raw}"
            ws['A4'].font = Font(name="Tahoma", size=12, bold=True)
            ws['A4'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A4:J4')
            ws['A5'] = f"Scenario: {scenario_name}"
            ws['A5'].font = Font(name="Tahoma", size=12, bold=True)
            ws['A5'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A5:J5')
            ws.append([])

            headers = ["Code","Description","Type","Kit Number","Module Number",
                       "Current Stock","Expiry Date","Batch Number",
                       "Qty Broken (Out)","Qty In (Mirror)"]
            ws.append(headers)
            for c in range(1,len(headers)+1):
                ws.cell(row=7,column=c).font = Font(name="Tahoma", size=11, bold=True)

            from openpyxl.styles import PatternFill
            for row_idx, r in enumerate(export_rows, start=8):
                ws.append([
                    r["code"], r["description"], r["type"], r["kit_number"],
                    r["module_number"], r["current_stock"], r["expiry_date"],
                    r["batch_number"], r["qty_out"], r["qty_in"]
                ])
                rtype = (r["type"] or "").lower()
                for col in range(1,len(headers)+1):
                    cell = ws.cell(row=row_idx, column=col)
                    cell.font = Font(name="Calibri", size=11, bold=(rtype in ("kit","module")))
                    if rtype=="kit":
                        cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
                    elif rtype=="module":
                        cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

            # Autofit
            for col in ws.columns:
                max_len=0
                letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value:
                        ln = len(str(cell.value))
                        if ln>max_len: max_len=ln
                ws.column_dimensions[letter].width = min(max_len+2, 50)

            wb.save(path)
            custom_popup(self.parent,"Success",f"Exported to {path}","info")
            self.status_var.set(f"Export successful: {path}")
        except Exception as e:
            logging.error(f"[BREAK] Export failed: {e}")
            custom_popup(self.parent,"Error",f"Export failed: {e}","error")


# -------------------- Main Harness --------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Break Kit/Module")
    app = type("App", (), {})()
    app.role = "admin"
    StockOutKit(root, app, role="admin")
    root.geometry("1400x850")
    root.mainloop()

