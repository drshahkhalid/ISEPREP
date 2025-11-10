import tkinter as tk
from tkinter import ttk, filedialog
import sqlite3
import logging
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from datetime import datetime
from popup_utils import custom_popup, custom_askyesno
import os
from popup_utils import custom_popup, custom_askyesno, custom_dialog
from db import connect_db
from manage_items import get_item_description, detect_type
from language_manager import lang
from popup_utils import custom_popup

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

def fetch_end_users():
    conn = connect_db()
    if conn is None:
        logging.error("[DISPATCH] DB connection failed (fetch_end_users)")
        return []
    cur = conn.cursor()
    try:
        cur.execute("SELECT name FROM end_users ORDER BY name")
        return [r[0] for r in cur.fetchall()]
    except sqlite3.Error as e:
        logging.error(f"[DISPATCH] fetch_end_users error: {e}")
        return []
    finally:
        cur.close()
        conn.close()

def fetch_third_parties():
    conn = connect_db()
    if conn is None:
        logging.error("[DISPATCH] DB connection failed (fetch_third_parties)")
        return []
    cur = conn.cursor()
    try:
        cur.execute("SELECT name FROM third_parties ORDER BY name")
        return [r[0] for r in cur.fetchall()]
    except sqlite3.Error as e:
        logging.error(f"[DISPATCH] fetch_third_parties error: {e}")
        return []
    finally:
        cur.close()
        conn.close()

def fetch_project_details():
    conn = connect_db()
    if conn is None:
        logging.error("[DISPATCH] DB connection failed (fetch_project_details)")
        return "Unknown Project", "Unknown Code"
    cur = conn.cursor()
    try:
        cur.execute("SELECT project_name, project_code FROM project_details LIMIT 1")
        row = cur.fetchone()
        return (row[0] if row and row[0] else "Unknown Project",
                row[1] if row and row[1] else "Unknown Code")
    except sqlite3.Error as e:
        logging.error(f"[DISPATCH] fetch_project_details error: {e}")
        return "Unknown Project", "Unknown Code"
    finally:
        cur.close()
        conn.close()

def log_transaction(unique_id, code, description, expiry_date, batch_number,
                    scenario, Kit, module, qty_out, out_type,
                    third_party, end_user, remarks, movement_type):
    conn = connect_db()
    if conn is None:
        raise ValueError("DB connection failed")
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO stock_transactions
            (Date, Time, unique_id, code, Description, Expiry_date, Batch_Number,
             Scenario, Kit, Module, Qty_IN, IN_Type, Qty_Out, Out_Type,
             Third_Party, End_User, Remarks, Movement_Type)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            datetime.today().strftime('%Y-%m-%d'),
            datetime.now().strftime('%H:%M:%S'),
            unique_id, code, description, expiry_date, batch_number,
            scenario, Kit, module,
            None, None, qty_out, out_type,
            third_party, end_user, remarks, movement_type
        ))
        conn.commit()
    except sqlite3.Error as e:
        conn.rollback()
        logging.error(f"[DISPATCH] Transaction log error: {e}")
        raise
    finally:
        cur.close()
        conn.close()
def configure_db_pragmas():
    """
    Apply SQLite pragmas to reduce 'database is locked' errors.
    Call once at startup (e.g., in __init__ or before first write).
    """
    conn = connect_db()
    if conn is None:
        return
    try:
        cur = conn.cursor()
        # WAL allows readers during writes; adjust synchronous for speed/safety tradeoff.
        cur.execute("PRAGMA journal_mode=WAL;")
        cur.execute("PRAGMA busy_timeout=5000;")  # ms
        # cur.execute("PRAGMA synchronous=NORMAL;")  # uncomment if acceptable
        conn.commit()
    except Exception:
        pass
    finally:
        try:
            conn.close()
        except:
            pass

class StockDispatchKit(tk.Frame):
    """
    Dispatch (Issue Out) Module
    Modes:
      - dispatch_Kit
      - issue_standalone
      - issue_module_scenario
      - issue_module_Kit
      - issue_items_Kit
      - issue_items_module
    """

    def __init__(self, parent, app, role="supervisor"):
        super().__init__(parent)
        self.parent = parent
        self.app = app
        self.role = role.lower()
    # Apply DB pragmas (optional)
        try:
            configure_db_pragmas()
        except Exception:
            pass
        # Scenario state
        self.scenario_map = self.fetch_scenario_map()
        self.selected_scenario_id = None
        self.selected_scenario_name = None

        # UI variable references
        self.scenario_var = None
        self.scenario_cb = None
        self.mode_var = None
        self.mode_cb = None
        self.Kit_var = None
        self.Kit_cb = None
        self.Kit_number_var = None
        self.Kit_number_cb = None
        self.module_var = None
        self.module_cb = None
        self.module_number_var = None
        self.module_number_cb = None
        self.trans_type_var = None
        self.trans_type_cb = None
        self.end_user_var = None
        self.end_user_cb = None
        self.third_party_var = None
        self.third_party_cb = None
        self.remarks_entry = None
        self.search_var = None
        self.search_entry = None
        self.search_listbox = None
        self.tree = None
        self.status_var = None
        self.editing_cell = None

        # Data caches
        self.row_data = {}          # iid -> metadata
        self.full_items = []        # raw enriched rows (Kit/Module/Item without headers)
        self.search_min_chars = 2

        if self.parent and self.parent.winfo_exists():
            self.pack(fill="both", expand=True)
            self.after(50, self.render_ui)

    # ------------------------------------------------------------------
    # Index / Parsing / Enrichment
    # ------------------------------------------------------------------
    def ensure_item_index(self, scenario_id):
        if hasattr(self, "_item_index_cache") and self._item_index_cache.get("scenario_id") == scenario_id:
            return
        self._item_index_cache = {"scenario_id": scenario_id, "flat_map": {}, "triple_map": {}}

        conn = connect_db()
        if conn is None:
            logging.warning("[DISPATCH] Cannot build item index (no DB connection).")
            return
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT code, Kit, module, item, treecode, level
                  FROM Kit_items
                 WHERE scenario_id = ?
            """, (scenario_id,))
            for row in cur.fetchall():
                code = (row["code"] or "").strip()
                Kit = (row["Kit"] or "").strip()
                module = (row["module"] or "").strip()
                item = (row["item"] or "").strip()
                treecode = row["treecode"]
                level = (row["level"] or "").strip().lower()
                entry = {
                    "code": code,
                    "Kit": Kit,
                    "module": module,
                    "item": item,
                    "treecode": treecode,
                    "level": level
                }
                for tok in (code, Kit, module, item):
                    if tok and tok.upper() != "NONE":
                        self._item_index_cache["flat_map"].setdefault(tok, entry)
                self._item_index_cache["triple_map"][(Kit or None,
                                                      module or None,
                                                      item or None)] = entry
            logging.info(f"[DISPATCH] item_index built: {len(self._item_index_cache['flat_map'])} tokens")
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] ensure_item_index error: {e}")
        finally:
            cur.close()
            conn.close()

    @staticmethod
    def parse_unique_id(unique_id: str):
        """
        Expected pattern:
          scenario / Kit / module / item / std_qty / exp_date / Kit_number / module_number
        std_qty parsed from index 4 if present & >0 else 1.
        """
        Kit_code = module_code = item_code = None
        std_qty = 1
        if not unique_id:
            return {"Kit_code": None, "module_code": None, "item_code": None, "std_qty": 1}
        parts = unique_id.split("/")
        if len(parts) >= 2:
            Kit_code = parts[1] or None
        if len(parts) >= 3:
            module_code = parts[2] or None
        if len(parts) >= 4:
            item_code = parts[3] or None
        if len(parts) >= 5:
            try:
                pv = int(parts[4])
                if pv > 0:
                    std_qty = pv
            except:
                std_qty = 1
        return {"Kit_code": Kit_code, "module_code": module_code, "item_code": item_code, "std_qty": std_qty}

    def enrich_stock_row(self, scenario_id, unique_id, final_qty, exp_date,
                         Kit_number, module_number):
        """
        Returns a dict with a normalized 'type' field: one of 'Kit', 'Module', 'Item'.
        Forced hierarchy: if module/item segments exist they override a broader detect_type result.
        """
        self.ensure_item_index(scenario_id)
        parsed = self.parse_unique_id(unique_id)
        Kit_code = parsed["Kit_code"]
        module_code = parsed["module_code"]
        item_code = parsed["item_code"]
        std_qty = parsed["std_qty"]

        # Determine display code & forced_type from unique_id structure
        if item_code and item_code.upper() != "NONE":
            display_code = item_code
            forced_type = "Item"
        elif module_code and module_code.upper() != "NONE":
            display_code = module_code
            forced_type = "Module"
        else:
            display_code = Kit_code
            forced_type = "Kit"

        # Lookup treecode (unchanged)
        treecode = None
        triple_key = (
            Kit_code if Kit_code else None,
            module_code if module_code else None,
            item_code if item_code else None
        )
        idx = getattr(self, "_item_index_cache", {})
        triple_map = idx.get("triple_map", {})
        flat_map = idx.get("flat_map", {})
        entry = triple_map.get(triple_key) or flat_map.get(display_code or "")
        if entry:
            treecode = entry.get("treecode")

        description = get_item_description(display_code or "")
        detected = detect_type(display_code or "", description) or forced_type

        # Normalize detected to canonical Title case if valid
        detected_upper = detected.upper()
        valid_upper = {"KIT": "Kit", "MODULE": "Module", "ITEM": "Item"}
        if detected_upper in valid_upper:
            detected_norm = valid_upper[detected_upper]
        else:
            detected_norm = forced_type  # fallback

        # If forced_type narrows (Module/Item), honor it over detected
        if forced_type in ("Module", "Item"):
            final_type = forced_type
        else:
            # forced_type == 'Kit'; keep detected_norm only if it is not narrower
            # (If detection said Module/Item but unique_id lacks those segments, detection is wrong; keep Kit)
            final_type = "Kit" if detected_norm not in ("Kit",) else detected_norm

        return {
            "unique_id": unique_id,
            "code": display_code or "",
            "description": description,
            "type": final_type,    # 'Kit', 'Module', or 'Item'
            "Kit": Kit_code or "-----",
            "module": module_code or "-----",
            "current_stock": final_qty,
            "expiry_date": exp_date or "",
            "batch_no": "",
            "treecode": treecode,
            "Kit_number": Kit_number,
            "module_number": module_number,
            "std_qty": std_qty if final_type == "Item" else None
        }
    # ------------------------------------------------------------------
    # Scenario / Modes
    # ------------------------------------------------------------------
    def fetch_scenario_map(self):
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed in fetch_scenario_map")
            return {}
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("SELECT scenario_id, name FROM scenarios ORDER BY name")
            rows = cur.fetchall()
            mapping = {str(r['scenario_id']): r['name'] for r in rows}
            logging.info(f"[DISPATCH] Loaded {len(mapping)} scenarios")
            return mapping
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] Error loading scenarios: {e}")
            return {}
        finally:
            cur.close()
            conn.close()

    def build_mode_definitions(self):
        scenario = self.selected_scenario_name or ""
        self.mode_definitions = [
            ("dispatch_Kit",
             lang.t("dispatch_Kit.mode_dispatch_Kit", "Dispatch Kit")),
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
        values = [f"{sid} - {name}" for sid, name in self.scenario_map.items()]
        logging.info(f"[DISPATCH] Populating scenario combobox with {len(values)} entries")
        self.scenario_cb['values'] = values
        logging.info(f"[DISPATCH] scenario_cb values set: {self.scenario_cb['values']}")
        if values:
            self.scenario_cb.current(0)
            self.on_scenario_selected()
        else:
            if self.status_var:
                self.status_var.set("No scenarios found (check DB).")

    def on_scenario_selected(self, event=None):
        sel = self.scenario_var.get()
        if not sel:
            self.selected_scenario_id = None
            self.selected_scenario_name = None
            return
        self.selected_scenario_id = sel.split(" - ")[0]
        self.selected_scenario_name = sel.split(" - ", 1)[1] if " - " in sel else ""
        logging.info(f"[DISPATCH] Scenario selected: id={self.selected_scenario_id} name={self.selected_scenario_name}")
        self.build_mode_definitions()
        self.mode_cb['values'] = [lbl for _, lbl in self.mode_definitions]
        if self.mode_definitions:
            self.mode_var.set(self.mode_definitions[0][1])
        self.on_mode_changed()

    def on_mode_changed(self, event=None):
        """
        Reset UI and enable only the first-level selectors appropriate
        to the chosen mode. Kit numbers and module numbers remain disabled
        until their parent (Kit or module) is selected.
        """
        mode_key = self.current_mode_key()
        logging.info(f"[DISPATCH] Mode changed to: {mode_key}")

        # Reset dependent comboboxes
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

        if mode_key == "issue_standalone":
            self.populate_standalone_items()
            return

        if mode_key in ("dispatch_Kit", "issue_items_Kit", "issue_module_Kit"):
            self.Kit_cb.config(state="readonly")
            self.Kit_cb['values'] = self.fetch_Kits(self.selected_scenario_id)

        if mode_key in ("issue_items_module", "issue_module_Kit", "issue_module_scenario"):
            self.module_cb.config(state="readonly")
            self.module_cb['values'] = self.fetch_all_modules(self.selected_scenario_id)

    # ------------------------------------------------------------------
    # Structural Helpers
    # ------------------------------------------------------------------
    def fetch_Kits(self, scenario_id):
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed fetch_Kits")
            return []
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT code FROM Kit_items
                WHERE scenario_id = ? AND level='primary'
                ORDER BY code
            """, (scenario_id,))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] fetch_Kits error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_all_modules(self, scenario_id):
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed fetch_all_modules")
            return []
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT code FROM Kit_items
                WHERE scenario_id=? AND level='secondary'
                ORDER BY code
            """, (scenario_id,))
            return [r[0] for r in cur.fetchall()]
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] fetch_all_modules error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_available_Kit_numbers(self, scenario_id, Kit_code=None):
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed fetch_available_Kit_numbers")
            return []
        cur = conn.cursor()
        try:
            if Kit_code:
                cur.execute("""
                    SELECT DISTINCT Kit_number
                      FROM stock_data
                     WHERE Kit_number IS NOT NULL
                       AND Kit_number != 'None'
                       AND unique_id LIKE ?
                       AND unique_id LIKE ?
                     ORDER BY Kit_number
                """, (f"{scenario_id}/%", f"{scenario_id}/{Kit_code}/%"))
            else:
                cur.execute("""
                    SELECT DISTINCT Kit_number
                      FROM stock_data
                     WHERE Kit_number IS NOT NULL
                       AND Kit_number != 'None'
                       AND unique_id LIKE ?
                     ORDER BY Kit_number
                """, (f"{scenario_id}/%",))
            vals = [r[0] for r in cur.fetchall()]
            logging.info(f"[DISPATCH] Kit numbers loaded: {len(vals)}")
            return vals
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] fetch_available_Kit_numbers error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_module_numbers(self, scenario_id, Kit_code=None, module_code=None):
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed fetch_module_numbers")
            return []
        cur = conn.cursor()
        try:
            where = ["module_number IS NOT NULL", "module_number != 'None'", "unique_id LIKE ?"]
            params = [f"{scenario_id}/%"]
            if Kit_code:
                where.append("unique_id LIKE ?")
                params.append(f"{scenario_id}/{Kit_code}/%")
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
            vals = [r[0] for r in cur.fetchall()]
            logging.info(f"[DISPATCH] Module numbers loaded: {len(vals)}")
            return vals
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] fetch_module_numbers error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    # ------------------------------------------------------------------
    # Stock Fetching
    # ------------------------------------------------------------------
    def fetch_stock_data_for_Kit_number(self, scenario_id, Kit_number, Kit_code=None):
        logging.info(f"[DISPATCH] fetch_stock_data_for_Kit_number Kit_number={Kit_number} Kit_code={Kit_code}")
        self.ensure_item_index(scenario_id)
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed fetch_stock_data_for_Kit_number")
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT unique_id, final_qty, exp_date, Kit_number, module_number
                  FROM stock_data
                 WHERE Kit_number = ?
                   AND unique_id LIKE ?
                   AND final_qty > 0
            """, (Kit_number, f"{scenario_id}/%"))
            rows = cur.fetchall()
            items = []
            for r in rows:
                items.append(self.enrich_stock_row(
                    scenario_id=scenario_id,
                    unique_id=r["unique_id"],
                    final_qty=r["final_qty"],
                    exp_date=r["exp_date"],
                    Kit_number=r["Kit_number"],
                    module_number=r["module_number"]
                ))
            logging.info(f"[DISPATCH] Kit_number={Kit_number} -> {len(items)} rows enriched")
            self._debug_log_items(f"Kit_number={Kit_number}", items)
            return items
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] fetch_stock_data_for_Kit_number error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_stock_data_for_module_number(self, scenario_id, module_number, Kit_code=None, module_code=None):
        logging.info(f"[DISPATCH] fetch_stock_data_for_module_number module_number={module_number} Kit_code={Kit_code} module_code={module_code}")
        self.ensure_item_index(scenario_id)
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed fetch_stock_data_for_module_number")
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT unique_id, final_qty, exp_date, Kit_number, module_number
                  FROM stock_data
                 WHERE module_number = ?
                   AND unique_id LIKE ?
                   AND final_qty > 0
            """, (module_number, f"{scenario_id}/%"))
            rows = cur.fetchall()
            items = []
            for r in rows:
                items.append(self.enrich_stock_row(
                    scenario_id=scenario_id,
                    unique_id=r["unique_id"],
                    final_qty=r["final_qty"],
                    exp_date=r["exp_date"],
                    Kit_number=r["Kit_number"],
                    module_number=r["module_number"]
                ))
            logging.info(f"[DISPATCH] module_number={module_number} -> {len(items)} rows enriched")
            self._debug_log_items(f"module_number={module_number}", items)
            return items
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] fetch_stock_data_for_module_number error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    def fetch_standalone_stock_items(self, scenario_id):
        logging.info(f"[DISPATCH] Fetching standalone stock scenario={scenario_id}")
        self.ensure_item_index(scenario_id)
        conn = connect_db()
        if conn is None:
            logging.error("[DISPATCH] DB connection failed fetch_standalone_stock_items")
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT unique_id, final_qty, exp_date, Kit_number, module_number
                  FROM stock_data
                 WHERE final_qty > 0
                   AND (Kit_number IS NULL OR Kit_number='None')
                   AND (module_number IS NULL OR module_number='None')
                   AND unique_id LIKE ?
            """, (f"{scenario_id}/%",))
            rows = cur.fetchall()
            items = []
            for r in rows:
                items.append(self.enrich_stock_row(
                    scenario_id=scenario_id,
                    unique_id=r["unique_id"],
                    final_qty=r["final_qty"],
                    exp_date=r["exp_date"],
                    Kit_number=r["Kit_number"],
                    module_number=r["module_number"]
                ))
            logging.info(f"[DISPATCH] standalone -> {len(items)} raw rows")
            return items
        except sqlite3.Error as e:
            logging.error(f"[DISPATCH] fetch_standalone_stock_items error: {e}")
            return []
        finally:
            cur.close()
            conn.close()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def render_ui(self):
        logging.info("Rendering Dispatch UI")
        if not self.parent:
            return
        for w in self.parent.winfo_children():
            try:
                w.destroy()
            except:
                pass

        # Title
        title_frame = tk.Frame(self.parent, bg="#F0F4F8")
        title_frame.pack(fill="x")
        tk.Label(title_frame,
                 text=lang.t("dispatch_Kit.title", "Dispatch Kit-Module"),
                 font=("Helvetica", 20, "bold"),
                 bg="#F0F4F8").pack(pady=(10, 0))

        # Instruction (moved here so it does not interfere with grid rows)
        instruct_frame = tk.Frame(self.parent, bg="#FFF9C4", highlightbackground="#E0D890",
                                  highlightthickness=1, bd=0)
        instruct_frame.pack(fill="x", padx=10, pady=(6, 4))
        tk.Label(
    instruct_frame,
    text="Cells marked with ★ are editable. Enter quantities only there; other cells are automatic.For Kits and modules quantity entered can be either 1 or 0",
    fg="#444",
    bg="#FFF9C4",
    font=("Helvetica", 10, "italic")
).pack(padx=8, pady=4, anchor="w")

        # Main content frame (grid layout)
        main = tk.Frame(self.parent, bg="#F0F4F8")
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # Scenario
        tk.Label(main, text=lang.t("receive_Kit.scenario", "Scenario:"), bg="#F0F4F8")\
            .grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.scenario_var = tk.StringVar()
        self.scenario_cb = ttk.Combobox(main, textvariable=self.scenario_var, state="readonly", width=40)
        self.scenario_cb.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        self.scenario_cb.bind("<<ComboboxSelected>>", self.on_scenario_selected)

        # Mode
        tk.Label(main, text=lang.t("receive_Kit.movement_type", "Movement Type:"), bg="#F0F4F8")\
            .grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.mode_var = tk.StringVar()
        self.mode_cb = ttk.Combobox(main, textvariable=self.mode_var, state="readonly", width=40)
        self.mode_cb.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky="w")
        self.mode_cb.bind("<<ComboboxSelected>>", self.on_mode_changed)

        # Kit selection
        self.Kit_var = tk.StringVar()
        self.Kit_cb = ttk.Combobox(main, textvariable=self.Kit_var, state="disabled", width=40)
        tk.Label(main, text=lang.t("receive_Kit.select_Kit", "Select Kit:"), bg="#F0F4F8")\
            .grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.Kit_cb.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.Kit_cb.bind("<<ComboboxSelected>>", self.on_Kit_selected)

        self.Kit_number_var = tk.StringVar()
        self.Kit_number_cb = ttk.Combobox(main, textvariable=self.Kit_number_var, state="disabled", width=20)
        tk.Label(main, text=lang.t("receive_Kit.select_Kit_number", "Select Kit Number:"), bg="#F0F4F8")\
            .grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.Kit_number_cb.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.Kit_number_cb.bind("<<ComboboxSelected>>", self.on_Kit_number_selected)

        # Module selection
        self.module_var = tk.StringVar()
        self.module_cb = ttk.Combobox(main, textvariable=self.module_var, state="disabled", width=40)
        tk.Label(main, text=lang.t("receive_Kit.select_module", "Select Module:"), bg="#F0F4F8")\
            .grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.module_cb.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.module_cb.bind("<<ComboboxSelected>>", self.on_module_selected)

        self.module_number_var = tk.StringVar()
        self.module_number_cb = ttk.Combobox(main, textvariable=self.module_number_var, state="disabled", width=20)
        tk.Label(main, text=lang.t("receive_Kit.select_module_number", "Select Module Number:"), bg="#F0F4F8")\
            .grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.module_number_cb.grid(row=3, column=3, padx=5, pady=5, sticky="w")
        self.module_number_cb.bind("<<ComboboxSelected>>", self.on_module_number_selected)

        # Out Type & dependencies
        type_frame = tk.Frame(main, bg="#F0F4F8")
        type_frame.grid(row=4, column=0, columnspan=4, pady=5, sticky="w")
        tk.Label(type_frame, text=lang.t("dispatch_Kit.out_type", "OUT Type:"), bg="#F0F4F8")\
            .grid(row=0, column=0, padx=5, sticky="w")
        self.trans_type_var = tk.StringVar()
        self.trans_type_cb = ttk.Combobox(
            type_frame,
            textvariable=self.trans_type_var,
            values=[
                lang.t("dispatch_Kit.out_issue_end_user", "Issue to End User"),
                lang.t("dispatch_Kit.out_expired", "Expired Items"),
                lang.t("dispatch_Kit.out_damaged", "Damaged Items"),
                lang.t("dispatch_Kit.out_cold_chain", "Cold Chain Break"),
                lang.t("dispatch_Kit.out_batch_recall", "Batch Recall"),
                lang.t("dispatch_Kit.out_theft", "Theft"),
                lang.t("dispatch_Kit.out_other_losses", "Other Losses"),
                lang.t("dispatch_Kit.out_donation", "Out Donation"),
                lang.t("dispatch_Kit.out_loan", "Loan"),
                lang.t("dispatch_Kit.out_return_borrowing", "Return of Borrowing"),
                lang.t("dispatch_Kit.out_quarantine", "Quarantine")
            ],
            state="readonly", width=30
        )
        self.trans_type_cb.grid(row=0, column=1, padx=5, pady=5)
        self.trans_type_cb.bind("<<ComboboxSelected>>", self.on_out_type_selected)

        tk.Label(type_frame, text=lang.t("receive_Kit.end_user", "End User:"), bg="#F0F4F8")\
            .grid(row=0, column=2, padx=5, sticky="w")
        self.end_user_var = tk.StringVar()
        self.end_user_cb = ttk.Combobox(type_frame, textvariable=self.end_user_var, state="disabled", width=30)
        self.end_user_cb['values'] = fetch_end_users()
        self.end_user_cb.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(type_frame, text=lang.t("receive_Kit.third_party", "Third Party:"), bg="#F0F4F8")\
            .grid(row=0, column=4, padx=5, sticky="w")
        self.third_party_var = tk.StringVar()
        self.third_party_cb = ttk.Combobox(type_frame, textvariable=self.third_party_var, state="disabled", width=30)
        self.third_party_cb['values'] = fetch_third_parties()
        self.third_party_cb.grid(row=0, column=5, padx=5, pady=5)

        tk.Label(type_frame, text=lang.t("receive_Kit.remarks", "Remarks:"), bg="#F0F4F8")\
            .grid(row=0, column=6, padx=5, sticky="w")
        self.remarks_entry = tk.Entry(type_frame, width=40, state="disabled")
        self.remarks_entry.grid(row=0, column=7, padx=5, pady=5)

        # Search
        tk.Label(main, text=lang.t("receive_Kit.item", "Kit/Module/Item:"), bg="#F0F4F8")\
            .grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(main, textvariable=self.search_var, width=40)
        self.search_entry.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        self.search_entry.bind("<KeyRelease>", self.search_items)

        tk.Button(main, text=lang.t("receive_Kit.clear_search", "Clear Search"),
                  bg="#7F8C8D", fg="white", command=self.clear_search)\
            .grid(row=5, column=2, padx=5, pady=5)

        self.search_listbox = tk.Listbox(main, height=5, width=60)
        self.search_listbox.grid(row=6, column=1, columnspan=3, padx=5, pady=5, sticky="we")

        cols = ("code", "description", "type", "Kit", "module",
                "current_stock", "expiry_date", "batch_no", "qty_to_issue", "unique_id")
        self.tree = ttk.Treeview(main, columns=cols, show="headings", height=18)

        headers = {
            "code": lang.t("dispatch_Kit.code", "Code"),
            "description": lang.t("dispatch_Kit.description", "Description"),
            "type": lang.t("dispatch_Kit.type", "Type"),
            "Kit": lang.t("dispatch_Kit.Kit", "Kit"),
            "module": lang.t("dispatch_Kit.module", "Module"),
            "current_stock": lang.t("dispatch_Kit.current_stock", "Current Stock"),
            "expiry_date": lang.t("dispatch_Kit.expiry_date", "Expiry Date"),
            "batch_no": lang.t("dispatch_Kit.batch_no", "Batch Number"),
            "qty_to_issue": lang.t("dispatch_Kit.qty_to_issue", "Quantity to Issue"),
            "unique_id": "Unique ID"
        }
        widths = {
            "code": 160, "description": 380, "type": 120, "Kit": 120, "module": 120,
            "current_stock": 110, "expiry_date": 150, "batch_no": 140, "qty_to_issue": 140,
            "unique_id": 0
        }
        aligns = {
            "code": "w", "description": "w", "type": "w", "Kit": "w", "module": "w",
            "current_stock": "e", "expiry_date": "w", "batch_no": "w", "qty_to_issue": "e",
            "unique_id": "w"
        }
        for c in cols:
            self.tree.heading(c, text=headers[c])
            self.tree.column(
                c,
                width=widths[c],
                anchor=aligns[c],
                stretch=(False if c == "unique_id" else True),
                minwidth=0 if c == "unique_id" else widths[c]
            )

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
        tk.Button(btnf, text=lang.t("receive_Kit.save", "Save"),
                  bg="#27AE60", fg="white",
                  command=self.save_all,
                  state="normal" if self.role in ["admin", "manager"] else "disabled").pack(side="left", padx=5)
        tk.Button(btnf, text=lang.t("receive_Kit.clear", "Clear"),
                  bg="#7F8C8D", fg="white", command=self.clear_form).pack(side="left", padx=5)
        tk.Button(btnf, text=lang.t("receive_Kit.export", "Export"),
                  bg="#2980B9", fg="white", command=self.export_data).pack(side="left", padx=5)

        self.status_var = tk.StringVar(value=lang.t("receive_Kit.ready", "Ready"))
        tk.Label(main, textvariable=self.status_var, relief="sunken",
                 anchor="w", bg="#F0F4F8").grid(row=10, column=0, columnspan=4, sticky="ew")

        self.load_scenarios()

    # ------------------------------------------------------------------
    # Headers + Display Assembly
    # ------------------------------------------------------------------
    def _build_with_headers(self, rows):
        def sort_key(it):
            return (
                it.get("Kit_number") or "",
                it.get("module_number") or "",
                it.get("treecode") or "",
                it.get("code") or ""
            )
        ordered = sorted(rows, key=sort_key)
        result = []
        seen_Kit = set()
        seen_module = set()
        for it in ordered:
            Kit_code = it.get("Kit") if it.get("Kit") and it.get("Kit") != "-----" else None
            module_code = it.get("module") if it.get("module") and it.get("module") != "-----" else None
            Kit_number = it.get("Kit_number")
            module_number = it.get("module_number")

            if Kit_code and Kit_number and (Kit_code, Kit_number) not in seen_Kit:
                result.append({
                    "is_header": True,
                    "header_level": "Kit",
                    "code": Kit_code,
                    "description": get_item_description(Kit_code),
                    "type": "Kit",
                    "Kit": Kit_number,
                    "module": "",
                    "current_stock": "",
                    "expiry_date": "",
                    "batch_no": "",
                    "unique_id": "",
                    "Kit_number": Kit_number,
                    "module_number": None,
                    "treecode": None,
                    "std_qty": None
                })
                seen_Kit.add((Kit_code, Kit_number))

            if module_code and module_number and (Kit_code, module_code, module_number, Kit_number) not in seen_module:
                result.append({
                    "is_header": True,
                    "header_level": "module",
                    "code": module_code,
                    "description": get_item_description(module_code),
                    "type": "Module",
                    "Kit": Kit_number or "",
                    "module": module_number,
                    "current_stock": "",
                    "expiry_date": "",
                    "batch_no": "",
                    "unique_id": "",
                    "Kit_number": Kit_number,
                    "module_number": module_number,
                    "treecode": None,
                    "std_qty": None
                })
                seen_module.add((Kit_code, module_code, module_number, Kit_number))

            result.append(it)
        return result

    # ------------------------------------------------------------------
    # Mode Rules & Quantity Logic
    # ------------------------------------------------------------------
    def get_mode_rules(self):
        """
        Returns mode rules. Type set kept in original mixed case for display,
        but downstream comparisons are case-insensitive.
        """
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
        elif mode in ("issue_module_scenario", "issue_module_Kit"):
            rules.update({
                "editable_types": {"Module"},
                "derive_items_from_modules": True
            })
        elif mode in ("issue_standalone", "issue_items_module", "issue_items_Kit"):
            rules.update({
                "editable_types": {"Item"}
            })
        return rules

    def initialize_quantities_and_highlight(self):
        """
        Force Kit data rows to default to 1 in dispatch_Kit mode (even if stock 0).
        Case-insensitive row type checks; highlight Kit & Module rows.
        """
        rules = self.get_mode_rules()
        mode_key = self.current_mode_key()

        # 1. Base init
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if meta.get("is_header"):
                continue
            vals = list(self.tree.item(iid, "values"))
            row_type_lower = (vals[2] or "").lower()

            try:
                stock = int(vals[5]) if vals[5] else 0
            except Exception:
                stock = 0

            if row_type_lower == "kit":
                if mode_key == "dispatch_Kit":
                    qty = 1
                else:
                    qty = 1 if ("kit" in {t.lower() for t in rules["editable_types"]} and stock > 0) else 0
            elif row_type_lower == "module":
                qty = 1 if ("module" in {t.lower() for t in rules["editable_types"]} and stock > 0) else 0
            elif row_type_lower == "item":
                qty = 0
            else:
                qty = 0

            vals[8] = str(qty)
            self.tree.item(iid, values=vals)

        # 2. Derivations
        if rules.get("derive_modules_from_Kit") and hasattr(self, "_derive_modules_from_Kits"):
            self._derive_modules_from_Kits()
        if rules.get("derive_items_from_modules"):
            self._derive_items_from_modules()

        # 3. Header highlight
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if not meta.get("is_header"):
                continue
            vals = self.tree.item(iid, "values")
            rt = (vals[2] or "").lower()
            if rt == "kit":
                self.tree.item(iid, tags=("Kit_header", "Kit_module_highlight"))
            elif rt == "module":
                self.tree.item(iid, tags=("module_header", "Kit_module_highlight"))

        # 4. Mark editable
        editable_types_lower = {t.lower() for t in rules["editable_types"]}
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if meta.get("is_header"):
                continue
            vals = list(self.tree.item(iid, "values"))
            rt_low = (vals[2] or "").lower()
            tags = []
            if rt_low in ("kit", "module"):
                tags.append("Kit_module_highlight")
            is_editable = (rt_low in editable_types_lower and meta.get("unique_id"))
            if is_editable:
                if not vals[8].startswith("★"):
                    vals[8] = f"★ {vals[8]}"
                tags.append("editable_row")
                self.tree.item(iid, values=vals, tags=tuple(tags))
            else:
                tags.append("non_editable")
                self.tree.item(iid, tags=tuple(tags))

    def _derive_modules_from_Kits(self):
        """
        Derive Module quantities from Kit quantities (case-insensitive).
        """
        kit_quantities = {}
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if meta.get("is_header"):
                continue
            vals = self.tree.item(iid, "values")
            if (vals[2] or "").lower() == "kit":
                raw = vals[8]
                if raw.startswith("★"):
                    raw = raw[2:].strip()
                kit_qty = int(raw) if raw.isdigit() else 0
                kit_quantities[meta.get("Kit_number")] = kit_qty

        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if meta.get("is_header"):
                continue
            vals = list(self.tree.item(iid, "values"))
            if (vals[2] or "").lower() == "module":
                base_qty = kit_quantities.get(meta.get("Kit_number"), 0)
                try:
                    stock = int(vals[5]) if vals[5] else 0
                except Exception:
                    stock = 0
                if base_qty > stock:
                    base_qty = stock
                vals[8] = str(base_qty)
                self.tree.item(iid, values=vals)

    def _reapply_editable_icons(self, rules):
        """
        Reapply star and highlight after propagation; case-insensitive.
        Ensures only one authoritative definition of this method.
        """
        editable_lower = {t.lower() for t in rules["editable_types"]}

        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            vals = list(self.tree.item(iid, "values"))
            row_type = vals[2] or ""
            rt_low = row_type.lower()

            if meta.get("is_header"):
                if rt_low in ("kit", "module"):
                    base_tag = "Kit_header" if rt_low == "kit" else "module_header"
                    self.tree.item(iid, tags=(base_tag, "Kit_module_highlight"))
                continue

            tags = []
            if rt_low in ("kit", "module"):
                tags.append("Kit_module_highlight")

            if rt_low in editable_lower and meta.get("unique_id"):
                cell_val = vals[8]
                core = cell_val[2:].strip() if cell_val.startswith("★") else cell_val.strip()
                if core == "":
                    core = "0"
                vals[8] = f"★ {core}"
                tags.append("editable_row")
                self.tree.item(iid, values=vals, tags=tuple(tags))
            else:
                tags.append("non_editable")
                self.tree.item(iid, tags=tuple(tags))

    def _derive_items_from_modules(self):
        """
        For each Module row, set Item rows quantity =
        std_qty * module_qty (capped by stock). Strip any leading star.
        """
        module_qty_map = {}
        # Collect module quantities
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if meta.get("is_header"):
                continue
            vals = self.tree.item(iid, "values")
            if vals[2] == "Module":
                raw = vals[8]
                if raw.startswith("★"):
                    raw = raw[2:].strip()
                mqty = int(raw) if raw.isdigit() else 0
                module_qty_map[meta.get("module_number")] = mqty

        # Apply to items
        for iid in self.tree.get_children():
            meta = self.row_data.get(iid, {})
            if meta.get("is_header"):
                continue
            if meta.get("row_type") != "Item":
                continue
            mnum = meta.get("module_number")
            if not mnum:
                continue
            mod_qty = module_qty_map.get(mnum, 0)
            std_qty = meta.get("std_qty") or 1
            desired = std_qty * mod_qty
            vals = list(self.tree.item(iid, "values"))
            try:
                stock = int(vals[5]) if vals[5] else 0
            except Exception:
                stock = 0
            if desired > stock:
                desired = stock
            vals[8] = str(desired)
            self.tree.item(iid, values=vals)

    # ------------------------------------------------------------------
    # Out Type Dependents
    # ------------------------------------------------------------------
    def on_out_type_selected(self, event=None):
        out_type = self.trans_type_var.get()
        third_party_required = {
            lang.t("dispatch_Kit.out_donation", "Out Donation"),
            lang.t("dispatch_Kit.out_loan", "Loan"),
            lang.t("dispatch_Kit.out_return_borrowing", "Return of Borrowing")
        }
        end_user_required = {
            lang.t("dispatch_Kit.out_issue_end_user", "Issue to End User")
        }
        remarks_required = {
            lang.t("dispatch_Kit.out_expired", "Expired Items"),
            lang.t("dispatch_Kit.out_damaged", "Damaged Items"),
            lang.t("dispatch_Kit.out_cold_chain", "Cold Chain Break"),
            lang.t("dispatch_Kit.out_batch_recall", "Batch Recall"),
            lang.t("dispatch_Kit.out_theft", "Theft"),
            lang.t("dispatch_Kit.out_other_losses", "Other Losses"),
            lang.t("dispatch_Kit.out_quarantine", "Quarantine")
        }

        self.end_user_cb.config(state="disabled")
        self.third_party_cb.config(state="disabled")
        self.remarks_entry.config(state="disabled")

        if out_type not in end_user_required:
            self.end_user_var.set("")
        if out_type not in third_party_required:
            self.third_party_var.set("")
        if out_type not in remarks_required:
            self.remarks_entry.delete(0, tk.END)

        if out_type in end_user_required:
            self.end_user_cb.config(state="readonly")
        if out_type in third_party_required:
            self.third_party_cb.config(state="readonly")
        if out_type in remarks_required:
            self.remarks_entry.config(state="normal")

    # ------------------------------------------------------------------
    # Event Handlers / Loading
    # ------------------------------------------------------------------

    def on_Kit_selected(self, event=None):
        """
        When a Kit code is selected, enable and populate the Kit number list.
        If cleared, disable the Kit number field.
        """
        Kit_code = (self.Kit_var.get() or "").strip()

        if not Kit_code:
            self.Kit_number_cb.config(state="disabled")
            self.Kit_number_cb['values'] = []
            self.Kit_number_var.set("")
            return

        self.Kit_number_cb.config(state="readonly")
        self.Kit_number_cb['values'] = self.fetch_available_Kit_numbers(
            self.selected_scenario_id,
            Kit_code
        )
        self.Kit_number_var.set("")

    def on_Kit_number_selected(self, event=None):
        Kit_number = (self.Kit_number_var.get() or "").strip()
        if not Kit_number:
            self.clear_table_only()
            self.full_items = []
            return
        Kit_code = self.Kit_var.get() or None
        items = self.fetch_stock_data_for_Kit_number(self.selected_scenario_id, Kit_number, Kit_code)
        self.full_items = items[:]
        self.populate_rows(self.full_items,
                           f"Loaded {len(self.full_items)} stock rows for Kit number {Kit_number}")

    def on_module_selected(self, event=None):
        """
        When a module code is selected, enable and populate the module number list.
        If cleared, disable the module number field.
        """
        module_code = (self.module_var.get() or "").strip()
        Kit_code = (self.Kit_var.get() or "").strip() or None

        if not module_code:
            self.module_number_cb.config(state="disabled")
            self.module_number_cb['values'] = []
            self.module_number_var.set("")
            return

        self.module_number_cb.config(state="readonly")
        self.module_number_cb['values'] = self.fetch_module_numbers(
            self.selected_scenario_id,
            Kit_code,
            module_code
        )
        self.module_number_var.set("")

    def on_module_number_selected(self, event=None):
        module_number = (self.module_number_var.get() or "").strip()
        mode_key = self.current_mode_key()
        if mode_key not in ("issue_items_module", "issue_module_scenario", "issue_module_Kit"):
            return
        if not module_number:
            self.clear_table_only()
            self.full_items = []
            return
        Kit_code = self.Kit_var.get() or None
        module_code = self.module_var.get() or None
        items = self.fetch_stock_data_for_module_number(
            self.selected_scenario_id, module_number, Kit_code, module_code
        )
        self.full_items = items[:]
        self.populate_rows(self.full_items,
                           f"Loaded {len(self.full_items)} stock rows for module number {module_number}")

    def populate_standalone_items(self):
        if not self.selected_scenario_id:
            return
        items = self.fetch_standalone_stock_items(self.selected_scenario_id)
        self.full_items = items[:]
        self.populate_rows(self.full_items,
                           lang.t("dispatch_Kit.loaded_standalone",
                                  f"Loaded {len(self.full_items)} standalone rows"))

    # ------------------------------------------------------------------
    # Table Helpers
    # ------------------------------------------------------------------
    def clear_table_only(self):
        if self.tree:
            self.tree.delete(*self.tree.get_children())
        self.row_data.clear()

    def populate_rows(self, items=None, status_msg=""):
        """
        Populate the tree, insert headers + data rows, configure row tags.
        Now: ALL Kit & Module rows (headers + data) get yellow highlight.
        """
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
                    row["unique_id"]
                )
                iid = self.tree.insert("", "end", values=values)
                self.row_data[iid] = {
                    "is_header": True,
                    "row_type": row["type"],
                    "Kit_number": row["Kit_number"],
                    "module_number": row["module_number"]
                }
            else:
                values = (
                    row["code"], row["description"], row["type"],
                    row["Kit"], row["module"],
                    row["current_stock"], row["expiry_date"], row["batch_no"], "",
                    row["unique_id"]
                )
                iid = self.tree.insert("", "end", values=values)
                self.row_data[iid] = {
                    "unique_id": row["unique_id"],
                    "Kit_number": row["Kit_number"],
                    "module_number": row["module_number"],
                    "current_stock": row["current_stock"],
                    "is_header": False,
                    "row_type": row["type"],
                    "std_qty": row.get("std_qty")
                }

        # Style tags
        # Header tags keep bold font; background will be overridden by Kit_module_highlight
        self.tree.tag_configure("Kit_header", font=("Helvetica", 10, "bold"))
        self.tree.tag_configure("module_header", font=("Helvetica", 10, "bold"))
        # Unified yellow highlight for any Kit / Module (header or data)
        self.tree.tag_configure("Kit_module_highlight", background="#FFF9C4")
        # Editable just sets foreground (no background to avoid overriding highlight)
        self.tree.tag_configure("editable_row", foreground="#000000")
        # Non-editable rows (items or disallowed types) gray text
        self.tree.tag_configure("non_editable", foreground="#666666")

        # Initialize quantities & mark editable/non-editable + highlight
        self.initialize_quantities_and_highlight()

        if status_msg:
            self.status_var.set(status_msg)
        else:
            self.status_var.set(
                lang.t("dispatch_Kit.showing_rows",
                       f"Showing {len(display_rows)} rows (incl. headers)")
            )

        # Header tag styles
        self.tree.tag_configure("Kit_header", background="#E3F6E1", font=("Helvetica", 10, "bold"))
        self.tree.tag_configure("module_header", background="#E1ECFC", font=("Helvetica", 10, "bold"))
        # Editable cell tag (light yellow)
        self.tree.tag_configure("editable_cell", background="#FFF9C4")

        # Initialize + propagate + highlight editable
        self.initialize_quantities_and_highlight()

        if status_msg:
            self.status_var.set(status_msg)
        else:
            self.status_var.set(lang.t("dispatch_Kit.showing_rows",
                                       f"Showing {len(display_rows)} rows (incl. headers)"))

    def get_selected_unique_ids(self):
        uids = []
        for iid in self.tree.selection():
            hidden_val = self.tree.set(iid, "unique_id")
            if hidden_val:
                uids.append(hidden_val)
            else:
                meta = self.row_data.get(iid, {})
                if meta.get("unique_id"):
                    uids.append(meta["unique_id"])
        return uids

    def _debug_log_items(self, label, items, limit=12):
        logging.info(f"[DISPATCH][DEBUG] {label}: total={len(items)}")
        for i, it in enumerate(items[:limit]):
            logging.info(
                f"[DISPATCH][DEBUG] {label} row{i}: "
                f"unique_id={it.get('unique_id')} code={it.get('code')} "
                f"type={it.get('type')} Kit={it.get('Kit')} module={it.get('module')} "
                f"Kit_no={it.get('Kit_number')} module_no={it.get('module_number')} std_qty={it.get('std_qty')}"
            )

    # ------------------------------------------------------------------
    # Editing
    # ------------------------------------------------------------------
    def _flatten_rows(self):
        out = []
        def dive(iids):
            for r in iids:
                out.append(r)
                dive(self.tree.get_children(r))
        dive(self.tree.get_children())
        return out

    def navigate_tree(self, event):
        if self.editing_cell:
            return
        rows = self._flatten_rows()
        if not rows:
            return
        sel = self.tree.selection()
        if not sel:
            self.tree.selection_set(rows[0])
            self.tree.focus(rows[0])
            self.start_edit_cell(rows[0], 8)
            return
        cur = sel[0]
        idx = rows.index(cur)
        if event.keysym == "Up" and idx > 0:
            self.tree.selection_set(rows[idx - 1])
            self.tree.focus(rows[idx - 1])
            self.start_edit_cell(rows[idx - 1], 8)
        elif event.keysym == "Down" and idx < len(rows) - 1:
            self.tree.selection_set(rows[idx + 1])
            self.tree.focus(rows[idx + 1])
            self.start_edit_cell(rows[idx + 1], 8)
        elif event.keysym in ("Return", "Tab"):
            self.start_edit_cell(cur, 8)
    def start_edit(self, event):
        """
        Wrapper bound to double-click / Return / Tab.
        Decides the target row & column, then calls start_edit_cell.
        Only qty_to_issue (column index 8) is allowed, further filtered
        inside start_edit_cell based on mode rules and editable_types.
        """
        # Keyboard trigger: edit currently selected row’s qty cell.
        if event.type == tk.EventType.KeyPress:
            sel = self.tree.selection()
            if not sel:
                return
            self.start_edit_cell(sel[0], 8)
            return

        # Mouse double-click: ensure it's a cell in qty_to_issue column
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)  # e.g. "#9" for qty_to_issue
        if not row_id or not col_id:
            return
        col_index = int(col_id.replace("#", "")) - 1
        if col_index != 8:
            return
        self.start_edit_cell(row_id, 8)

    def start_edit_cell(self, row_id, col_index):
        """
        Edit Quantity to Issue (column 8).
        - Kit & Module: only 0 or 1
        - Item: non-negative integer, capped by stock
        """
        if col_index != 8:
            return

        rules = self.get_mode_rules()
        editable_lower = {t.lower() for t in rules["editable_types"]}
        meta = self.row_data.get(row_id, {})
        if meta.get("is_header") or not meta.get("unique_id"):
            return

        vals = self.tree.item(row_id, "values")
        rt_low = (vals[2] or "").lower()
        if rt_low not in editable_lower:
            return

        bbox = self.tree.bbox(row_id, f"#{col_index + 1}")
        if not bbox:
            return

        x, y, w, h = bbox
        raw_old = self.tree.set(row_id, "qty_to_issue")
        old_clean = raw_old[2:].strip() if raw_old.startswith("★") else raw_old.strip()

        if self.editing_cell:
            try:
                self.editing_cell.destroy()
            except Exception:
                pass

        entry = tk.Entry(self.tree, font=("Helvetica", 10), background="#FFFBE0")
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, old_clean if old_clean else "")
        entry.focus()
        self.editing_cell = entry

        def set_status(msg):
            if self.status_var:
                self.status_var.set(msg)

        def save(_=None):
            val = entry.get().strip()
            try:
                stock = int(vals[5]) if vals[5] else 0
            except Exception:
                stock = 0

            if rt_low in ("kit", "module"):
                if val not in ("0", "1"):
                    set_status("Only 0 or 1 allowed – auto-corrected.")
                    val = old_clean if old_clean in ("0", "1") else ("1" if stock > 0 else "0")
                if stock == 0 and val == "1":
                    val = "0"
            else:  # item
                if not val.isdigit():
                    set_status("Invalid number – set to 0.")
                    val = "0"
                else:
                    iv = int(val)
                    if iv < 0:
                        iv = 0
                        set_status("Negative not allowed – set to 0.")
                    if iv > stock:
                        iv = stock
                        set_status("Exceeded stock – capped.")
                    val = str(iv)

            self.tree.set(row_id, "qty_to_issue", f"★ {val}")
            entry.destroy()
            self.editing_cell = None

            # Propagation
            if rt_low == "kit" and rules.get("derive_modules_from_Kit") and hasattr(self, "_derive_modules_from_Kits"):
                self._derive_modules_from_Kits()
                if rules.get("derive_items_from_modules"):
                    self._derive_items_from_modules()
                self._reapply_editable_icons(rules)
            elif rt_low == "module" and rules.get("derive_items_from_modules"):
                self._derive_items_from_modules()
                self._reapply_editable_icons(rules)

        entry.bind("<Return>", save)
        entry.bind("<Tab>", save)
        entry.bind("<FocusOut>", save)
        entry.bind("<Escape>", lambda e: (entry.destroy(), setattr(self, "editing_cell", None)))

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------
    def search_items(self, event=None):
        query_raw = (self.search_var.get() or "").strip()
        query = query_raw.lower()
        if query == "":
            count = len(self.full_items)
            self.populate_rows(self.full_items,
                               lang.t("dispatch_Kit.showing_rows_reset",
                                      f"Showing {count} rows (reset)"))
            return
        if len(query) < self.search_min_chars:
            self.status_var.set(
                lang.t("dispatch_Kit.search_min_chars",
                       f"Type at least {self.search_min_chars} characters to search...")
            )
            return
        filtered = []
        for it in self.full_items:
            code_l = it['code'].lower()
            desc_l = it['description'].lower()
            if query in code_l or query in desc_l:
                filtered.append(it)
        count = len(filtered)
        msg = lang.t("dispatch_Kit.found_items_count",
                     f"Found {count} matching rows")
        self.populate_rows(filtered, msg)

    def _insert_transaction_issue(self, cur, *, unique_id, code, description,
                                  expiry_date, batch_number, scenario, kit_number,
                                  module_number, qty_out, out_type,
                                  third_party, end_user, remarks, movement_type,
                                  ts_date, ts_time, document_number):
        """
        Insert a single stock issue transaction using an existing cursor.
        Now includes document_number column.
        """
        cur.execute("""
            INSERT INTO stock_transactions
            (Date, Time, unique_id, code, Description, Expiry_date, Batch_Number,
             Scenario, Kit, Module, Qty_IN, IN_Type, Qty_Out, Out_Type,
             Third_Party, End_User, Remarks, Movement_Type, document_number)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            ts_date, ts_time,
            unique_id, code, description, expiry_date, batch_number,
            scenario, kit_number, module_number,
            None, None, qty_out, out_type,
            third_party, end_user, remarks, movement_type, document_number
        ))

    # ------------------------------------------------------------------
    # Generate Document Number
    # ------------------------------------------------------------------

    def generate_document_number(self, out_type_text: str) -> str:
        """
        Generate a document number for OUT (dispatch) operations.

        Format:
            YYYY/MM/<PROJECT_CODE>/<ABBR>/<SERIAL>

        - PROJECT_CODE retrieved from project_details.project_code
        - ABBR: derived from out_type_text using a mapping; if no direct match,
                build from significant uppercase initials (skip small stopwords)
        - SERIAL: 4-digit increment per prefix (0001, 0002, ...)

        Stores in self.current_document_number and returns it.
        """
        project_name, project_code = fetch_project_details()
        project_code = (project_code or "PRJ").strip().upper()

        # Base mapping (keys are English canonical forms; we fuzzy-normalize input)
        base_map = {
            "Issue to End User": "OEU",
            "Expired Items": "OEXP",
            "Damaged Items": "ODMG",
            "Cold Chain Break": "OCCB",
            "Batch Recall": "OBRC",
            "Theft": "OTHF",
            "Other Losses": "OLS",
            "Out Donation": "ODN",
            "Loan": "OLOAN",
            "Return of Borrowing": "OROB",
            "Quarantine": "OQRT"
        }

        raw = (out_type_text or "").strip()
        import re
        norm_raw = re.sub(r'[^a-z0-9]+', '', raw.lower())
        abbr = None
        for k, v in base_map.items():
            if re.sub(r'[^a-z0-9]+', '', k.lower()) == norm_raw:
                abbr = v
                break

        if not abbr:
            # Dynamic abbreviation: keep 'MSF' intact, take first letter of other words
            stop = {"OF", "FROM", "THE", "AND", "DE", "DU", "DES", "LA", "LE", "LES"}
            parts = []
            for token in re.split(r'\s+', raw.upper()):
                if not token or token in stop:
                    continue
                if token == "MSF":
                    parts.append("MSF")
                else:
                    parts.append(token[0])
            if not parts:
                abbr = (raw[:4].upper() or "DOC").replace(" ", "")
            else:
                # Merge, respecting MSF
                merged = []
                for p in parts:
                    merged.append(p)
                abbr = "".join(merged)
            if len(abbr) > 8:
                abbr = abbr[:8]

        now = datetime.now()
        prefix = f"{now.year:04d}/{now.month:02d}/{project_code}/{abbr}"

        # Find latest serial for this prefix
        conn = connect_db()
        serial_num = 1
        if conn is not None:
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
                    last_serial = row[0].rsplit("/", 1)[-1]
                    if last_serial.isdigit():
                        serial_num = int(last_serial) + 1
            except Exception as e:
                logging.error(f"[DISPATCH] Error fetching last document_number: {e}")
            finally:
                cur.close()
                conn.close()

        document_number = f"{prefix}/{serial_num:04d}"
        self.current_document_number = document_number
        return document_number

    # ------------------------------------------------------------------
    # Save (Issue)
    # ------------------------------------------------------------------
    def save_all(self):
        """
        Issue (dispatch) selected quantities.

        CHANGE: Adds document_number creation (one per batch).
        All other behavior retained. Export formatting untouched.
        """
        logging.info("[DISPATCH] save_all called")
        if self.role not in ["admin", "manager"]:
            custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                         lang.t("receive_Kit.no_permission", "Only admin or manager roles can save changes."), "error")
            return

        out_type = (self.trans_type_var.get() or "").strip()
        if not out_type:
            custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                         lang.t("dispatch_kit.no_out_type", "OUT Type is mandatory."), "error")
            return

        third_party_required = {
            lang.t("dispatch_Kit.out_donation", "Out Donation"),
            lang.t("dispatch_Kit.out_loan", "Loan"),
            lang.t("dispatch_Kit.out_return_borrowing", "Return of Borrowing")
        }
        end_user_required = {
            lang.t("dispatch_Kit.out_issue_end_user", "Issue to End User")
        }
        remarks_required = {
            lang.t("dispatch_Kit.out_expired", "Expired Items"),
            lang.t("dispatch_Kit.out_damaged", "Damaged Items"),
            lang.t("dispatch_Kit.out_cold_chain", "Cold Chain Break"),
            lang.t("dispatch_Kit.out_batch_recall", "Batch Recall"),
            lang.t("dispatch_Kit.out_theft", "Theft"),
            lang.t("dispatch_Kit.out_other_losses", "Other Losses"),
            lang.t("dispatch_Kit.out_quarantine", "Quarantine")
        }

        end_user = (self.end_user_var.get() or "").strip()
        third_party = (self.third_party_var.get() or "").strip()
        remarks = (self.remarks_entry.get() or "").strip()

        if out_type in end_user_required and not end_user:
            custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                         lang.t("dispatch_Kit.err_end_user_required", "End User is required for this Out Type."), "error")
            return
        if out_type in third_party_required and not third_party:
            custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                         lang.t("dispatch_Kit.err_third_party_required", "Third Party is required for this Out Type."), "error")
            return
        if out_type in remarks_required and (len(remarks) < 3):
            custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                         lang.t("dispatch_Kit.err_remarks_required", "Remarks are required (min 3 chars) for this Out Type."), "error")
            return

        # Collect rows to issue
        rows_to_issue = []
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if not vals:
                continue
            code, desc, type_field, kit_col, module_col, current_stock, exp_date, batch_no, qty_to_issue, tree_unique_id = vals
            meta = self.row_data.get(iid, {})
            unique_id = meta.get("unique_id") or tree_unique_id
            if not unique_id:
                continue
            raw_q = qty_to_issue[2:].strip() if qty_to_issue.startswith("★") else qty_to_issue
            if not raw_q.isdigit():
                continue
            q_int = int(raw_q)
            if q_int <= 0:
                continue
            try:
                stock_int = int(current_stock) if str(current_stock).isdigit() else 0
            except Exception:
                stock_int = 0
            rows_to_issue.append((iid, code, desc, type_field, stock_int, q_int, exp_date, batch_no, unique_id))

        if not rows_to_issue:
            custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                         lang.t("dispatch_Kit.no_issue_qty", "No quantities entered to issue."), "error")
            return

        over = [code for (_, code, _, _, stock, qty, _, _, _) in rows_to_issue if qty > stock and stock > 0]
        if over:
            custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                         lang.t("dispatch_Kit.over_issue",
                                f"Cannot issue more than stock for: {', '.join(over)}"), "error")
            return

        scenario_name = self.scenario_map.get(self.selected_scenario_id, "")
        movement_label = self.mode_var.get()

        # Generate ONE outbound document number for the whole batch
        doc_number = self.generate_document_number(out_type)
        self.status_var.set(f"Pending dispatch... Document Number: {doc_number}")

        import time, sqlite3
        max_attempts = 4
        for attempt in range(1, max_attempts + 1):
            conn = connect_db()
            if conn is None:
                custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                             lang.t("receive_Kit.db_error", "Database connection failed"), "error")
                return
            try:
                conn.execute("PRAGMA busy_timeout=5000;")
                cur = conn.cursor()
                now_date = datetime.today().strftime('%Y-%m-%d')
                now_time = datetime.now().strftime('%H:%M:%S')

                for (iid, code, desc, type_field, stock, qty, exp_date, batch_no, unique_id) in rows_to_issue:
                    # Update stock_data (ensure available final_qty by checking qty_out only logically)
                    cur.execute("""
                        SELECT final_qty FROM stock_data WHERE unique_id = ?
                    """, (unique_id,))
                    row = cur.fetchone()
                    if not row or row[0] is None or row[0] < qty:
                        raise ValueError(f"Insufficient stock or concurrency issue for {code}")

                    cur.execute("""
                        UPDATE stock_data
                           SET qty_out = qty_out + ?,
                               updated_at = ?
                         WHERE unique_id = ?
                           AND (qty_in - qty_out) >= ?
                    """, (qty, f"{now_date} {now_time}", unique_id, qty))
                    if cur.rowcount == 0:
                        raise ValueError(f"Concurrent change or insufficient stock for {code}")

                    rd = self.row_data.get(iid, {})
                    kit_number = rd.get('Kit_number') or rd.get('kit_number') or kit_col or None
                    module_number = rd.get('module_number') or module_col or None

                    self._insert_transaction_issue(
                        cur,
                        unique_id=unique_id,
                        code=code,
                        description=desc,
                        expiry_date=exp_date if exp_date else None,
                        batch_number=batch_no if batch_no else None,
                        scenario=scenario_name,
                        kit_number=kit_number,
                        module_number=module_number,
                        qty_out=qty,
                        out_type=out_type,
                        third_party=third_party if third_party else None,
                        end_user=end_user if end_user else None,
                        remarks=remarks if remarks else None,
                        movement_type=movement_label,
                        ts_date=now_date,
                        ts_time=now_time,
                        document_number=doc_number
                    )

                conn.commit()
                custom_popup(self.parent, lang.t("dialog_titles.success", "Success"),
                             lang.t("dispatch_Kit.issue_success", "Stock issued successfully."), "info")
                self.status_var.set(f"Issue complete. Document Number: {doc_number}")

                if custom_askyesno(self.parent,
                                   lang.t("dialog_titles.confirm", "Confirm"),
                                   lang.t("dispatch_Kit.ask_export", "Do you want to export the issuance to Excel?")) == "yes":
                    export_tuples = [(iid, code, desc, stock, qty, exp_date, batch_no)
                                     for (iid, code, desc, _type, stock, qty, exp_date, batch_no, _uid) in rows_to_issue]
                    self.export_data(export_tuples)

                self.clear_form()
                return
            except sqlite3.OperationalError as e:
                if "locked" in str(e).lower():
                    logging.warning(f"[DISPATCH] Database locked attempt {attempt}/{max_attempts}; retrying...")
                    try:
                        conn.rollback()
                    except:
                        pass
                    time.sleep(0.8 * attempt)
                else:
                    try:
                        conn.rollback()
                    except:
                        pass
                    logging.error(f"[DISPATCH] Issue failed: {e}")
                    custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                                 lang.t("dispatch_Kit.issue_failed", f"Issue failed: {e}"), "error")
                    return
            except Exception as e:
                try:
                    conn.rollback()
                except:
                    pass
                logging.error(f"[DISPATCH] Issue failed: {e}")
                custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                             lang.t("dispatch_Kit.issue_failed", f"Issue failed: {e}"), "error")
                return
            finally:
                try:
                    cur.close()
                except:
                    pass
                try:
                    conn.close()
                except:
                    pass

        custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                     lang.t("dispatch_Kit.issue_failed", "Issue failed: database remained locked."), "error")

    # ------------------------------------------------------------------
    # Utility / Clear / Export
    # ------------------------------------------------------------------
    def clear_search(self):
        self.search_var.set("")
        self.populate_rows(self.full_items,
                           lang.t("dispatch_Kit.showing_rows_reset",
                                  f"Showing {len(self.full_items)} rows (reset)"))

    def clear_form(self):
        self.clear_table_only()
        self.scenario_var.set("")
        self.mode_var.set("")
        self.Kit_var.set("")
        self.Kit_number_var.set("")
        self.module_var.set("")
        self.module_number_var.set("")
        self.trans_type_var.set("")
        self.end_user_var.set("")
        self.third_party_var.set("")
        self.remarks_entry.config(state="normal")
        self.remarks_entry.delete(0, tk.END)
        self.remarks_entry.config(state="disabled")
        self.status_var.set(lang.t("receive_Kit.ready", "Ready"))
        self.scenario_map = self.fetch_scenario_map()
        self.load_scenarios()

    def export_data(self, rows_to_issue=None):
        """
        Export issued rows to Excel (no unique_id column).
        rows_to_issue: optional list of tuples (iid, code, desc, stock, qty, exp_date, batch_no)
                       If None, will re-scan tree (post-save normally we pass it in).
        Adds Document Number (if available) after Date with 8 spaces separation.
        """
        logging.info("[DISPATCH] export_data called")
        if self.parent is None or not self.parent.winfo_exists():
            logging.error("Parent window is None or does not exist in export_data")
            return
        try:
            # Always scan the tree recursively to get all rows
            export_rows = []
            all_iids = []

            def collect_iids(item=''):
                for child in self.tree.get_children(item):
                    all_iids.append(child)
                    collect_iids(child)
            collect_iids()

            for iid in all_iids:
                vals = self.tree.item(iid, "values")
                if not vals or len(vals) < 10:
                    logging.warning(f"[DISPATCH] Skipping invalid row {iid}: {vals}")
                    continue
                code, desc, tfield, kit_col, module_col, current_stock, exp_date, batch_no, qty_to_issue, _uid = vals
                raw_q = qty_to_issue[2:].strip() if qty_to_issue.startswith("★") else qty_to_issue
                qty = int(raw_q) if raw_q.isdigit() else 0
                rd = self.row_data.get(iid, {})
                kit_no = rd.get('Kit_number') or rd.get('kit_number') or kit_col or ""
                mod_no = rd.get('module_number') or module_col or ""
                export_rows.append({
                    "iid": iid,
                    "code": code,
                    "description": desc,
                    "type": tfield or "Item",
                    "kit_number": kit_no,
                    "module_number": mod_no,
                    "current_stock": int(current_stock or 0),
                    "expiry_date": exp_date or "",
                    "batch_number": batch_no or "",
                    "qty_issued": qty
                })

            # If rows_to_issue is provided, update the corresponding entries
            if rows_to_issue is not None:
                for row in rows_to_issue:
                    if len(row) != 7:
                        logging.warning(f"[DISPATCH] Skipping invalid provided row: {row}")
                        continue
                    rid, code, desc, stock, qty, exp_date, batch_no = row
                    for er in export_rows:
                        if er["iid"] == rid:
                            er["current_stock"] = stock
                            er["qty_issued"] = qty
                            er["expiry_date"] = exp_date or er["expiry_date"]
                            er["batch_number"] = batch_no or er["batch_number"]
                            break

            # Check if there is any issued qty
            has_issued = any(r["qty_issued"] > 0 for r in export_rows)
            if not has_issued:
                custom_popup(self.parent, lang.t("dialog_titles.error", "Error"),
                             lang.t("dispatch_kit.no_issue_qty", "No quantities entered to issue."), "error")
                return

            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            out_type_raw = self.trans_type_var.get() or lang.t("dispatch_kit.unknown", "Unknown")
            movement_type_raw = self.mode_var.get() or lang.t("dispatch_kit.unknown", "Unknown")
            scenario_name = self.selected_scenario_name or "N/A"
            doc_number = getattr(self, "current_document_number", None)

            # Helper to sanitize for filenames
            import re
            def sanitize(s: str) -> str:
                s = re.sub(r'[^A-Za-z0-9]+', '_', s)
                s = re.sub(r'_+', '_', s)
                return s.strip('_') or "Unknown"

            out_type_slug = sanitize(out_type_raw)
            movement_type_slug = sanitize(movement_type_raw)

            # Ask for path
            default_dir = "D:/ISEPREP"
            if not os.path.exists(default_dir):
                os.makedirs(default_dir)

            file_name = f"Dispatch_{movement_type_slug}_{out_type_slug}_{current_time.replace(':', '-')}.xlsx"

            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile=file_name,
                initialdir=default_dir
            )
            if not path:
                self.status_var.set(lang.t("dispatch_kit.export_cancelled", "Export cancelled"))
                return

            wb = openpyxl.Workbook()
            ws = wb.active

            # Worksheet title (Excel limit 31 chars)
            ws_title_base = "Dispatch"
            ws_title = f"{ws_title_base[:15]}-{movement_type_slug[:12]}"
            ws.title = ws_title

            # Row 1: Date + 8 spaces + Document Number (if available)
            if doc_number:
                ws['A1'] = f"Date: {current_time}{' ' * 8}Document Number: {doc_number}"
            else:
                ws['A1'] = f"Date: {current_time}"
            ws['A1'].font = Font(name="Calibri", size=11)
            ws['A1'].alignment = Alignment(horizontal="left")

            project_name, project_code = fetch_project_details()

            # A2: Movement line
            ws['A2'] = f"{ws_title_base} – Movement: {movement_type_raw}"
            ws['A2'].font = Font(name="Tahoma", size=14, bold=True)
            ws['A2'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A2:I2')

            # A3: Project
            ws['A3'] = f"{project_name} - {project_code}"
            ws['A3'].font = Font(name="Tahoma", size=14, bold=True)
            ws['A3'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A3:I3')

            # A4: OUT Type
            ws['A4'] = f"{lang.t('dispatch_kit.out_type', 'OUT Type')}: {out_type_raw}"
            ws['A4'].font = Font(name="Tahoma", size=12, bold=True)
            ws['A4'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A4:I4')

            # A5: Scenario
            ws['A5'] = f"Scenario: {scenario_name}"
            ws['A5'].font = Font(name="Tahoma", size=12, bold=True)
            ws['A5'].alignment = Alignment(horizontal="right")
            ws.merge_cells('A5:I5')

            ws.append([])  # Empty row 6
            ws['A6'].font = Font(name="Tahoma", size=11, bold=True)

            headers = [
                lang.t("dispatch_kit.code", "Code"),
                lang.t("dispatch_kit.description", "Description"),
                lang.t("dispatch_kit.type", "Type"),
                lang.t("dispatch_kit.kit", "Kit Number"),
                lang.t("dispatch_kit.module", "Module Number"),
                lang.t("dispatch_kit.current_stock", "Current Stock"),
                lang.t("dispatch_kit.expiry_date", "Expiry Date"),
                lang.t("dispatch_kit.batch_no", "Batch Number"),
                lang.t("dispatch_kit.qty_to_issue", "Qty Issued")
            ]
            ws.append(headers)

            # Header row styling (row 7)
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=7, column=col)
                cell.font = Font(name="Tahoma", size=11, bold=True)

            from openpyxl.styles import PatternFill
            for row_idx, row in enumerate(export_rows, start=8):
                ws.append([
                    row["code"],
                    row["description"],
                    row["type"],
                    row["kit_number"],
                    row["module_number"],
                    row["current_stock"],
                    row["expiry_date"],
                    row["batch_number"],
                    row["qty_issued"]
                ])
                row_type = row["type"].lower() if row["type"] else ""
                for col in range(1, len(headers) + 1):
                    cell = ws.cell(row=row_idx, column=col)
                    cell.font = Font(name="Calibri", size=11, bold=(row_type in ("kit", "module")))
                    if row_type == "kit":
                        cell.fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
                    elif row_type == "module":
                        cell.fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

            # Autofit (approx)
            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        l = len(str(cell.value)) if cell.value is not None else 0
                        if l > max_length:
                            max_length = l
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToHeight = 0
            ws.page_setup.fitToWidth = 1
            ws.print_title_rows = '1:7'
            ws.oddFooter.center.text = "&P of &N"
            ws.evenFooter.center.text = "&P of &N"

            wb.save(path)
            custom_popup(self.parent,
                         lang.t("dialog_titles.success", "Success"),
                         lang.t("dispatch_kit.export_success", f"Exported to {path}"),
                         "info")
            self.status_var.set(lang.t("dispatch_kit.export_success", f"Export successful: {path}"))
            logging.info(f"[DISPATCH] Exported file: {path}")
        except Exception as e:
            logging.error(f"[DISPATCH] Export failed: {e}")
            custom_popup(self.parent,
                         lang.t("dialog_titles.error", "Error"),
                         lang.t("dispatch_kit.export_failed", f"Export failed: {str(e)}"),
                         "error")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Dispatch")
    app = type("App", (), {})()
    app.role = "admin"
    StockDispatchKit(root, app, role="admin")
    root.mainloop()
