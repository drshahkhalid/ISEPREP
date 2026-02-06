import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import time
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

from db import connect_db
from language_manager import lang
from stock_data import parse_expiry
from manage_items import get_item_description, detect_type
from popup_utils import custom_popup, custom_askyesno, custom_dialog

# Optional custom popups
try:
    from popup_utils import custom_popup, custom_askyesno
except ImportError:

    def custom_popup(parent, title, message, kind="info"):
        if kind == "error":
            messagebox.showerror(title, message)
        elif kind == "warning":
            messagebox.showwarning(title, message)
        else:
            messagebox.showinfo(title, message)

    def custom_askyesno(parent, title, message):
        return "yes" if messagebox.askyesno(title, message) else "no"


_SCENARIO_CACHE = {}


def scenario_id_to_name(scenario_id: str) -> str:
    if not scenario_id:
        return lang.t("stock_inv.unknown", "Unknown")
    if scenario_id in _SCENARIO_CACHE:
        return _SCENARIO_CACHE[scenario_id]
    conn = connect_db()
    if conn is None:
        return scenario_id
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute("SELECT name FROM scenarios WHERE scenario_id = ?", (scenario_id,))
        row = cur.fetchone()
        name = row["name"] if row and row["name"] else scenario_id
        _SCENARIO_CACHE[scenario_id] = name
        return name
    finally:
        cur.close()
        conn.close()


def get_active_designation(code):
    if not code:
        return lang.t("stock_inv.no_description", "No Description")
    conn = connect_db()
    if conn is None:
        return code
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute(
            """SELECT designation, designation_en, designation_fr, designation_sp
                          FROM items_list WHERE code = ?""",
            (code,),
        )
        row = cursor.fetchone()
        if not row:
            return code
        lang_code = lang.lang_code.lower()
        col_map = {
            "en": "designation_en",
            "fr": "designation_fr",
            "es": "designation_sp",
            "sp": "designation_sp",
        }
        preferred = col_map.get(lang_code, "designation_en")
        if row[preferred]:
            return row[preferred]
        if row["designation_en"]:
            return row["designation_en"]
        return row["designation"] or code
    finally:
        cursor.close()
        conn.close()


def check_expiry_required(code):
    if not code:
        return False
    conn = connect_db()
    if conn is None:
        return False
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT remarks FROM items_list WHERE code = ?", (code,))
        row = cursor.fetchone()
        return bool(row and row["remarks"] and "exp" in row["remarks"].lower())
    finally:
        cursor.close()
        conn.close()


def get_std_qty(code, scenario_name):
    conn = connect_db()
    if conn is None:
        return 0
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        if scenario_name == lang.t("stock_inv.all_scenarios", "All Scenarios"):
            cursor.execute(
                "SELECT SUM(quantity) AS total_qty FROM compositions WHERE code = ?",
                (code,),
            )
            row = cursor.fetchone()
            return row["total_qty"] if row and row["total_qty"] is not None else 0
        else:
            cursor.execute(
                """
                SELECT quantity FROM compositions
                 WHERE code = ? AND scenario_id = (
                     SELECT scenario_id FROM scenarios WHERE name = ?
                 )""",
                (code, scenario_name),
            )
            row = cursor.fetchone()
            return row["quantity"] if row and row["quantity"] is not None else 0
    finally:
        cursor.close()
        conn.close()


def get_treecode(scenario_id, kit_code, module_code, item_code):
    """
    Fetch treecode from kit_items for the given scenario, kit, module, item.
    """
    conn = connect_db()
    if conn is None:
        return None
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute(
            """
            SELECT treecode FROM kit_items
            WHERE scenario_id = ? AND kit = ? AND module = ? AND item = ?
            LIMIT 1
        """,
            (scenario_id, kit_code or None, module_code or None, item_code),
        )
        row = cur.fetchone()
        return row["treecode"] if row else None
    finally:
        cur.close()
        conn.close()


def get_item_type(code: str) -> str:
    if not code:
        return lang.t("stock_inv.item", "Item")
    conn = connect_db()
    if conn is None:
        return lang.t("stock_inv.item", "Item")
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT type, designation_en FROM items_list WHERE code = ?", (code,)
        )
        row = cur.fetchone()
        if row and row["type"]:
            t = row["type"].strip().upper()
            if t in ("KIT", "MODULE", "ITEM"):
                return lang.t(f"stock_inv.{t.lower()}", t.title())
        desc = row["designation_en"] if row else get_item_description(code)
        dt = detect_type(code, desc) if desc else "ITEM"
        return {
            "KIT": lang.t("stock_inv.kit", "Kit"),
            "MODULE": lang.t("stock_inv.module", "Module"),
            "ITEM": lang.t("stock_inv.item", "Item"),
        }.get(dt.upper(), lang.t("stock_inv.item", "Item"))
    finally:
        cur.close()
        conn.close()


def validate_expiry_for_save(code, expiry_date):
    """
    For items requiring expiry (remarks contains 'exp'), an expiry date must exist and be future.
    """
    if not check_expiry_required(code):
        return True
    parsed = parse_expiry(expiry_date) if expiry_date else None
    if not parsed:
        return False
    # parsed may return datetime or date; convert to date
    if hasattr(parsed, "date"):
        parsed_date = parsed.date()
    else:
        parsed_date = parsed
    return parsed_date > datetime.now().date()


def parse_inventory_unique_id(unique_id: str) -> dict:
    parts = unique_id.split("/") if unique_id else []
    out = {
        "scenario_id": None,
        "scenario_name": None,
        "kit_code": None,
        "module_code": None,
        "item_code": None,
        "std_qty": None,
        "exp_date": None,
        "kit_number": None,
        "module_number": None,
        "treecode": None,
    }
    if not parts:
        return out
    if len(parts) >= 1:
        out["scenario_id"] = parts[0]
        out["scenario_name"] = scenario_id_to_name(parts[0])
    if len(parts) >= 2 and parts[1] not in ("None", ""):
        out["kit_code"] = parts[1]
    if len(parts) >= 3 and parts[2] not in ("None", ""):
        out["module_code"] = parts[2]
    if len(parts) >= 4 and parts[3] not in ("None", ""):
        out["item_code"] = parts[3]
    if len(parts) >= 5:
        try:
            out["std_qty"] = int(parts[4])
        except:
            out["std_qty"] = None
    if len(parts) >= 6 and parts[5] not in ("None", ""):
        out["exp_date"] = parts[5]
    if len(parts) >= 7 and parts[6] not in ("None", ""):
        out["kit_number"] = parts[6]
    if len(parts) >= 8 and parts[7] not in ("None", ""):
        out["module_number"] = parts[7]
    if len(parts) >= 9 and parts[8] not in ("None", ""):
        out["treecode"] = parts[8]
    return out


def construct_unique_id(
    scenario_id: str,
    kit_code: str,
    module_code: str,
    item_code: str,
    std_qty: int,
    exp_date: str,
    kit_number: str = None,
    module_number: str = None,
    force_box_format: bool = False,
    treecode: str = None,
) -> str:
    base = [
        scenario_id or "0",
        kit_code or "None",
        module_code or "None",
        item_code or "None",
        str(std_qty if std_qty is not None else 0),
        exp_date or "None",
    ]
    if kit_number or module_number or force_box_format:
        base.append(kit_number or "None")
        base.append(module_number or "None")
        if treecode:
            base.append(treecode)
    return "/".join(base)


class StockInventory(tk.Frame):
    def __init__(self, parent, app, role="supervisor"):
        super().__init__(parent)
        self.app = app
        self.role = role.lower()
        self.scenario_map = self.fetch_scenario_map()
        self.ctx_menu = None
        self.ctx_row = None
        self.base_physical_inputs = {}
        self.user_row_states = {}
        self.temp_row_counter = 0
        self.pack(fill="both", expand=True)
        self.render_ui()

    # ---------- DB fetchers ----------
    def fetch_scenario_map(self):
        conn = connect_db()
        if conn is None:
            return {}
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("SELECT scenario_id, name FROM scenarios")
            return {str(r["scenario_id"]): r["name"] for r in cur.fetchall()}
        finally:
            cur.close()
            conn.close()

    def fetch_kit_numbers(self, scenario_name=None):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            sql = """SELECT DISTINCT kit_number
                       FROM stock_data
                      WHERE kit_number IS NOT NULL
                        AND kit_number != 'None'"""
            params = []
            if scenario_name and scenario_name != lang.t(
                "stock_inv.all_scenarios", "All Scenarios"
            ):
                sql += " AND scenario = ?"
                params.append(scenario_name)
            sql += " ORDER BY kit_number"
            cur.execute(sql, params)
            return [r["kit_number"] for r in cur.fetchall()]
        finally:
            cur.close()
            conn.close()

    def fetch_module_numbers(self, scenario_name=None, kit_number=None):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            sql = """SELECT DISTINCT module_number
                       FROM stock_data
                      WHERE module_number IS NOT NULL
                        AND module_number != 'None'"""
            params = []
            if scenario_name and scenario_name != lang.t(
                "stock_inv.all_scenarios", "All Scenarios"
            ):
                sql += " AND scenario = ?"
                params.append(scenario_name)
            if kit_number and kit_number not in (
                lang.t("stock_inv.all_kits", "All Kits"),
                lang.t("stock_inv.stand_alone_items", "Stand alone items"),
            ):
                sql += " AND kit_number = ?"
                params.append(kit_number)
            sql += " ORDER BY module_number"
            cur.execute(sql, params)
            return [r["module_number"] for r in cur.fetchall()]
        finally:
            cur.close()
            conn.close()

    def fetch_project_details(self):
        try:
            conn = connect_db()
            if conn is None:
                return (
                    lang.t("stock_inv.unknown", "Unknown"),
                    lang.t("stock_inv.unknown", "Unknown"),
                )
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            try:
                cur.execute(
                    "SELECT project_name, project_code FROM project_details ORDER BY id DESC LIMIT 1"
                )
                row = cur.fetchone()
                return (
                    (row["project_name"], row["project_code"])
                    if row
                    else (
                        lang.t("stock_inv.unknown", "Unknown"),
                        lang.t("stock_inv.unknown", "Unknown"),
                    )
                )
            finally:
                cur.close()
                conn.close()
        except Exception as e:
            # Log or handle the error, return defaults
            print(f"Error fetching project details: {e}")
            return (
                lang.t("stock_inv.unknown", "Unknown"),
                lang.t("stock_inv.unknown", "Unknown"),
            )

    # ---------- Dropdown refresh ----------
    def refresh_kit_dropdown(self):
        scenario_val = self.scenario_var.get()
        kits = self.fetch_kit_numbers(scenario_val)
        all_label = lang.t("stock_inv.all_kits", "All Kits")
        standalone_label = lang.t("stock_inv.stand_alone_items", "Stand alone items")
        kits = [k for k in kits if k not in (all_label, standalone_label)]
        kits_sorted = sorted(kits) + [standalone_label, all_label]
        current = self.kit_number_var.get()
        self.kit_number_cb["values"] = kits_sorted
        if current not in kits_sorted:
            self.kit_number_var.set(all_label)

    def refresh_module_dropdown(self):
        scenario_val = self.scenario_var.get()
        kit_val = self.kit_number_var.get()
        modules = self.fetch_module_numbers(
            scenario_val,
            (
                kit_val
                if kit_val
                not in (
                    lang.t("stock_inv.all_kits", "All Kits"),
                    lang.t("stock_inv.stand_alone_items", "Stand alone items"),
                )
                else None
            ),
        )
        all_label = lang.t("stock_inv.all_modules", "All Modules")
        if all_label not in modules:
            modules.append(all_label)
        modules_sorted = [m for m in modules if m != all_label] + [all_label]
        self.module_number_cb["values"] = modules_sorted
        if self.module_number_var.get() not in modules_sorted:
            self.module_number_var.set(all_label)

    # ---------- Search base ----------
    def fetch_search_results(self, query):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            sql = """
                SELECT DISTINCT c.code
                  FROM compositions c
                  JOIN items_list i ON c.code = i.code
                 WHERE c.quantity > 0
            """
            params = []
            if query:
                lang_code = lang.lang_code.lower()
                col_map = {
                    "en": "designation_en",
                    "fr": "designation_fr",
                    "es": "designation_sp",
                    "sp": "designation_sp",
                }
                active_designation = col_map.get(lang_code, "designation_en")
                sql += f"""
                   AND (
                        LOWER(c.code) LIKE ?
                     OR LOWER(i.{active_designation}) LIKE ?
                     OR (i.{active_designation} IS NULL AND LOWER(i.designation_en) LIKE ?)
                     OR (i.{active_designation} IS NULL AND i.designation_en IS NULL AND LOWER(i.designation) LIKE ?)
                   )
                """
                like = f"%{query.lower()}%"
                params.extend([like, like, like, like])
            if self.scenario_var.get() != lang.t(
                "stock_inv.all_scenarios", "All Scenarios"
            ):
                sql += " AND c.scenario_id = (SELECT scenario_id FROM scenarios WHERE name = ?)"
                params.append(self.scenario_var.get())
            sql += " ORDER BY c.code"
            cur.execute(sql, params)
            rows = cur.fetchall()
            return [
                {"code": r["code"], "description": get_active_designation(r["code"])}
                for r in rows
            ]
        finally:
            cur.close()
            conn.close()

    # ---------- Items in stock with filters ----------
    def fetch_items_in_stock(self, code=None):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            sql = """SELECT unique_id, final_qty FROM stock_data WHERE final_qty != 0"""
            params = []
            if code:
                sql += " AND unique_id LIKE ?"
                params.append(f"%/{code}/%")
            cur.execute(sql, params)
            rows = cur.fetchall()

            scenario_filter = self.scenario_var.get()
            kit_filter = self.kit_number_var.get()
            module_filter = self.module_number_var.get()
            mgmt_mode = self.mgmt_mode_var.get()

            standalone_label = lang.t(
                "stock_inv.stand_alone_items", "Stand alone items"
            )
            all_kits_label = lang.t("stock_inv.all_kits", "All Kits")
            all_modules_label = lang.t("stock_inv.all_modules", "All Modules")

            want_standalone = kit_filter == standalone_label
            on_shelf = mgmt_mode == lang.t("stock_inv.management_on_shelf", "On-Shelf")
            in_box = mgmt_mode == lang.t("stock_inv.management_in_box", "In-Box")

            six_segment = []
            eight_segment = []

            for r in rows:
                uid = r["unique_id"]
                plen = len(uid.split("/"))
                parsed = parse_inventory_unique_id(uid)
                scenario_name = parsed["scenario_name"] or lang.t(
                    "stock_inv.unknown", "Unknown"
                )

                if (
                    scenario_filter
                    != lang.t("stock_inv.all_scenarios", "All Scenarios")
                    and scenario_name != scenario_filter
                ):
                    continue

                kit_number_val = parsed["kit_number"]
                module_number_val = parsed["module_number"]

                if on_shelf and plen != 6:
                    continue
                if in_box and plen != 8:
                    continue

                if want_standalone:
                    if kit_number_val or module_number_val:
                        continue
                else:
                    if not on_shelf:
                        if kit_filter not in (all_kits_label, standalone_label):
                            if (kit_number_val or "-----") != kit_filter:
                                continue
                        if module_filter != all_modules_label:
                            if (module_number_val or "-----") != module_filter:
                                continue

                display_code = (
                    parsed["item_code"]
                    or parsed["module_code"]
                    or parsed["kit_code"]
                    or ""
                )
                description = get_active_designation(display_code)
                item_type = get_item_type(display_code)
                current_stock = r["final_qty"]

                record = {
                    "unique_id": uid,
                    "code": display_code,
                    "description": description,
                    "type": item_type,
                    "scenario": scenario_name,
                    "kit_number": (kit_number_val or "-----"),
                    "module_number": (module_number_val or "-----"),
                    "current_stock": current_stock,
                    "exp_date": parsed["exp_date"] or "",
                    "std_qty": parsed["std_qty"],
                }
                if plen == 6:
                    six_segment.append(record)
                else:
                    eight_segment.append(record)

            return six_segment + eight_segment
        finally:
            cur.close()
            conn.close()

    # ---------- State preservation helpers ----------
    def capture_current_rows(self):
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if not vals:
                continue
            uid = vals[0]
            base_ph = self.base_physical_inputs.get(
                iid, int(vals[9]) if str(vals[9]).isdigit() else 0
            )
            self.user_row_states[uid] = {
                "unique_id": uid,
                "code": vals[1],
                "description": vals[2],
                "type": vals[3],
                "scenario": vals[4],
                "kit_number": vals[5],
                "module_number": vals[6],
                "current_stock": vals[7],
                "exp_date": vals[8],
                "physical_qty": vals[9],
                "updated_exp_date": vals[10],
                "discrepancy": vals[11],
                "remarks": vals[12],
                "std_qty": vals[13],
                "base_physical": base_ph,
                "is_custom": uid.startswith("temp::") or vals[7] in ("0", "", "None"),
            }

    def build_filtered_dataset(self):
        fetched = self.fetch_items_in_stock()
        fetched_map = {r["unique_id"]: r for r in fetched}
        merged = []
        for item in fetched:
            state = self.user_row_states.get(item["unique_id"])
            if state:
                phys = state["physical_qty"]
                upd = state["updated_exp_date"]
                disc = state["discrepancy"]
                remarks = state.get("remarks", "")
                base_ph = state.get(
                    "base_physical", int(phys) if str(phys).isdigit() else 0
                )
            else:
                phys = ""  # blank by default
                upd = ""
                disc = ""  # blank discrepancy
                remarks = ""
                base_ph = 0
            merged.append((item, phys, upd, disc, remarks, base_ph))
        for uid, state in self.user_row_states.items():
            if uid not in fetched_map and state.get("is_custom"):
                scen_filter = self.scenario_var.get()
                if (
                    scen_filter != lang.t("stock_inv.all_scenarios", "All Scenarios")
                    and state["scenario"] != scen_filter
                ):
                    continue
                merged.append(
                    (
                        {
                            "unique_id": state["unique_id"],
                            "code": state["code"],
                            "description": state["description"],
                            "type": state["type"],
                            "scenario": state["scenario"],
                            "kit_number": state["kit_number"],
                            "module_number": state["module_number"],
                            "current_stock": state["current_stock"],
                            "exp_date": state["exp_date"],
                            "std_qty": state["std_qty"],
                        },
                        state["physical_qty"],
                        state["updated_exp_date"],
                        state["discrepancy"],
                        state.get("remarks", ""),
                        state["base_physical"],
                    )
                )
        return merged

    def rebuild_tree_preserving_state(self):
        self.capture_current_rows()
        self.tree.delete(*self.tree.get_children())
        dataset = self.build_filtered_dataset()

        # Check if we are in "Complete Inventory" mode
        is_complete_inventory = self.inv_type_var.get() == lang.t(
            "stock_inv.complete_inventory", "Complete Inventory"
        )

        for item, phys, upd, disc, remarks, base_ph in dataset:
            # If in complete inventory mode, force physical quantity and base input to 0 by default
            if is_complete_inventory:
                phys_to_insert = "0"
                base_ph_to_insert = 0
            else:
                phys_to_insert = phys
                base_ph_to_insert = base_ph

            iid = self._insert_from_state(
                item, phys_to_insert, upd, disc, remarks, base_ph_to_insert
            )

            # If we forced a zero, also update the stored user state
            if is_complete_inventory:
                uid = item["unique_id"]
                if uid in self.user_row_states:
                    self.user_row_states[uid]["physical_qty"] = "0"
                    self.user_row_states[uid]["base_physical"] = 0

        # The recompute function will handle all final calculations and discrepancies
        self.recompute_all_physical_quantities()

        self.status_var.set(
            lang.t("stock_inv.loaded_records", "Loaded {count} records").format(
                count=len(self.tree.get_children())
            )
        )

    def _insert_from_state(
        self, item_dict, physical_qty, updated_exp, discrepancy, remarks, base_physical
    ):
        phys_str = physical_qty if physical_qty and str(physical_qty).isdigit() else ""
        disc_str = discrepancy if discrepancy not in ("", None) else ""
        iid = self.tree.insert(
            "",
            "end",
            values=(
                item_dict["unique_id"],
                item_dict["code"],
                item_dict["description"],
                item_dict["type"],
                item_dict["scenario"],
                item_dict.get("kit_number", "-----"),
                item_dict.get("module_number", "-----"),
                item_dict["current_stock"],
                item_dict["exp_date"],
                phys_str,
                updated_exp,
                disc_str,
                remarks,
                item_dict["std_qty"],
            ),
        )
        t = (item_dict["type"] or "").upper()
        if t == "KIT":
            self.tree.item(iid, tags=("kit_row",))
        elif t == "MODULE":
            self.tree.item(iid, tags=("module_row",))
        else:
            self.tree.item(iid, tags=())
        self.base_physical_inputs[iid] = base_physical
        return iid

    # ---------- In-box multiplier logic ----------

    def recompute_all_physical_quantities(self):
        """
        Universal quantity calculation engine.
        - Determines logic based on unique_id structure (8-segment vs. other).
        - For 8-segment 'physical' rows:
          - KIT/MODULE base inputs are coerced to 0 or 1 for factors.
          - Final quantity is calculated as: base * kit_factor * module_factor.
        - For all other rows, the final quantity is the user's base input.
        - Updates discrepancies and highlights for all rows after computation.
        - Includes enhanced popups and state synchronization.
        """
        # --- Phase 1: Calculate all final quantities ---
        kit_factors = {}
        module_factors = {}

        all_iids = self.tree.get_children()

        # First pass: Identify all KIT/MODULE rows and calculate their factors (0 or 1)
        for iid in all_iids:
            vals = self.tree.item(iid, "values")
            if not vals:
                continue

            unique_id = vals[0]
            is_physical = unique_id and unique_id.count("/") == 7
            if not is_physical:
                continue

            row_type = (vals[3] or "").upper()
            base_input = self.base_physical_inputs.get(iid, 0)

            if row_type == "KIT":
                kit_number = vals[5]
                if kit_number and kit_number != "-----":
                    final_qty = 1 if base_input > 0 else 0
                    kit_factors[kit_number] = final_qty
                    self.tree.set(iid, "physical_qty", str(final_qty))

            elif row_type == "MODULE":
                module_number = vals[6]
                kit_number = vals[5]
                if module_number and module_number != "-----":
                    parent_kit_factor = kit_factors.get(kit_number, 1)
                    base_factor = 1 if base_input > 0 else 0
                    final_qty = base_factor * parent_kit_factor
                    module_factors[module_number] = final_qty
                    self.tree.set(iid, "physical_qty", str(final_qty))

                    # Enhanced popup for MODULEs zeroed out by a KIT
                    if base_input > 0 and final_qty == 0 and parent_kit_factor == 0:
                        custom_popup(
                            self,
                            lang.t("dialog_titles.info", "Info"),
                            lang.t(
                                "stock_inv.kit_zero_module_zero",
                                "Quantity for {code} {desc} will remain 0 as the kit number {kit} has 0 quantity.",
                            ).format(code=vals[1], desc=vals[2], kit=kit_number),
                            "info",
                        )
                        # State Synchronization: Reset base input to prevent repeated popups
                        self.base_physical_inputs[iid] = 0

        # Second pass: Calculate quantities for all other rows (ITEMS and non-physical)
        for iid in all_iids:
            vals = self.tree.item(iid, "values")
            if not vals:
                continue

            unique_id = vals[0]
            is_physical = unique_id and unique_id.count("/") == 7
            base_input = self.base_physical_inputs.get(iid, 0)
            row_type = (vals[3] or "").upper()

            final_qty = base_input  # Default for non-physical rows

            if is_physical and row_type == "ITEM":
                kit_number = vals[5]
                module_number = vals[6]

                parent_kit_factor = kit_factors.get(kit_number, 1)
                parent_module_factor = module_factors.get(module_number, 1)

                final_qty = base_input * parent_kit_factor * parent_module_factor
                self.tree.set(iid, "physical_qty", str(final_qty))

                if base_input > 0 and final_qty == 0:
                    reason = ""
                    if parent_module_factor == 0:
                        reason = (
                            lang.t("stock_inv.module_number", "module number")
                            + f" {module_number}"
                        )
                    elif parent_kit_factor == 0:
                        reason = (
                            lang.t("stock_inv.kit_number", "kit number")
                            + f" {kit_number}"
                        )

                    if reason:
                        custom_popup(
                            self,
                            lang.t("dialog_titles.info", "Info"),
                            lang.t(
                                "stock_inv.item_zero_reason",
                                "Quantity for {code} {desc} will remain 0 as the {reason} has 0 quantity.",
                            ).format(code=vals[1], desc=vals[2], reason=reason),
                            "info",
                        )
                        # State Synchronization: Reset base input
                        self.base_physical_inputs[iid] = 0

            elif not is_physical:
                self.tree.set(iid, "physical_qty", str(final_qty))

        # --- Phase 2: Update discrepancies for all rows ---
        for iid in all_iids:
            vals = self.tree.item(iid, "values")
            if not vals:
                continue

            current_stock = int(vals[7]) if str(vals[7]).isdigit() else 0
            physical_qty = int(vals[9]) if str(vals[9]).isdigit() else 0
            discrepancy = physical_qty - current_stock

            self.tree.set(
                iid, "discrepancy", "" if discrepancy == 0 else str(discrepancy)
            )
            self._update_state_from_row(iid)

        # --- Phase 3: Re-apply all visual highlights ---
        self._highlight_missing_required_expiry()

    # ---------- Document number generation ----------
    def generate_document_number(self):
        project_name, project_code = self.fetch_project_details()
        project_code = (project_code or "PRJ").strip().upper()
        inv_type_label = self.inv_type_var.get()
        abbr = (
            "CINV"
            if inv_type_label
            == lang.t("stock_inv.complete_inventory", "Complete Inventory")
            else "PINV"
        )
        now = datetime.now()
        prefix = f"{now.year:04d}/{now.month:02d}/{project_code}/{abbr}"
        conn = connect_db()
        serial = 1
        if conn:
            cur = conn.cursor()
            try:
                cur.execute(
                    """
                    SELECT document_number FROM stock_transactions
                    WHERE document_number LIKE ?
                    ORDER BY document_number DESC LIMIT 1
                """,
                    (prefix + "/%",),
                )
                row = cur.fetchone()
                if row and row[0]:
                    last_serial = row[0].rsplit("/", 1)[-1]
                    if last_serial.isdigit():
                        serial = int(last_serial) + 1
            finally:
                cur.close()
                conn.close()
        return f"{prefix}/{serial:04d}"

    # ---------- Export ----------
    def export_to_excel(self, rows_to_export=None, document_number=None):
        try:
            default_dir = "D:/ISEPREP"
            os.makedirs(default_dir, exist_ok=True)
            inv_type = self.inv_type_var.get().replace(" ", "_")
            mgmt_mode = self.mgmt_mode_var.get().replace(" ", "_")
            current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            file_name = (
                f"IsEPREP_Stock-Inventory_{inv_type}_{mgmt_mode}_{current_time}.xlsx"
            )
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[(lang.t("stock_inv.excel_files", "Excel Files"), "*.xlsx")],
                title=lang.t("stock_inv.save_excel", "Save Excel"),
                initialfile=file_name,
                initialdir=default_dir,
            )
            if not path:
                self.status_var.set(
                    lang.t("stock_inv.export_cancelled", "Export cancelled")
                )
                return

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = lang.t("stock_inv.stock_inventory", "Stock Inventory")

            project_name, project_code = self.fetch_project_details()
            inv_type_label = self.inv_type_var.get()
            mgmt_mode_label = self.mgmt_mode_var.get()
            scenario = self.scenario_var.get()
            kit_filter = self.kit_number_var.get()
            module_filter = self.module_number_var.get()
            current_date = datetime.now().strftime("%Y-%m-%d")

            ws["A1"] = lang.t("stock_inv.stock_inventory", "Stock Inventory")
            ws["A1"].font = Font(name="Tahoma", size=14, bold=True)
            ws["A1"].alignment = Alignment(horizontal="right")
            ws.merge_cells("A1:L1")

            ws["A2"] = f"{project_name} - {project_code}"
            ws["A2"].font = Font(name="Tahoma", size=14)
            ws["A2"].alignment = Alignment(horizontal="right")
            ws.merge_cells("A2:L2")

            ws["A3"] = (
                f"{lang.t('stock_inv.inventory_type', 'Inventory Type')}: {inv_type_label}, "
                f"{lang.t('stock_inv.management_mode', 'Management Mode')}: {mgmt_mode_label}, "
                f"{lang.t('stock_inv.scenario', 'Scenario')}: {scenario}, "
                f"{lang.t('stock_inv.kit_number', 'Kit Number')}: {kit_filter}, "
                f"{lang.t('stock_inv.module_number', 'Module Number')}: {module_filter}"
            )
            ws["A3"].font = Font(name="Tahoma")
            ws["A3"].alignment = Alignment(horizontal="right")
            ws.merge_cells("A3:L3")

            ws["A4"] = (
                f"{lang.t('stock_inv.inventory_date', 'Inventory Date')}: {current_date}"
            )
            ws["A4"].font = Font(name="Tahoma")
            ws["A4"].alignment = Alignment(horizontal="right")
            ws.merge_cells("A4:L4")

            if document_number:
                ws["A5"] = (
                    f"{lang.t('stock_inv.document_number', 'Document Number')}: {document_number}"
                )
                ws["A5"].font = Font(name="Tahoma")
                ws["A5"].alignment = Alignment(horizontal="right")
                ws.merge_cells("A5:L5")
                ws.append([])

            headers = [
                lang.t("stock_inv.code", "Code"),
                lang.t("stock_inv.description", "Description"),
                lang.t("stock_inv.type", "Type"),
                lang.t("stock_inv.scenario", "Scenario"),
                lang.t("stock_inv.kit_number", "Kit Number"),
                lang.t("stock_inv.module_number", "Module Number"),
                lang.t("stock_inv.current_stock", "Current Stock"),
                lang.t("stock_inv.expiry_date", "Expiry Date"),
                lang.t("stock_inv.physical_qty", "Physical Quantity"),
                lang.t("stock_inv.updated_exp_date", "Updated Expiry Date"),
                lang.t("stock_inv.discrepancy", "Discrepancy"),
                lang.t("stock_inv.remarks", "Remarks"),
            ]
            ws.append(headers)
            kit_fill = PatternFill(
                start_color="D8F5D0", end_color="D8F5D0", fill_type="solid"
            )
            module_fill = PatternFill(
                start_color="D5ECFF", end_color="D5ECFF", fill_type="solid"
            )
            exp_warn_fill = PatternFill(
                start_color="FFD8D8", end_color="FFD8D8", fill_type="solid"
            )

            rows_data = rows_to_export or [
                {
                    "code": vals[1],
                    "description": vals[2],
                    "type": vals[3],
                    "scenario": vals[4],
                    "kit_number": vals[5],
                    "module_number": vals[6],
                    "current_stock": vals[7],
                    "exp_date": vals[8],
                    "physical_qty": vals[9],
                    "updated_exp_date": vals[10],
                    "discrepancy": vals[11],
                    "remarks": vals[12],
                }
                for item in self.tree.get_children()
                if (vals := self.tree.item(item)["values"])
            ]

            row_start = ws.max_row + 1
            for idx, row in enumerate(rows_data, start=row_start):
                ws.append(
                    [
                        row["code"],
                        row["description"],
                        row["type"],
                        row["scenario"],
                        row["kit_number"],
                        row["module_number"],
                        row["current_stock"],
                        row["exp_date"],
                        row["physical_qty"],
                        row["updated_exp_date"],
                        row["discrepancy"],
                        row["remarks"],
                    ]
                )
                t = (row["type"] or "").lower()
                fill = None
                if t == "kit":
                    fill = kit_fill
                elif t == "module":
                    fill = module_fill
                # highlight missing required expiry
                if check_expiry_required(row["code"]) and not (
                    row["updated_exp_date"] or row["exp_date"]
                ):
                    fill = exp_warn_fill
                if fill:
                    for c in ws[f"A{idx}:L{idx}"]:
                        for cell in c:
                            cell.fill = fill

            widths = [100, 300, 100, 120, 120, 130, 110, 110, 120, 130, 110, 200]
            for i, w in enumerate(widths, start=1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w / 7

            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToHeight = 0
            ws.page_setup.fitToWidth = 1

            wb.save(path)
            wb.close()
            custom_popup(
                self,
                lang.t("dialog_titles.success", "Success"),
                lang.t("stock_inv.exported", "Exported to {path}").format(path=path),
                "info",
            )
            self.status_var.set(
                lang.t("stock_inv.export_success", "Export successful: {path}").format(
                    path=path
                )
            )
        except Exception as e:
            custom_popup(
                self,
                lang.t("dialog_titles.error", "Error"),
                lang.t("stock_inv.export_failed", "Export failed: {error}").format(
                    error=str(e)
                ),
                "error",
            )
            self.status_var.set(
                lang.t("stock_inv.export_error", "Export error: {error}").format(
                    error=str(e)
                )
            )

    # ---------- UI ----------
    def render_ui(self):
        for w in self.winfo_children():
            w.destroy()

        tk.Label(
            self,
            text=lang.t("stock_inv.title", "Stock Inventory Adjustment"),
            font=("Helvetica", 20, "bold"),
            bg="#F5F5F5",
        ).pack(pady=10, anchor="w")

        filter_frame = tk.Frame(self, bg="#F5F5F5")
        filter_frame.pack(pady=5, fill="x")

        tk.Label(
            filter_frame, text=lang.t("stock_inv.scenario", "Scenario:"), bg="#F5F5F5"
        ).grid(row=0, column=0, padx=5, sticky="w")
        self.scenario_var = tk.StringVar(
            value=lang.t("stock_inv.all_scenarios", "All Scenarios")
        )
        scenarios = list(self.scenario_map.values())
        all_scen = lang.t("stock_inv.all_scenarios", "All Scenarios")
        if all_scen not in scenarios:
            scenarios.append(all_scen)
        self.scenario_cb = ttk.Combobox(
            filter_frame,
            textvariable=self.scenario_var,
            values=scenarios,
            state="readonly",
            width=20,
        )
        self.scenario_cb.grid(row=0, column=1, padx=5, pady=5)
        self.scenario_cb.bind("<<ComboboxSelected>>", self.on_scenario_selected)

        tk.Label(
            filter_frame,
            text=lang.t("stock_inv.management_mode", "Management Mode:"),
            bg="#F5F5F5",
        ).grid(row=0, column=2, padx=5, sticky="w")
        self.mgmt_mode_var = tk.StringVar(
            value=lang.t("stock_inv.management_all", "All")
        )
        self.mgmt_mode_cb = ttk.Combobox(
            filter_frame,
            textvariable=self.mgmt_mode_var,
            values=[
                lang.t("stock_inv.management_all", "All"),
                lang.t("stock_inv.management_on_shelf", "On-Shelf"),
                lang.t("stock_inv.management_in_box", "In-Box"),
            ],
            state="readonly",
            width=14,
        )
        self.mgmt_mode_cb.grid(row=0, column=3, padx=5, pady=5)
        self.mgmt_mode_cb.bind("<<ComboboxSelected>>", self.on_management_mode_change)

        tk.Label(
            filter_frame,
            text=lang.t("stock_inv.kit_number", "Kit Number:"),
            bg="#F5F5F5",
        ).grid(row=0, column=4, padx=5, sticky="w")
        self.kit_number_var = tk.StringVar(
            value=lang.t("stock_inv.all_kits", "All Kits")
        )
        self.kit_number_cb = ttk.Combobox(
            filter_frame,
            textvariable=self.kit_number_var,
            values=[lang.t("stock_inv.all_kits", "All Kits")],
            state="readonly",
            width=14,
        )
        self.kit_number_cb.grid(row=0, column=5, padx=5, pady=5)
        self.kit_number_cb.bind("<<ComboboxSelected>>", self.on_filter_change)

        tk.Label(
            filter_frame,
            text=lang.t("stock_inv.module_number", "Module Number:"),
            bg="#F5F5F5",
        ).grid(row=0, column=6, padx=5, sticky="w")
        self.module_number_var = tk.StringVar(
            value=lang.t("stock_inv.all_modules", "All Modules")
        )
        self.module_number_cb = ttk.Combobox(
            filter_frame,
            textvariable=self.module_number_var,
            values=[lang.t("stock_inv.all_modules", "All Modules")],
            state="readonly",
            width=14,
        )
        self.module_number_cb.grid(row=0, column=7, padx=5, pady=5)
        self.module_number_cb.bind("<<ComboboxSelected>>", self.on_filter_change)

        # Inventory Type
        type_frame = tk.Frame(self, bg="#F5F5F5")
        type_frame.pack(pady=5, fill="x")
        tk.Label(
            type_frame,
            text=lang.t("stock_inv.inventory_type", "Inventory Type:"),
            bg="#F5F5F5",
        ).grid(row=0, column=0, padx=5, sticky="w")
        self.inv_type_var = tk.StringVar(
            value=lang.t("stock_inv.complete_inventory", "Complete Inventory")
        )
        self.inv_type_cb = ttk.Combobox(
            type_frame,
            textvariable=self.inv_type_var,
            state="readonly",
            values=[
                lang.t("stock_inv.complete_inventory", "Complete Inventory"),
                lang.t("stock_inv.partial_inventory", "Partial Inventory"),
            ],
            width=28,
        )
        self.inv_type_cb.grid(row=0, column=1, padx=5, pady=5)
        self.inv_type_cb.bind("<<ComboboxSelected>>", self.on_inv_type_selected)

        # Search (partial only)
        self.search_frame = tk.Frame(self, bg="#F5F5F5")
        tk.Label(
            self.search_frame,
            text=lang.t("stock_inv.search", "Search Code or Description:"),
            bg="#F5F5F5",
        ).grid(row=0, column=0, padx=5, sticky="w")
        self.code_entry = tk.Entry(self.search_frame, width=55)
        self.code_entry.grid(row=0, column=1, padx=5, pady=5)
        self.code_entry.bind("<KeyRelease>", self.on_search_keyrelease)
        self.code_entry.bind("<Return>", self.select_first_search_result)
        self.search_listbox = tk.Listbox(self.search_frame, height=5, width=50)
        self.search_listbox.grid(row=0, column=2, padx=5, pady=5)
        self.search_listbox.bind("<<ListboxSelect>>", self.on_search_select)

        self.button_frame = tk.Frame(self, bg="#F5F5F5")
        self.button_frame.pack(fill="x")
        tk.Button(
            self.button_frame,
            text=lang.t("stock_inv.add_missing_item", "Add Missing Item"),
            bg="#2980B9",
            fg="white",
            command=self.add_missing_item,
        ).pack(side="right", padx=5, pady=5)

        # Tree
        tree_frame = tk.Frame(self)
        tree_frame.pack(expand=True, fill="both", pady=10)
        self.cols = (
            "unique_id",
            "code",
            "description",
            "type",
            "scenario",
            "kit_number",
            "module_number",
            "current_stock",
            "exp_date",
            "physical_qty",
            "updated_exp_date",
            "discrepancy",
            "remarks",
            "std_qty",
        )
        self.display_cols = (
            "code",
            "description",
            "type",
            "scenario",
            "kit_number",
            "module_number",
            "current_stock",
            "exp_date",
            "physical_qty",
            "updated_exp_date",
            "discrepancy",
            "remarks",
        )
        self.tree = ttk.Treeview(
            tree_frame,
            columns=self.cols,
            show="headings",
            height=18,
            displaycolumns=self.display_cols,
        )
        self.tree.tag_configure("light_red", background="#FF9999")
        self.tree.tag_configure("kit_row", background="#D8F5D0")
        self.tree.tag_configure("module_row", background="#D5ECFF")

        headers = {
            "unique_id": lang.t("stock_inv.unique_id", "Unique ID"),
            "code": lang.t("stock_inv.code", "Code"),
            "description": lang.t("stock_inv.description", "Description"),
            "type": lang.t("stock_inv.type", "Type"),
            "scenario": lang.t("stock_inv.scenario", "Scenario"),
            "kit_number": lang.t("stock_inv.kit_number", "Kit Number"),
            "module_number": lang.t("stock_inv.module_number", "Module Number"),
            "current_stock": lang.t("stock_inv.current_stock", "Current Stock"),
            "exp_date": lang.t("stock_inv.expiry_date", "Expiry Date"),
            "physical_qty": lang.t("stock_inv.physical_qty", "Physical Quantity"),
            "updated_exp_date": lang.t(
                "stock_inv.updated_exp_date", "Updated Expiry Date"
            ),
            "discrepancy": lang.t("stock_inv.discrepancy", "Discrepancy"),
            "remarks": lang.t("stock_inv.remarks", "Remarks"),
            "std_qty": lang.t("stock_inv.std_qty", "Standard Qty"),
        }
        widths = {
            "unique_id": 0,
            "code": 110,
            "description": 360,
            "type": 80,
            "scenario": 110,
            "kit_number": 110,
            "module_number": 120,
            "current_stock": 90,
            "exp_date": 100,
            "physical_qty": 110,
            "updated_exp_date": 120,
            "discrepancy": 90,
            "remarks": 180,
            "std_qty": 80,
        }
        for c in self.cols:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=widths.get(c, 100), anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        btn_frame = tk.Frame(self, bg="#F5F5F5")
        btn_frame.pack(pady=5, fill="x")
        actions = [
            (
                "Save Adjustments",
                "#27AE60",
                self.save_all,
                "normal" if self.role in ["admin", "manager"] else "disabled",
            ),
            ("Clear All", "#7F8C8D", self.clear_form, "normal"),
            (
                "Export to Excel",
                "#3B82F6",
                self.export_to_excel,
                "normal" if self.role in ["admin", "manager"] else "disabled",
            ),
        ]
        for txt, bg, cmd, st in actions:
            tk.Button(
                btn_frame,
                text=lang.t(f"stock_inv.{txt.lower().replace(' ', '_')}", txt),
                bg=bg,
                fg="white",
                command=cmd,
                state=st,
            ).pack(side="left", padx=5)

        batch_frame = tk.Frame(self, bg="#F5F5F5")
        batch_frame.pack(pady=5, fill="x")
        self.batch_info_var = tk.StringVar()
        tk.Label(
            batch_frame,
            text=lang.t("stock_inv.batch_info", "Batch Information:"),
            bg="#F5F5F5",
        ).grid(row=0, column=0, padx=5, sticky="w")
        tk.Label(
            batch_frame,
            textvariable=self.batch_info_var,
            bg="#F5F5F5",
            wraplength=600,
            anchor="w",
        ).grid(row=0, column=1, padx=5, sticky="w")

        self.status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self.status_var, bg="#F5F5F5", anchor="w").pack(
            pady=5, anchor="w"
        )

        self.tree.bind("<Double-1>", self.start_edit)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.search_listbox.bind("<<ListboxSelect>>", self.on_search_select)
        self.tree.bind("<Tab>", self.on_tab_press)

        self.refresh_kit_dropdown()
        self.refresh_module_dropdown()
        self.on_inv_type_selected()

    def _is_valid_future_updated_exp(self, updated_exp_str: str) -> bool:
        """
        Return True only if updated_exp_str is a valid parsable future date.
        Empty / None returns False.
        """
        if not updated_exp_str:
            return False
        parsed = parse_expiry(updated_exp_str)
        if not parsed:
            return False
        d = parsed.date() if hasattr(parsed, "date") else parsed
        return d > datetime.now().date()

    def _highlight_missing_required_expiry(self):
        """
        Highlight rows (light_red) ONLY when:
          - Item requires expiry (remarks contains 'exp')
          - Physical quantity > 0
          - AND (updated_exp_date is missing OR invalid / not future)
        If updated_exp_date is present AND valid future -> remove highlight.
        Rows with phys qty 0 or blank are never highlighted.
        Other rows keep kit/module coloring.
        """
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if not vals:
                continue
            code = vals[1]
            phys = vals[9]
            updated = vals[10]
            t = (vals[3] or "").upper()
            phys_int = int(phys) if phys.isdigit() else 0
            requires = check_expiry_required(code)

            if requires and phys_int > 0:
                if self._is_valid_future_updated_exp(updated):
                    # Valid updated future expiry -> normal coloring
                    if t == "KIT":
                        self.tree.item(iid, tags=("kit_row",))
                    elif t == "MODULE":
                        self.tree.item(iid, tags=("module_row",))
                    else:
                        self.tree.item(iid, tags=())
                else:
                    # Missing or invalid updated future expiry -> highlight
                    self.tree.item(iid, tags=("light_red",))
            else:
                # Not requiring OR phys not positive -> normal coloring
                if t == "KIT":
                    self.tree.item(iid, tags=("kit_row",))
                elif t == "MODULE":
                    self.tree.item(iid, tags=("module_row",))
                else:
                    self.tree.item(iid, tags=())

    # ---------- Tab navigation ----------
    def on_tab_press(self, event):
        row_id = self.tree.focus()
        if not row_id:
            return "break"
        cols_order = ["physical_qty", "updated_exp_date", "remarks"]
        # Determine current column from focus not reliable; start next cycle
        # Always move to next editable column
        current_values = self.tree.item(row_id, "values")
        # Find which one is currently being edited by scanning focus widget - simplify: always start with first blank or cycle
        # Just start editing first editable if none editing
        self._begin_inline_edit(row_id, cols_order[0])
        return "break"

    # ---------- Context menu ----------
    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.ctx_row = row_id
        if self.ctx_menu:
            self.ctx_menu.destroy()
        self.ctx_menu = tk.Menu(self, tearoff=0)
        self.ctx_menu.add_command(
            label=lang.t("stock_inv.edit_physical", "Edit Physical Qty"),
            command=lambda: self._begin_inline_edit(row_id, "physical_qty"),
        )
        self.ctx_menu.add_command(
            label=lang.t("stock_inv.edit_expiry", "Edit Updated Expiry"),
            command=lambda: self._begin_inline_edit(row_id, "updated_exp_date"),
        )
        self.ctx_menu.add_command(
            label=lang.t("stock_inv.add_new_row", "Add New Row"),
            command=lambda: self.add_new_row_below(row_id),
        )
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(
            label=lang.t("stock_inv.clear_physical", "Clear Physical Qty"),
            command=lambda: self._clear_physical(row_id),
        )
        self.ctx_menu.add_command(
            label=lang.t("stock_inv.remove_row", "Remove Row"),
            command=lambda: (self.tree.delete(row_id), self._remove_state(row_id)),
        )
        self.ctx_menu.tk_popup(event.x_root, event.y_root)

    # ---------- State utility ----------
    def _remove_state(self, row_id):
        vals = self.tree.item(row_id, "values")
        if vals:
            uid = vals[0]
            self.user_row_states.pop(uid, None)
        self.base_physical_inputs.pop(row_id, None)

    def _update_state_from_row(self, row_id):
        vals = self.tree.item(row_id, "values")
        if not vals:
            return
        uid = vals[0]
        st = self.user_row_states.get(uid)
        if not st:
            return
        st["physical_qty"] = vals[9]
        st["updated_exp_date"] = vals[10]
        st["discrepancy"] = vals[11]
        st["remarks"] = vals[12]
        st["base_physical"] = self.base_physical_inputs.get(
            row_id, st.get("base_physical", 0)
        )

    # ---------- Insert row ----------
    def insert_tree_row(self, item, physical_qty=""):
        # Always blank by default
        iid = self.tree.insert(
            "",
            "end",
            values=(
                item["unique_id"],
                item["code"],
                item["description"],
                item["type"],
                item["scenario"],
                item.get("kit_number", "-----"),
                item.get("module_number", "-----"),
                item["current_stock"],
                item["exp_date"],
                "",  # physical qty blank
                "",  # updated expiry blank
                "",  # discrepancy blank
                "",  # remarks blank
                item["std_qty"],
            ),
        )
        t = item["type"].upper()
        if t == "KIT":
            self.tree.item(iid, tags=("kit_row",))
        elif t == "MODULE":
            self.tree.item(iid, tags=("module_row",))
        else:
            self.tree.item(iid, tags=())
        self.base_physical_inputs[iid] = 0
        self.user_row_states[item["unique_id"]] = {
            "unique_id": item["unique_id"],
            "code": item["code"],
            "description": item["description"],
            "type": item["type"],
            "scenario": item["scenario"],
            "kit_number": item.get("kit_number", "-----"),
            "module_number": item.get("module_number", "-----"),
            "current_stock": item["current_stock"],
            "exp_date": item["exp_date"],
            "physical_qty": "",
            "updated_exp_date": "",
            "discrepancy": "",
            "remarks": "",
            "std_qty": item["std_qty"],
            "base_physical": 0,
            "is_custom": item["current_stock"] == 0,
        }
        return True

    def add_new_row_below(self, row_id):
        vals = self.tree.item(row_id, "values")
        if not vals:
            return
        index = self.tree.index(row_id)
        original_uid = vals[0]
        new_uid = original_uid
        # Duplicate but keep blank physical / discrepancy as they appear
        iid = self.tree.insert("", index + 1, values=vals)
        t = (vals[3] or "").upper()
        if t == "KIT":
            self.tree.item(iid, tags=("kit_row",))
        elif t == "MODULE":
            self.tree.item(iid, tags=("module_row",))
        base_ph = 0
        self.base_physical_inputs[iid] = base_ph
        self.user_row_states[new_uid] = {
            "unique_id": new_uid,
            "code": vals[1],
            "description": vals[2],
            "type": vals[3],
            "scenario": vals[4],
            "kit_number": vals[5],
            "module_number": vals[6],
            "current_stock": vals[7],
            "exp_date": vals[8],
            "physical_qty": vals[9],
            "updated_exp_date": vals[10],
            "discrepancy": vals[11],
            "remarks": vals[12],
            "std_qty": vals[13],
            "base_physical": base_ph,
            "is_custom": True,
        }
        if self.mgmt_mode_var.get() == lang.t("stock_inv.management_in_box", "In-Box"):
            self.recompute_all_physical_quantities()
        self._highlight_missing_required_expiry()
        self.status_var.set(
            lang.t("stock_inv.added_new_row", "Added new row below {code}").format(
                code=vals[1]
            )
        )

    def _clear_physical(self, row_id):
        vals = self.tree.item(row_id, "values")
        if not vals:
            return
        uid = vals[0]
        self.base_physical_inputs[row_id] = 0
        self.tree.set(row_id, "physical_qty", "")
        self.tree.set(row_id, "discrepancy", "")
        st = self.user_row_states.get(uid)
        if st:
            st["physical_qty"] = ""
            st["discrepancy"] = ""
            st["base_physical"] = 0
        t = (vals[3] or "").upper()
        if t == "KIT":
            self.tree.item(row_id, tags=("kit_row",))
        elif t == "MODULE":
            self.tree.item(row_id, tags=("module_row",))
        if self.mgmt_mode_var.get() == lang.t("stock_inv.management_in_box", "In-Box"):
            self.recompute_all_physical_quantities()
        self._highlight_missing_required_expiry()

    # ---------- Inline edit ----------
    def _begin_inline_edit(self, row_id, col_key):
        """
        Handles inline editing for specified columns.
        This now correctly accepts '0' for KIT/MODULE rows and triggers the
        universal re-computation logic.
        """
        if col_key not in self.display_cols:
            return
        display_index = self.display_cols.index(col_key)
        bbox = self.tree.bbox(row_id, f"#{display_index+1}")
        if not bbox:
            return
        x, y, w, h = bbox
        current_val = self.tree.set(row_id, col_key)

        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        if current_val:
            entry.insert(0, current_val)
        entry.focus()

        def save_edit(evt=None):
            new_val = entry.get().strip()

            if col_key == "physical_qty":
                # Store the user's raw input as the base for calculations
                base_qty = int(new_val) if new_val.isdigit() else 0
                self.base_physical_inputs[row_id] = base_qty

                # ALWAYS trigger the single, correct recalculation function
                self.recompute_all_physical_quantities()

            elif col_key == "updated_exp_date":
                if new_val:
                    parsed = parse_expiry(new_val)
                    new_val = parsed.strftime("%Y-%m-%d") if parsed else ""
                self.tree.set(row_id, col_key, new_val)

            elif col_key == "remarks":
                self.tree.set(row_id, col_key, new_val[:255])

            self._update_state_from_row(row_id)
            self._highlight_missing_required_expiry()
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", lambda e: entry.destroy())

    def start_edit(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id or not col:
            return
        display_cols = list(self.display_cols)
        col_index = int(col.replace("#", "")) - 1
        if col_index >= len(display_cols):
            return
        col_key = display_cols[col_index]
        if col_key not in ("physical_qty", "updated_exp_date", "remarks"):
            return
        self._begin_inline_edit(row_id, col_key)

    # ---------- Filter / mode events ----------
    def on_scenario_selected(self, event=None):
        self.refresh_kit_dropdown()
        self.refresh_module_dropdown()
        self.rebuild_tree_preserving_state()

    def on_management_mode_change(self, event=None):
        mgmt_mode = self.mgmt_mode_var.get()
        on_shelf = mgmt_mode == lang.t("stock_inv.management_on_shelf", "On-Shelf")
        if on_shelf:
            self.kit_number_var.set(lang.t("stock_inv.all_kits", "All Kits"))
            self.module_number_var.set(lang.t("stock_inv.all_modules", "All Modules"))
            self.kit_number_cb.config(state="disabled")
            self.module_number_cb.config(state="disabled")
        else:
            self.kit_number_cb.config(state="readonly")
            self.module_number_cb.config(state="readonly")
        self.rebuild_tree_preserving_state()

    def on_filter_change(self, event=None):
        widget = event.widget if event else None
        mgmt_mode = self.mgmt_mode_var.get()
        on_shelf = mgmt_mode == lang.t("stock_inv.management_on_shelf", "On-Shelf")
        if not on_shelf and widget == getattr(self, "kit_number_cb", None):
            self.refresh_module_dropdown()
        if self.inv_type_var.get() == lang.t(
            "stock_inv.complete_inventory", "Complete Inventory"
        ):
            self.rebuild_tree_preserving_state()

    def on_inv_type_selected(self, event=None):
        inv_type = self.inv_type_var.get()
        if inv_type == lang.t("stock_inv.complete_inventory", "Complete Inventory"):
            self.rebuild_tree_preserving_state()
            self.search_frame.pack_forget()
            self.batch_info_var.set(
                lang.t(
                    "stock_inv.complete_info",
                    "Complete inventory: All items with current stock loaded.",
                )
            )
        else:
            # Partial inventory: clear tree and states to start fresh
            self.tree.delete(*self.tree.get_children())
            self.user_row_states.clear()
            self.base_physical_inputs.clear()
            self.search_frame.pack(pady=5, fill="x")
            self.status_var.set(
                lang.t(
                    "stock_inv.partial_info",
                    "Enter a code or description to search items or add a missing item.",
                )
            )
            self.batch_info_var.set(
                lang.t(
                    "stock_inv.partial_batch_info",
                    "Partial inventory: Search to add items or use 'Add Missing Item' for items not in stock.",
                )
            )

    # ---------- Search interactions ----------
    def on_search_keyrelease(self, event=None):
        query = self.code_entry.get().strip()
        self.search_listbox.delete(0, tk.END)
        if len(query) < 2:
            return
        results = self.fetch_search_results(query)
        for r in results:
            self.search_listbox.insert(tk.END, f"{r['code']} - {r['description']}")
        self.status_var.set(
            lang.t(
                "stock_inv.found_items", "Found {count} items matching '{query}'"
            ).format(count=len(results), query=query)
        )

    def select_first_search_result(self, event=None):
        if self.search_listbox.size() > 0:
            self.search_listbox.selection_set(0)
            self.on_search_select()

    def on_search_select(self, event=None):
        sel = self.search_listbox.curselection()
        if not sel:
            return
        line = self.search_listbox.get(sel[0])
        code = line.split(" - ")[0]
        self.capture_current_rows()
        items = self.fetch_items_in_stock(code)
        if not items:
            custom_popup(
                self,
                lang.t("dialog_titles.info", "Info"),
                lang.t(
                    "stock_inv.no_stock",
                    "No stock available for code {code}. Use 'Add Missing Item' to add it.",
                ).format(code=code),
                "info",
            )
            return
        added = 0
        for rec in items:
            if rec["unique_id"] not in self.user_row_states:
                self.insert_tree_row(rec, physical_qty="")
                added += 1
        self._highlight_missing_required_expiry()
        self.status_var.set(
            lang.t("stock_inv.added_items", "Added items for code {code}").format(
                code=code
            )
        )
        self.batch_info_var.set(
            lang.t(
                "stock_inv.added_lines",
                "Added {count} lines for code {code} including components.",
            ).format(count=added, code=code)
        )
        self.code_entry.delete(0, tk.END)
        self.search_listbox.delete(0, tk.END)

    # ---------- Add missing item ----------
    def add_missing_item(self):
        dialog = tk.Toplevel(self)
        dialog.title(lang.t("stock_inv.add_missing_item", "Add Missing Item"))
        dialog.geometry("430x620")
        dialog.transient(self)
        dialog.grab_set()

        tk.Label(dialog, text=lang.t("stock_inv.scenario", "Scenario:")).pack(pady=4)
        scenario_values = list(self.scenario_map.values())
        scen_cb_var = tk.StringVar(
            value=(
                self.scenario_var.get()
                if self.scenario_var.get()
                != lang.t("stock_inv.all_scenarios", "All Scenarios")
                else (scenario_values[0] if scenario_values else "")
            )
        )
        scen_cb = ttk.Combobox(
            dialog,
            textvariable=scen_cb_var,
            values=scenario_values,
            state="readonly",
            width=30,
        )
        scen_cb.pack(pady=2)

        tk.Label(dialog, text=lang.t("stock_inv.kit_code", "Kit Code:")).pack(pady=4)
        kit_code_var = tk.StringVar()
        tk.Entry(dialog, textvariable=kit_code_var, width=30).pack()

        tk.Label(dialog, text=lang.t("stock_inv.module_code", "Module Code:")).pack(
            pady=4
        )
        module_code_var = tk.StringVar()
        tk.Entry(dialog, textvariable=module_code_var, width=30).pack()

        tk.Label(
            dialog, text=lang.t("stock_inv.code", "Item Code - Description:")
        ).pack(pady=4)
        code_var = tk.StringVar()
        code_entry = tk.Entry(dialog, textvariable=code_var, width=30)
        code_entry.pack()

        code_listbox = tk.Listbox(dialog, height=6, width=40)
        code_listbox.pack(pady=4)

        tk.Label(dialog, text=lang.t("stock_inv.std_qty", "Standard Quantity:")).pack(
            pady=4
        )
        std_qty_var = tk.StringVar(value="0")
        tk.Label(dialog, textvariable=std_qty_var, width=30).pack()

        tk.Label(
            dialog, text=lang.t("stock_inv.physical_qty", "Physical Quantity:")
        ).pack(pady=4)
        physical_qty_var = tk.StringVar(value="")  # blank default
        tk.Entry(dialog, textvariable=physical_qty_var, width=30).pack()

        tk.Label(
            dialog, text=lang.t("stock_inv.expiry_date", "Expiry (MM/YYYY):")
        ).pack(pady=4)
        exp_date_var = tk.StringVar()
        tk.Entry(dialog, textvariable=exp_date_var, width=30).pack()

        tk.Label(dialog, text=lang.t("stock_inv.kit_number", "Kit Number:")).pack(
            pady=4
        )
        kit_number_var = tk.StringVar(
            value=(
                self.kit_number_var.get()
                if self.kit_number_var.get() != lang.t("stock_inv.all_kits", "All Kits")
                else ""
            )
        )
        tk.Entry(dialog, textvariable=kit_number_var, width=30).pack()

        tk.Label(dialog, text=lang.t("stock_inv.module_number", "Module Number:")).pack(
            pady=4
        )
        module_number_var = tk.StringVar(
            value=(
                self.module_number_var.get()
                if self.module_number_var.get()
                != lang.t("stock_inv.all_modules", "All Modules")
                else ""
            )
        )
        tk.Entry(dialog, textvariable=module_number_var, width=30).pack()

        def update_code_list(event=None):
            query = code_var.get().strip()
            code_listbox.delete(0, tk.END)
            if len(query) >= 2:
                results = self.fetch_search_results(query)
                for res in results:
                    code_listbox.insert(tk.END, f"{res['code']} - {res['description']}")
            scenario_sel = scen_cb_var.get()
            code_only = (
                code_var.get().split(" - ")[0]
                if " - " in code_var.get()
                else code_var.get()
            )
            if code_only:
                std_qty_var.set(str(get_std_qty(code_only, scenario_sel)))

        def on_code_pick(evt=None):
            sel = code_listbox.curselection()
            if not sel:
                return
            val = code_listbox.get(sel[0])
            code_var.set(val.split(" - ")[0])
            update_code_list()

        code_entry.bind("<KeyRelease>", update_code_list)
        code_listbox.bind("<<ListboxSelect>>", on_code_pick)
        code_listbox.bind("<Double-1>", on_code_pick)
        scen_cb.bind("<<ComboboxSelected>>", update_code_list)

        def submit_new():
            scenario_name = scen_cb_var.get()
            if not scenario_name:
                custom_popup(
                    self,
                    lang.t("dialog_titles.error", "Error"),
                    lang.t("stock_inv.no_scenario", "Please select a scenario."),
                    "error",
                )
                return
            scenario_id = None
            for sid, nm in self.scenario_map.items():
                if nm == scenario_name:
                    scenario_id = sid
                    break
            if not scenario_id:
                custom_popup(
                    self,
                    lang.t("dialog_titles.error", "Error"),
                    lang.t("stock_inv.no_scenario", "Scenario not found."),
                    "error",
                )
                return
            item_code = code_var.get().strip()
            if not item_code:
                custom_popup(
                    self,
                    lang.t("dialog_titles.error", "Error"),
                    lang.t("stock_inv.no_code", "Please select a code."),
                    "error",
                )
                return
            physical = physical_qty_var.get().strip()
            if not physical or not physical.isdigit() or int(physical) <= 0:
                custom_popup(
                    self,
                    lang.t("dialog_titles.error", "Error"),
                    lang.t(
                        "stock_inv.invalid_qty",
                        "Physical quantity must be a positive integer.",
                    ),
                    "error",
                )
                return
            physical = int(physical)
            exp_raw = exp_date_var.get().strip()
            # If item requires expiry -> must supply valid future expiry
            if check_expiry_required(item_code):
                if not validate_expiry_for_save(item_code, exp_raw):
                    custom_popup(
                        self,
                        lang.t("dialog_titles.error", "Error"),
                        lang.t(
                            "stock_inv.invalid_expiry",
                            "A valid future expiry date is required.",
                        ),
                        "error",
                    )
                    return

            parsed_exp = parse_expiry(exp_raw)
            exp_iso = parsed_exp.strftime("%Y-%m-%d") if parsed_exp else "None"
            std_qty_val = int(std_qty_var.get()) if std_qty_var.get().isdigit() else 0

            mgmt_mode = self.mgmt_mode_var.get()
            force_box = mgmt_mode == lang.t("stock_inv.management_in_box", "In-Box")
            had_box = (
                kit_number_var.get().strip()
                or module_number_var.get().strip()
                or force_box
            )
            treecode = (
                get_treecode(
                    scenario_id,
                    kit_code_var.get().strip() or None,
                    module_code_var.get().strip() or None,
                    item_code,
                )
                if had_box
                else None
            )
            unique_id = construct_unique_id(
                scenario_id=scenario_id,
                kit_code=kit_code_var.get().strip() or None,
                module_code=module_code_var.get().strip() or None,
                item_code=item_code,
                std_qty=std_qty_val,
                exp_date=exp_iso,
                kit_number=kit_number_var.get().strip() or None,
                module_number=module_number_var.get().strip() or None,
                force_box_format=force_box,
                treecode=treecode,
            )

            description = get_active_designation(item_code)
            item_type = get_item_type(item_code)
            record = {
                "unique_id": unique_id,
                "code": item_code,
                "description": description,
                "type": item_type,
                "scenario": scenario_name,
                "kit_number": kit_number_var.get().strip() or "-----",
                "module_number": module_number_var.get().strip() or "-----",
                "current_stock": 0,
                "exp_date": exp_iso if exp_iso != "None" else "",
                "std_qty": std_qty_val,
            }
            self.insert_tree_row(record, physical_qty=str(physical))
            self._highlight_missing_required_expiry()
            self.batch_info_var.set(
                lang.t(
                    "stock_inv.added_new_item_info",
                    "Added new item {code} with physical quantity {qty}",
                ).format(code=item_code, qty=physical)
            )
            self.status_var.set(
                lang.t("stock_inv.added_new_item", "Added new item {code}").format(
                    code=item_code
                )
            )
            dialog.destroy()

    def _zero_old_expiry_line(self, cur, unique_id, has_discrepancy):
        """
        Force the existing stock_data line to zero final stock:
            final_qty = 0  ==> set qty_out = qty_in and discrepancy = 0 (if column exists)
        """
        if has_discrepancy:
            cur.execute(
                """
                UPDATE stock_data
                   SET qty_out = qty_in,
                       discrepancy = 0
                 WHERE unique_id = ?
            """,
                (unique_id,),
            )
        else:
            cur.execute(
                """
                UPDATE stock_data
                   SET qty_out = qty_in
                 WHERE unique_id = ?
            """,
                (unique_id,),
            )

    def _zero_batch_via_qty_out(self, cur, unique_id, has_discrepancy):
        """
        Force a batch to zero stock using qty_out (preferred):
            qty_out = qty_in
            discrepancy = 0 (if exists)
        Safe idempotent: running multiple times keeps it zeroed.
        """
        if has_discrepancy:
            cur.execute(
                """
                UPDATE stock_data
                   SET qty_out = qty_in,
                       discrepancy = 0
                 WHERE unique_id = ?
            """,
                (unique_id,),
            )
        else:
            cur.execute(
                """
                UPDATE stock_data
                   SET qty_out = qty_in
                 WHERE unique_id = ?
            """,
                (unique_id,),
            )

    def save_all(self):
        """
        Save inventory adjustments using "Bake-In & Reset" strategy.
        
        Strategy:
        - Convert all adjustments into qty_in/qty_out movements
        - Reset discrepancy = 0 after baking in
        - Database triggers auto-calculate final_qty = qty_in - qty_out + 0
        
        Handles:
        - Complete & Partial inventory types
        - 6-segment (on-shelf) & 8-segment (in-box) unique_ids
        - Expiry changes (close old batch, create new)
        - New items not in stock (insert with qty_in = physical)
        
        Validates:
        - Items requiring expiry have valid future updated_exp_date
        - Only admin/manager roles can save
        """

        # Role validation
        if self.role.lower() not in ["admin", "manager"]:
            custom_popup(
                self,
                lang.t("dialog_titles.error", "Error"),
                lang.t(
                    "stock_inv.restricted",
                    "Only admin or manager roles can save changes.",
                ),
                "error",
            )
            return

        # Check if there are rows to save
        rows = self.tree.get_children()
        if not rows:
            custom_popup(
                self,
                lang.t("dialog_titles.error", "Error"),
                lang.t("stock_inv.no_rows", "No rows to save."),
                "error",
            )
            return

        # Validate expiry dates for items requiring them
        blocking = []
        for rid in rows:
            vals = self.tree.item(rid, "values")
            code = vals[1]
            phys_str = vals[9]
            updated_exp = vals[10]
            phys_int = int(phys_str) if phys_str.isdigit() else 0
            if check_expiry_required(code) and phys_int > 0:
                if not self._is_valid_future_updated_exp(updated_exp):
                    blocking.append(code)
        if blocking:
            self._highlight_missing_required_expiry()
            custom_popup(
                self,
                lang.t("dialog_titles.error", "Error"),
                lang.t(
                    "stock_inv.invalid_expiry",
                    "Valid future updated expiry dates are required for items: {items}",
                ).format(items=", ".join(sorted(set(blocking)))),
                "error",
            )
            return

        # Generate document number for transactions
        doc_number = self.generate_document_number()

        # Connect to database
        conn = connect_db()
        if conn is None:
            custom_popup(
                self,
                lang.t("dialog_titles.error", "Error"),
                lang.t("stock_inv.db_error", "Database connection failed"),
                "error",
            )
            return
        cur = conn.cursor()

        errors = []
        max_retries = 4

        def attempt(sql, params):
            """
            Execute SQL with retry logic for database locks.
            
            Args:
                sql (str): SQL query to execute
                params (tuple): Query parameters
                
            Returns:
                bool: True if successful, False if failed after retries
            """
            for attempt_i in range(max_retries):
                try:
                    cur.execute(sql, params)
                    return True
                except sqlite3.OperationalError as e:
                    if "locked" in str(e).lower() and attempt_i < max_retries - 1:
                        time.sleep(0.35 * (attempt_i + 1))
                        continue
                    errors.append(str(e))
                    return False

        # Determine inventory type abbreviation for logging
        inv_type_label = self.inv_type_var.get()
        inv_abbr = (
            "Complete INV"
            if inv_type_label
            == lang.t("stock_inv.complete_inventory", "Complete Inventory")
            else "Partial INV"
        )

        def final_qty(row):
            """
            Calculate final quantity matching trigger logic: qty_in - qty_out.
            
            Args:
                row (tuple): Stock data row with (qty_in, qty_out, ...)
                
            Returns:
                int: Final quantity calculated as qty_in - qty_out
            """
            return (row[0] or 0) - (row[1] or 0)

        def log_transaction(
            uid,
            exp,
            qty_in=None,
            qty_out=None,
            discrepancy=None,
            remarks=None,
            code=None,
            scen=None,
            kit=None,
            mod=None,
        ):
            """
            Log transaction to stock_transactions table with discrepancy for audit trail.
            
            Args:
                discrepancy: Variance (positive=surplus, negative=shortage)
            """
            # Only log if there's actual movement
            if not qty_in and not qty_out:
                return
            
            attempt(
                """
                INSERT INTO stock_transactions
                (Date, Time, document_number, unique_id, code, Description, Expiry_date, Batch_Number,
                    Scenario, Kit, Module, Qty_IN, IN_Type, Qty_Out, Out_Type,
                    Third_Party, End_User, Discrepancy, Remarks, Movement_Type)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
                (
                    datetime.today().strftime("%Y-%m-%d"),
                    datetime.now().strftime("%H:%M:%S"),
                    doc_number,
                    uid,
                    code,
                    get_active_designation(code),
                    exp or None,
                    None,
                    scen,  # scenario_id
                    kit if kit != "-----" else None,
                    mod if mod != "-----" else None,
                    qty_in if qty_in and qty_in > 0 else None,
                    inv_abbr if qty_in and qty_in > 0 else None,
                    qty_out if qty_out and qty_out > 0 else None,
                    inv_abbr if qty_out and qty_out > 0 else None,
                    None,
                    None,
                    discrepancy if discrepancy else None,
                    remarks if remarks else None,
                    "stock_inv",
                ),
            )

        try:
            cur.execute("BEGIN TRANSACTION")

            for rid in rows:
                vals = self.tree.item(rid, "values")
                (
                    unique_id,
                    code,
                    _desc,
                    _type,
                    scenario_name,
                    kit_number,
                    module_number,
                    _curr_stock_ui,
                    exp_date_orig,
                    phys_str,
                    updated_exp,
                    _disc_ui,
                    remarks,
                    std_qty_ui,
                ) = vals

                # Parse physical quantity
                physical = int(phys_str) if phys_str.isdigit() else None
                if physical is None or physical < 0:
                    # Log warning for data quality tracking
                    print(f"Warning: Skipping row with invalid physical quantity for item {code}: '{phys_str}'")
                    continue  # Skip rows without valid physical quantity

                # Parse unique_id components
                parsed = parse_inventory_unique_id(unique_id)
                scenario_id = parsed["scenario_id"]
                kit_code = parsed["kit_code"]
                module_code = parsed["module_code"]
                item_code = parsed["item_code"]
                old_exp = parsed["exp_date"] or ""
                new_exp = updated_exp or ""
                expiry_changed = new_exp and new_exp != old_exp
                std_qty_val = (
                    parsed["std_qty"]
                    if parsed["std_qty"] is not None
                    else (int(std_qty_ui) if str(std_qty_ui).isdigit() else 0)
                )

                # Fetch existing stock_data row
                cur.execute(
                    "SELECT qty_in, qty_out FROM stock_data WHERE unique_id = ?",
                    (unique_id,),
                )
                existing = cur.fetchone()

                # Determine if this is in-box format (8-segment unique_id)
                had_box = (
                    parsed["kit_number"]
                    or parsed["module_number"]
                    or len(unique_id.split("/")) >= 8
                )

                # Helper to insert new batch with qty_in = physical
                def insert_new_batch(new_uid, exp_val, qty_in_amount):
                    """
                    Insert new batch. Triggers calculate final_qty = qty_in - qty_out.
                    
                    Args:
                        new_uid (str): Unique identifier for the new batch
                        exp_val (str): Expiry date in ISO format
                        qty_in_amount (int): Physical quantity to insert as qty_in
                    """
                    cols = [
                        "unique_id",
                        "scenario",
                        "kit_number",
                        "module_number",
                        "kit",
                        "module",
                        "item",
                        "std_qty",
                        "qty_in",
                        "qty_out",
                        "exp_date",
                    ]
                    vals_ = [
                        new_uid,
                        scenario_name,
                        parsed["kit_number"],
                        parsed["module_number"],
                        kit_code,
                        module_code,
                        item_code,
                        std_qty_val,
                        qty_in_amount,  # qty_in = physical
                        0,  # qty_out = 0
                        exp_val,
                    ]
                    attempt(
                        f"INSERT INTO stock_data ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})",
                        tuple(vals_),
                    )

                # CASE 1: New row (item not in stock)
                if not existing:
                    if physical > 0:
                        # Construct unique_id with correct expiry
                        target_exp = new_exp if new_exp else old_exp
                        treecode = (
                            get_treecode(scenario_id, kit_code, module_code, item_code)
                            if had_box
                            else None
                        )
                        new_uid = construct_unique_id(
                            scenario_id=scenario_id,
                            kit_code=kit_code,
                            module_code=module_code,
                            item_code=item_code,
                            std_qty=std_qty_val,
                            exp_date=target_exp,
                            kit_number=parsed["kit_number"] if had_box else None,
                            module_number=parsed["module_number"] if had_box else None,
                            force_box_format=had_box,
                            treecode=treecode,
                        )
                        # Insert with qty_in = physical
                        insert_new_batch(new_uid, target_exp, physical)
                        # Log as Qty_IN movement with discrepancy
                        log_transaction(
                            new_uid,
                            target_exp,
                            qty_in=physical,
                            discrepancy=physical,
                            code=code,
                            scen=scenario_id,
                            kit=kit_number,
                            mod=module_number,
                            remarks=remarks or f"New batch from {inv_abbr}",
                        )
                    continue

                # Get current final quantity
                old_qty_in = existing[0] or 0
                old_qty_out = existing[1] or 0
                old_final = final_qty(existing)

                # CASE 2: Existing row with expiry change - Smart Split Logic
                if expiry_changed:
                    # Determine split strategy based on physical vs old_final
                    
                    if physical < old_final:
                        # SCENARIO 1: SPLIT - Keep difference in old batch
                        qty_to_move = old_final - physical  # Amount moving to new batch
                        new_qty_out = old_qty_out + qty_to_move
                        
                        # Update old batch (keeps physical amount with old expiry)
                        attempt(
                            "UPDATE stock_data SET qty_out = ? WHERE unique_id = ?",
                            (new_qty_out, unique_id),
                        )
                        
                        # Log removal from old batch
                        log_transaction(
                            unique_id,
                            old_exp,
                            qty_out=qty_to_move,
                            discrepancy=-qty_to_move,
                            code=code,
                            scen=scenario_id,
                            kit=kit_number,
                            mod=module_number,
                            remarks=f"Split: moving {qty_to_move} units to new expiry {new_exp}",
                        )
                        
                        # Create new batch with the moved amount
                        treecode = (
                            get_treecode(scenario_id, kit_code, module_code, item_code)
                            if had_box
                            else None
                        )
                        new_uid = construct_unique_id(
                            scenario_id=scenario_id,
                            kit_code=kit_code,
                            module_code=module_code,
                            item_code=item_code,
                            std_qty=std_qty_val,
                            exp_date=new_exp,
                            kit_number=parsed["kit_number"] if had_box else None,
                            module_number=parsed["module_number"] if had_box else None,
                            force_box_format=had_box,
                            treecode=treecode,
                        )
                        insert_new_batch(new_uid, new_exp, qty_to_move)
                        
                        # Log new batch creation
                        log_transaction(
                            new_uid,
                            new_exp,
                            qty_in=qty_to_move,
                            discrepancy=qty_to_move,
                            code=code,
                            scen=scenario_id,
                            kit=kit_number,
                            mod=module_number,
                            remarks=f"New batch from split of {old_exp}",
                        )
                        
                    elif physical == old_final:
                        # SCENARIO 2: MOVE ALL - Close old batch, create new with same qty
                        new_qty_out = old_qty_out + old_final
                        
                        # Close old batch
                        attempt(
                            "UPDATE stock_data SET qty_out = ? WHERE unique_id = ?",
                            (new_qty_out, unique_id),
                        )
                        
                        # Log closing
                        log_transaction(
                            unique_id,
                            old_exp,
                            qty_out=old_final,
                            discrepancy=-old_final,
                            code=code,
                            scen=scenario_id,
                            kit=kit_number,
                            mod=module_number,
                            remarks=f"Closed: moving all stock to new expiry {new_exp}",
                        )
                        
                        # Create new batch
                        treecode = (
                            get_treecode(scenario_id, kit_code, module_code, item_code)
                            if had_box
                            else None
                        )
                        new_uid = construct_unique_id(
                            scenario_id=scenario_id,
                            kit_code=kit_code,
                            module_code=module_code,
                            item_code=item_code,
                            std_qty=std_qty_val,
                            exp_date=new_exp,
                            kit_number=parsed["kit_number"] if had_box else None,
                            module_number=parsed["module_number"] if had_box else None,
                            force_box_format=had_box,
                            treecode=treecode,
                        )
                        insert_new_batch(new_uid, new_exp, physical)
                        
                        # Log new batch
                        log_transaction(
                            new_uid,
                            new_exp,
                            qty_in=physical,
                            discrepancy=physical,
                            code=code,
                            scen=scenario_id,
                            kit=kit_number,
                            mod=module_number,
                            remarks=f"New batch replacing {old_exp}",
                        )
                        
                    else:  # physical > old_final
                        # SCENARIO 3: MOVE + ADD - Close old batch, create new with surplus
                        surplus = physical - old_final
                        new_qty_out = old_qty_out + old_final
                        
                        # Close old batch
                        attempt(
                            "UPDATE stock_data SET qty_out = ? WHERE unique_id = ?",
                            (new_qty_out, unique_id),
                        )
                        
                        # Log closing
                        log_transaction(
                            unique_id,
                            old_exp,
                            qty_out=old_final,
                            discrepancy=-old_final,
                            code=code,
                            scen=scenario_id,
                            kit=kit_number,
                            mod=module_number,
                            remarks=f"Closed: moving stock to new expiry {new_exp} with surplus",
                        )
                        
                        # Create new batch with surplus
                        treecode = (
                            get_treecode(scenario_id, kit_code, module_code, item_code)
                            if had_box
                            else None
                        )
                        new_uid = construct_unique_id(
                            scenario_id=scenario_id,
                            kit_code=kit_code,
                            module_code=module_code,
                            item_code=item_code,
                            std_qty=std_qty_val,
                            exp_date=new_exp,
                            kit_number=parsed["kit_number"] if had_box else None,
                            module_number=parsed["module_number"] if had_box else None,
                            force_box_format=had_box,
                            treecode=treecode,
                        )
                        insert_new_batch(new_uid, new_exp, physical)
                        
                        # Log new batch with surplus
                        log_transaction(
                            new_uid,
                            new_exp,
                            qty_in=physical,
                            discrepancy=physical,
                            code=code,
                            scen=scenario_id,
                            kit=kit_number,
                            mod=module_number,
                            remarks=f"New batch with surplus +{surplus} from {old_exp}",
                        )
                    
                    continue  # Skip to next row

                # CASE 3: Existing row (no expiry change)
                # Calculate adjustment needed
                adjustment = physical - old_final
                
                if adjustment > 0:
                    # Need to ADD stock - increase qty_in
                    new_qty_in = old_qty_in + adjustment
                    attempt(
                        "UPDATE stock_data SET qty_in = ? WHERE unique_id = ?",
                        (new_qty_in, unique_id),
                    )
                    # Log as Qty_IN movement
                    log_transaction(
                        unique_id,
                        old_exp,
                        qty_in=adjustment,
                        discrepancy=adjustment,
                        code=code,
                        scen=scenario_id,
                        kit=kit_number,
                        mod=module_number,
                        remarks=remarks,
                    )
                    
                elif adjustment < 0:
                    # Need to REMOVE stock - increase qty_out
                    new_qty_out = old_qty_out + abs(adjustment)
                    attempt(
                        "UPDATE stock_data SET qty_out = ? WHERE unique_id = ?",
                        (new_qty_out, unique_id),
                    )
                    # Log as Qty_Out movement
                    log_transaction(
                        unique_id,
                        old_exp,
                        qty_out=abs(adjustment),
                        discrepancy=adjustment,
                        code=code,
                        scen=scenario_id,
                        kit=kit_number,
                        mod=module_number,
                        remarks=remarks,
                    )
                    
                else:
                    # No change needed
                    pass

            conn.commit()

            if errors:
                custom_popup(
                    self,
                    lang.t("dialog_titles.error", "Error"),
                    "Some rows failed:\n" + "\n".join(errors),
                    "error",
                )
            else:
                custom_popup(
                    self,
                    lang.t("dialog_titles.success", "Success"),
                    lang.t(
                        "stock_inv.success", "Inventory adjustments saved successfully."
                    ),
                    "info",
                )
                self.user_row_states.clear()
                self.base_physical_inputs.clear()
                self.export_to_excel(document_number=doc_number)
                self.clear_form()

        except Exception as e:
            conn.rollback()
            custom_popup(
                self,
                lang.t("dialog_titles.error", "Error"),
                lang.t("stock_inv.save_failed", "Failed to save: {error}").format(
                    error=str(e)
                ),
                "error",
            )
        finally:
            cur.close()
            conn.close()

    # ---------- Clear form ----------
    def clear_form(self):
        self.tree.delete(*self.tree.get_children())
        self.code_entry.delete(0, tk.END)
        self.search_listbox.delete(0, tk.END)
        self.inv_type_var.set(
            lang.t("stock_inv.complete_inventory", "Complete Inventory")
        )
        self.scenario_var.set(lang.t("stock_inv.all_scenarios", "All Scenarios"))
        self.mgmt_mode_var.set(lang.t("stock_inv.management_all", "All"))
        self.kit_number_var.set(lang.t("stock_inv.all_kits", "All Kits"))
        self.module_number_var.set(lang.t("stock_inv.all_modules", "All Modules"))
        self.refresh_kit_dropdown()
        self.refresh_module_dropdown()
        self.search_frame.pack_forget()
        self.batch_info_var.set("")
        self.status_var.set("")
        self.user_row_states.clear()
        self.base_physical_inputs.clear()
        self.on_inv_type_selected()


# ---------- Standalone run ----------
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Stock Inventory Adjustment")
    app = type("App", (), {})()
    app.role = "admin"
    StockInventory(root, app, role="admin")
    root.geometry("1550x900")
    root.mainloop()
