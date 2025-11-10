import tkinter as tk
from tkinter import ttk
import sqlite3
from datetime import datetime
from db import connect_db
from language_manager import lang as lang_lang   # (kept original alias style)
from popup_utils import custom_popup, custom_askyesno
import time

# -----------------------------------------------------------------------------------
# Configuration / Constants
# -----------------------------------------------------------------------------------
BG_MAIN          = "#F0F4F8"
BG_PANEL         = "#FFFFFF"
COLOR_PRIMARY    = "#2C3E50"
COLOR_ACCENT     = "#2563EB"
COLOR_BORDER     = "#D0D7DE"
COLOR_ROW_ALT    = "#F7FAFC"
COLOR_ROW_NORM   = "#FFFFFF"
BTN_BG_ADD       = "#27AE60"
BTN_BG_EDIT      = "#2980B9"
BTN_BG_DELETE    = "#C0392B"
BTN_BG_DISABLED  = "#95A5A6"
MAX_SCENARIOS    = 15
SIMILARITY_THRESHOLD = 0.70  # 70%

# Roles / symbols that are NOT allowed to modify scenarios
RESTRICTED_ROLES = {"manager", "supervisor", "~", "$"}

# -----------------------------------------------------------------------------------
# Levenshtein-based similarity (case-insensitive)
# -----------------------------------------------------------------------------------
def levenshtein_distance(a: str, b: str) -> int:
    a = a.lower()
    b = b.lower()
    if a == b:
        return 0
    la, lb = len(a), len(b)
    if la == 0:
        return lb
    if lb == 0:
        return la
    prev = list(range(lb + 1))
    for i, ca in enumerate(a, start=1):
        curr = [i]
        for j, cb in enumerate(b, start=1):
            cost = 0 if ca == cb else 1
            curr.append(min(
                prev[j] + 1,      # deletion
                curr[j-1] + 1,    # insertion
                prev[j-1] + cost  # substitution
            ))
        prev = curr
    return prev[-1]

def similarity_ratio(a: str, b: str) -> float:
    a = (a or "").strip().lower()
    b = (b or "").strip().lower()
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    dist = levenshtein_distance(a, b)
    return 1.0 - (dist / max(len(a), len(b)))


class Scenarios(tk.Frame):
    """
    Scenario Management UI with restricted modification:
      - Max 15 scenarios
      - Unique scenario names (case-insensitive)
      - Warn if new/edited name >= 70% similar to existing different names
      - Target population must be integer or blank
      - Table shows up to 15 rows (blank fillers)
      - Local time display for last updated
      - Users with role symbol "~" (manager) or "$" (supervisor) CANNOT add/edit/delete
        (Also covers canonical names 'manager' / 'supervisor')
    """
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG_MAIN)
        self.app = app
        self.role = getattr(app, "role", "user") or "user"
        self.pack(fill="both", expand=True)
        self._configure_styles()
        self._build_ui()
        self.load_data()

    # --------------------------------------------------------------------------------
    # Permission helper
    # --------------------------------------------------------------------------------
    def _can_modify(self) -> bool:
        r = (self.role or "").lower()
        return r not in RESTRICTED_ROLES

    # --------------------------------------------------------------------------------
    # Styles
    # --------------------------------------------------------------------------------
    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Sc.Treeview",
            background=BG_PANEL,
            fieldbackground=BG_PANEL,
            foreground=COLOR_PRIMARY,
            rowheight=26,
            bordercolor=COLOR_BORDER,
            relief="flat",
            font=("Helvetica", 10)
        )
        style.map("Sc.Treeview",
                  background=[("selected", COLOR_ACCENT)],
                  foreground=[("selected", "#FFFFFF")])

        style.configure(
            "Sc.Treeview.Heading",
            background="#E5E8EB",
            foreground=COLOR_PRIMARY,
            font=("Helvetica", 11, "bold"),
            relief="flat"
        )

        style.configure(
            "Sc.TButton",
            font=("Helvetica", 10, "bold"),
            padding=6
        )
        style.configure("Sc.TEntry", font=("Helvetica", 10))
        style.configure("Sc.TCombobox", font=("Helvetica", 10))

    # --------------------------------------------------------------------------------
    # UI Build
    # --------------------------------------------------------------------------------
    def _build_ui(self):
        # Title
        tk.Label(
            self,
            text=lang_lang.t("scenarios.title", "Manage Scenarios"),
            font=("Helvetica", 20, "bold"),
            bg=BG_MAIN,
            fg=COLOR_PRIMARY,
            anchor="w",
            justify="left"
        ).pack(fill="x", pady=(8, 4), padx=12)

        # Last updated label
        self.last_updated_var = tk.StringVar(value="")
        tk.Label(
            self,
            textvariable=self.last_updated_var,
            font=("Helvetica", 10, "italic"),
            bg=BG_MAIN,
            fg="#7F8C8D",
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(0, 8))

        # Table frame with border
        table_outer = tk.Frame(self, bg=COLOR_BORDER, bd=1, relief="solid")
        table_outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        columns = (
            "scenario_id",
            "name",
            "activity_type",
            "target_population",
            "stock_location",
            "responsible_person"
        )
        headers = {
            "scenario_id":        lang_lang.t("scenarios.id", "ID"),
            "name":               lang_lang.t("scenarios.name", "Scenario Name"),
            "activity_type":      lang_lang.t("scenarios.activity_type", "Activity Type"),
            "target_population":  lang_lang.t("scenarios.target_population", "Target Population"),
            "stock_location":     lang_lang.t("scenarios.stock_location", "Stock Location"),
            "responsible_person": lang_lang.t("scenarios.responsible_person", "Person Responsible")
        }

        self.tree = ttk.Treeview(
            table_outer,
            columns=columns,
            show="headings",
            height=MAX_SCENARIOS,
            style="Sc.Treeview"
        )

        self.tree.column("scenario_id", width=70, anchor="w")
        self.tree.column("name", width=180, anchor="w")
        self.tree.column("activity_type", width=130, anchor="w")
        self.tree.column("target_population", width=130, anchor="e")
        self.tree.column("stock_location", width=180, anchor="w")
        self.tree.column("responsible_person", width=180, anchor="w")
        for col in columns:
            self.tree.heading(col, text=headers[col])

        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(table_outer, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        # Buttons
        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill="x", padx=12, pady=(0, 6))

        can_modify = self._can_modify()

        self.btn_add = tk.Button(
            btn_frame, text=lang_lang.t("scenarios.add", "Add Scenario"),
            bg=BTN_BG_ADD if can_modify else BTN_BG_DISABLED,
            fg="#FFFFFF", relief="flat",
            activebackground="#1E874B" if can_modify else BTN_BG_DISABLED,
            command=(self.add_scenario if can_modify else self._denied_popup)
        )
        self.btn_add.pack(side="left", padx=4)

        self.btn_edit = tk.Button(
            btn_frame, text=lang_lang.t("scenarios.edit", "Edit Scenario"),
            bg=BTN_BG_EDIT if can_modify else BTN_BG_DISABLED,
            fg="#FFFFFF", relief="flat",
            activebackground="#1F5D82" if can_modify else BTN_BG_DISABLED,
            command=(self.edit_scenario if can_modify else self._denied_popup)
        )
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = tk.Button(
            btn_frame, text=lang_lang.t("scenarios.delete", "Delete Scenario"),
            bg=BTN_BG_DELETE if can_modify else BTN_BG_DISABLED,
            fg="#FFFFFF", relief="flat",
            activebackground="#962D22" if can_modify else BTN_BG_DISABLED,
            command=(self.delete_scenario if can_modify else self._denied_popup)
        )
        self.btn_delete.pack(side="left", padx=4)

        # Limit note
        tk.Label(
            self,
            text=lang_lang.t("scenarios.limit_note", "Note: Maximum 15 active scenarios allowed."),
            font=("Helvetica", 9, "italic"),
            bg=BG_MAIN,
            fg="#7F8C8D",
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(0, 4))

        # If restricted, show an info label
        if not can_modify:
            tk.Label(
                self,
                text=lang_lang.t("scenarios.read_only_notice",
                                 "Read-only: Your role is not permitted to modify scenarios."),
                font=("Helvetica", 9, "italic"),
                bg=BG_MAIN,
                fg="#A855F7",
                anchor="w",
                justify="left"
            ).pack(fill="x", padx=12, pady=(2, 6))

    def _denied_popup(self):
        custom_popup(
            self,
            lang_lang.t("dialog_titles.restricted", "Restricted"),
            lang_lang.t("scenarios.restricted_msg",
                        "Your role is not permitted to modify scenarios."),
            "warning"
        )

    # --------------------------------------------------------------------------------
    # Local time formatting helper
    # --------------------------------------------------------------------------------
    def _utc_to_local_display(self, ts: str) -> str:
        if not ts or ts == "N/A":
            return "N/A"
        try:
            dt_obj = None
            for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
                try:
                    dt_obj = datetime.strptime(ts, fmt)
                    break
                except ValueError:
                    continue
            if dt_obj is None:
                return ts
            epoch_utc = dt_obj.timestamp()
            offset = (datetime.fromtimestamp(epoch_utc) - datetime.utcfromtimestamp(epoch_utc))
            local_dt = dt_obj + offset
            return local_dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return ts

    # --------------------------------------------------------------------------------
    # Data Loading
    # --------------------------------------------------------------------------------
    def load_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = connect_db()
        if conn is None:
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.db_error", "Database connection failed"), "error")
            return
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("""
                SELECT scenario_id, name, activity_type, target_population, stock_location, responsible_person, created_at
                FROM scenarios
                ORDER BY scenario_id
            """)
            rows = cur.fetchall()
            for idx, r in enumerate(rows):
                vals = (
                    r["scenario_id"],
                    r["name"],
                    r["activity_type"],
                    r["target_population"] if r["target_population"] is not None else "",
                    r["stock_location"],
                    r["responsible_person"]
                )
                tag = "alt" if idx % 2 else "norm"
                self.tree.insert("", "end", values=vals, tags=(tag,))
            filler = MAX_SCENARIOS - len(rows)
            for i in range(filler):
                tag = "alt" if (len(rows) + i) % 2 else "norm"
                self.tree.insert("", "end", values=("", "", "", "", "", ""), tags=(tag,))

            self.tree.tag_configure("norm", background=COLOR_ROW_NORM)
            self.tree.tag_configure("alt", background=COLOR_ROW_ALT)

            last_ts = max([row["created_at"] for row in rows if row["created_at"]]) if rows else None
            last_local = self._utc_to_local_display(last_ts) if last_ts else "N/A"
            self.last_updated_var.set(
                lang_lang.t("scenarios.last_updated", "Last updated: {0}").format(last_local)
            )

        except sqlite3.Error as e:
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.last_updated_error",
                                     "Error fetching last updated: {0}").format(str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # --------------------------------------------------------------------------------
    # Add (guarded)
    # --------------------------------------------------------------------------------
    def add_scenario(self):
        if not self._can_modify():
            self._denied_popup()
            return
        if not self._can_add_more():
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.limit_reached",
                                     "Maximum of 15 active scenarios reached. Adopt / edit an existing one."),
                         "error")
            return
        self._open_form(edit=False)

    def _can_add_more(self) -> bool:
        conn = connect_db()
        if conn is None:
            return False
        cur = conn.cursor()
        try:
            cur.execute("SELECT COUNT(*) FROM scenarios")
            count = cur.fetchone()[0]
            return count < MAX_SCENARIOS
        finally:
            cur.close()
            conn.close()

    # --------------------------------------------------------------------------------
    # Edit (guarded)
    # --------------------------------------------------------------------------------
    def edit_scenario(self):
        if not self._can_modify():
            self._denied_popup()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.select_edit", "Please select a scenario to edit"),
                         "error")
            return
        values = self.tree.item(sel[0])["values"]
        if not values or values[0] == "":
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.select_edit", "Please select a scenario to edit"),
                         "error")
            return
        self._open_form(edit=True, scenario_data=values)

    # --------------------------------------------------------------------------------
    # Delete (guarded)
    # --------------------------------------------------------------------------------
    def delete_scenario(self):
        if not self._can_modify():
            self._denied_popup()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.select_delete", "Please select a scenario to delete"),
                         "error")
            return
        values = self.tree.item(sel[0])["values"]
        if not values or values[0] == "":
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.select_delete", "Please select a scenario to delete"),
                         "error")
            return
        scenario_id = values[0]
        ans = custom_askyesno(
            self,
            lang_lang.t("scenarios.confirm", "Confirm Delete"),
            lang_lang.t("scenarios.confirm_msg", "Are you sure you want to delete this scenario?")
        )
        if ans != "yes":
            return
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"),
                         lang_lang.t("scenarios.db_error", "Database connection failed"), "error")
            return
        cur = conn.cursor()
        try:
            cur.execute("DELETE FROM scenarios WHERE scenario_id = ?", (scenario_id,))
            conn.commit()
        except sqlite3.Error as e:
            conn.rollback()
            custom_popup(self, lang_lang.t("dialog_titles.error", "Error"), str(e), "error")
        finally:
            cur.close()
            conn.close()
        self.load_data()

    # --------------------------------------------------------------------------------
    # Utilities for name checking
    # --------------------------------------------------------------------------------
    def _fetch_existing_names(self):
        conn = connect_db()
        if conn is None:
            return []
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute("SELECT scenario_id, name FROM scenarios")
            rows = cur.fetchall()
            return [(r["scenario_id"], r["name"]) for r in rows]
        finally:
            cur.close()
            conn.close()

    def _is_duplicate_name(self, name: str, exclude_id=None) -> bool:
        name_l = (name or "").strip().lower()
        for sid, existing_name in self._fetch_existing_names():
            if exclude_id and sid == exclude_id:
                continue
            if existing_name.strip().lower() == name_l:
                return True
        return False

    def _find_similar_names(self, name: str, exclude_id=None):
        similar = []
        for sid, existing_name in self._fetch_existing_names():
            if exclude_id and sid == exclude_id:
                continue
            ratio = similarity_ratio(name, existing_name)
            if 1.0 > ratio >= SIMILARITY_THRESHOLD:
                similar.append(existing_name)
        return similar

    # --------------------------------------------------------------------------------
    # Form (Add/Edit)
    # --------------------------------------------------------------------------------
    def _open_form(self, edit=False, scenario_data=None):
        if not self._can_modify():
            self._denied_popup()
            return

        form = tk.Toplevel(self)
        form.title(
            lang_lang.t("scenarios.edit", "Edit Scenario") if edit
            else lang_lang.t("scenarios.add", "Add Scenario")
        )
        form.configure(bg=BG_MAIN)
        form.geometry("430x520")
        form.transient(self)
        form.grab_set()

        def label(parent, text):
            return tk.Label(parent, text=text, bg=BG_MAIN, fg=COLOR_PRIMARY,
                            font=("Helvetica", 10), anchor="w", justify="left")

        frame = tk.Frame(form, bg=BG_MAIN)
        frame.pack(fill="both", expand=True, padx=16, pady=16)

        # Name
        label(frame, lang_lang.t("scenarios.name", "Scenario Name")).pack(fill="x", pady=(0, 4))
        name_entry = ttk.Entry(frame, style="Sc.TEntry")
        name_entry.pack(fill="x", pady=(0, 8))

        # Activity Type
        label(frame, lang_lang.t("scenarios.activity_type", "Activity Type")).pack(fill="x", pady=(0, 4))
        activity_types = [
            "IPD", "OPD", "IPD+OPD", "Surgical", "Trauma",
            "Nutrition", "Epidemic", "Vaccination", "Displacement",
            "Field Hospital", "Maternity", "NCD", "MCP", "Bulk stock"
        ]
        type_cb = ttk.Combobox(frame, values=activity_types, state="readonly", style="Sc.TCombobox")
        type_cb.pack(fill="x", pady=(0, 8))

        # Target Population
        label(frame, lang_lang.t("scenarios.target_population", "Target Population")).pack(fill="x", pady=(0, 4))
        population_entry = ttk.Entry(frame, style="Sc.TEntry")
        population_entry.pack(fill="x", pady=(0, 8))

        # Stock Location
        label(frame, lang_lang.t("scenarios.stock_location", "Stock Location")).pack(fill="x", pady=(0, 4))
        location_entry = ttk.Entry(frame, style="Sc.TEntry")
        location_entry.pack(fill="x", pady=(0, 8))

        # Responsible
        label(frame, lang_lang.t("scenarios.responsible_person", "Person Responsible")).pack(fill="x", pady=(0, 4))
        responsible_entry = ttk.Entry(frame, style="Sc.TEntry")
        responsible_entry.pack(fill="x", pady=(0, 8))

        # Pre-fill if edit
        if edit and scenario_data:
            name_entry.insert(0, str(scenario_data[1]))
            type_cb.set(str(scenario_data[2]))
            if scenario_data[3] not in ("", None):
                population_entry.insert(0, str(scenario_data[3]))
            location_entry.insert(0, str(scenario_data[4]))
            responsible_entry.insert(0, str(scenario_data[5]))

        def validate_int_or_blank(value):
            v = value.strip()
            return (v == "") or v.isdigit()

        def save():
            name = name_entry.get().strip()
            activity = type_cb.get().strip()
            population_raw = population_entry.get().strip()
            location = location_entry.get().strip()
            responsible = responsible_entry.get().strip()

            if not (name and activity and location and responsible):
                custom_popup(form,
                             lang_lang.t("dialog_titles.error", "Error"),
                             lang_lang.t("scenarios.required",
                                         "Name, Activity Type, Stock Location, and Person Responsible are required"),
                             "error")
                return

            if not validate_int_or_blank(population_raw):
                custom_popup(form,
                             lang_lang.t("dialog_titles.error", "Error"),
                             lang_lang.t("scenarios.population_integer",
                                         "Target Population must be an integer (or blank)"),
                             "error")
                return

            exclude_id = scenario_data[0] if (edit and scenario_data) else None
            if self._is_duplicate_name(name, exclude_id=exclude_id):
                custom_popup(form,
                             lang_lang.t("dialog_titles.error", "Error"),
                             lang_lang.t("scenarios.duplicate", "Scenario name already exists"),
                             "error")
                return

            similar = self._find_similar_names(name, exclude_id=exclude_id)
            if similar:
                msg = lang_lang.t(
                    "scenarios.similarity_warning",
                    "Looks like there is a similar scenario already present ({names}). Are you sure to add this scenario?"
                ).format(names=", ".join(similar[:3]))
                ans = custom_askyesno(form,
                                      lang_lang.t("scenarios.similar_found", "Similar Scenario"),
                                      msg)
                if ans != "yes":
                    return

            population_val = int(population_raw) if population_raw else None

            conn = connect_db()
            if conn is None:
                custom_popup(form,
                             lang_lang.t("dialog_titles.error", "Error"),
                             lang_lang.t("scenarios.db_error", "Database connection failed"),
                             "error")
                return
            cur = conn.cursor()
            try:
                if edit and scenario_data:
                    cur.execute("""
                        UPDATE scenarios
                        SET name=?, activity_type=?, target_population=?, stock_location=?, responsible_person=?, created_at=CURRENT_TIMESTAMP
                        WHERE scenario_id=?
                    """, (name, activity, population_val, location, responsible, scenario_data[0]))
                else:
                    cur.execute("SELECT COUNT(*) FROM scenarios")
                    c = cur.fetchone()[0]
                    if c >= MAX_SCENARIOS:
                        custom_popup(form,
                                     lang_lang.t("dialog_titles.error", "Error"),
                                     lang_lang.t("scenarios.limit_reached",
                                                 "Maximum of 15 active scenarios reached. Adopt / edit an existing one."),
                                     "error")
                        cur.close()
                        conn.close()
                        return
                    cur.execute("""
                        INSERT INTO scenarios (name, activity_type, target_population, stock_location, responsible_person, created_at)
                        VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    """, (name, activity, population_val, location, responsible))
                conn.commit()
            except sqlite3.Error as e:
                conn.rollback()
                custom_popup(form,
                             lang_lang.t("dialog_titles.error", "Error"),
                             str(e),
                             "error")
                return
            finally:
                cur.close()
                conn.close()

            form.destroy()
            self.load_data()

        ttk.Button(
            frame,
            text=lang_lang.t("save", "Save"),
            style="Sc.TButton",
            command=save
        ).pack(fill="x", pady=18)

        name_entry.focus()

    # --------------------------------------------------------------------------------
    # Public refresh
    # --------------------------------------------------------------------------------
    def refresh(self):
        self.load_data()


# Standalone test runner
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Scenarios Management (Restricted Test)")
    class DummyApp:
        # Try "~" or "$" here to test restriction
        role = "~"
    app = DummyApp()
    Scenarios(root, app)
    root.geometry("1000x620")
    root.mainloop()