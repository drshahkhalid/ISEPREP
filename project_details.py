import tkinter as tk
from tkinter import ttk
import sqlite3
from db import connect_db
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno, custom_dialog
import logging

logger = logging.getLogger(__name__)

class ProjectDetailsWindow(tk.Toplevel):
    """
    Project details configuration window.

    New features added:
      - Buffer (months) spinbox (0–9) after Cover Period.
      - Role-based edit rules:
          * project_name, project_code, eprep_type editable only by roles: admin, hq, coordinator
          * update button enabled only for roles: admin, hq
      - Save button only shown when no project exists yet (initial setup).
      - Automatic prompt if any required field missing.
      - Graceful schema adaptation: adds buffer_months column if missing.
      - Personalised popups via custom_popup.
    """
    NAME_CODE_TYPE_EDIT_ROLES = {"admin", "hq", "coordinator"}
    UPDATE_BUTTON_ROLES = {"admin", "hq"}

    def __init__(self, parent, current_user: dict):
        super().__init__(parent)
        self.parent = parent
        self.current_user = current_user or {}
        self.role = (self.current_user.get("role") or "").lower()

        self.title(lang.t("project_details.window_title",
                          fallback="Project Details"))
        self.configure(bg="#F0F4F8")
        self.geometry("560x430")
        self.resizable(False, False)

        # Make sure schema has buffer_months
        self.ensure_schema()

        # Data
        self.project_data = self.load_project()

        # UI
        self.create_widgets()
        self.populate_fields()
        self.apply_role_permissions()
        self.update_parent_title()
        self.validate_and_prompt_missing()

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.focus_force()

    # ------------- Translation helper -------------
    def t(self, key, fallback=None, **kwargs):
        return lang.t(f"project_details.{key}", fallback=fallback, **kwargs)

    # ------------- Schema adaptation -------------
    def ensure_schema(self):
        """
        Add buffer_months column if missing.
        """
        conn = connect_db()
        if not conn:
            return
        cur = conn.cursor()
        try:
            cur.execute("PRAGMA table_info(project_details)")
            cols = {r[1].lower() for r in cur.fetchall()}
            if "buffer_months" not in cols:
                cur.execute("ALTER TABLE project_details ADD COLUMN buffer_months INTEGER DEFAULT 0")
                conn.commit()
        except sqlite3.Error as e:
            logger.error(f"[ensure_schema] {e}")
        finally:
            cur.close()
            conn.close()

    # ------------- UI Construction -------------
    def create_widgets(self):
        pad_lbl = {'padx': 10, 'pady': 4, 'sticky': 'w'}
        pad_field = {'padx': 10, 'pady': 4, 'sticky': 'w'}

        # Mapping internal enum -> localized label
        self.eprep_types_map = {
            "By Kits": self.t("types.by_kits", fallback="By Kits"),
            "By Items": self.t("types.by_items", fallback="By Items"),
            "Hybrid":   self.t("types.hybrid",   fallback="Hybrid")
        }
        self.freq_map = {
            "Once a year":  self.t("frequencies.once",   fallback="Once a year"),
            "Twice a year": self.t("frequencies.twice",  fallback="Twice a year"),
            "Three times a year": self.t("frequencies.thrice", fallback="Three times a year")
        }

        row = 0
        lbl_style = {'bg': "#F0F4F8", 'font': ("Helvetica", 10, "bold")}
        # Project Name
        tk.Label(self, text=self.t("project_name", fallback="Project Name"), **lbl_style).grid(row=row, column=0, **pad_lbl)
        self.entry_name = tk.Entry(self, width=42)
        self.entry_name.grid(row=row, column=1, columnspan=2, **pad_field)

        row += 1
        tk.Label(self, text=self.t("project_code", fallback="Project Code"), **lbl_style).grid(row=row, column=0, **pad_lbl)
        self.entry_code = tk.Entry(self, width=20)
        self.entry_code.grid(row=row, column=1, **pad_field)

        row += 1
        tk.Label(self, text=self.t("eprep_type", fallback="Management Type"), **lbl_style).grid(row=row, column=0, **pad_lbl)
        self.combo_type = ttk.Combobox(self, state="readonly", width=28,
                                       values=list(self.eprep_types_map.values()))
        self.combo_type.grid(row=row, column=1, columnspan=2, **pad_field)

        row += 1
        tk.Label(self, text=self.t("replenishment_frequency", fallback="Replenishment Frequency"),
                 **lbl_style).grid(row=row, column=0, **pad_lbl)
        self.combo_frequency = ttk.Combobox(self, state="readonly", width=28,
                                            values=list(self.freq_map.values()))
        self.combo_frequency.grid(row=row, column=1, columnspan=2, **pad_field)

        row += 1
        tk.Label(self, text=self.t("lead_time", fallback="Lead Time (Months)"), **lbl_style)\
            .grid(row=row, column=0, **pad_lbl)
        self.spin_lead = tk.Spinbox(self, from_=1, to=15, width=6, justify="right")
        self.spin_lead.grid(row=row, column=1, **pad_field)

        row += 1
        tk.Label(self, text=self.t("cover_period", fallback="Cover Period (Months)"), **lbl_style)\
            .grid(row=row, column=0, **pad_lbl)
        self.spin_cover = tk.Spinbox(self, from_=1, to=12, width=6, justify="right")
        self.spin_cover.grid(row=row, column=1, **pad_field)

        row += 1
        tk.Label(self, text=self.t("buffer_months", fallback="Buffer (Months)"), **lbl_style)\
            .grid(row=row, column=0, **pad_lbl)
        self.spin_buffer = tk.Spinbox(self, from_=0, to=9, width=6, justify="right")
        self.spin_buffer.grid(row=row, column=1, **pad_field)

        # Buttons Frame
        row += 1
        btn_frame = tk.Frame(self, bg="#F0F4F8")
        btn_frame.grid(row=row, column=0, columnspan=3, pady=18)

        self.btn_save = tk.Button(btn_frame,
                                  text=self.t("save_button", fallback="Save"),
                                  width=12,
                                  command=self.save_project,
                                  bg="#27AE60", fg="white", relief="raised")
        self.btn_update = tk.Button(btn_frame,
                                    text=self.t("update_button", fallback="Update"),
                                    width=12,
                                    command=self.update_project,
                                    bg="#2980B9", fg="white", relief="raised")

        if self.project_data:
            # Existing project: show Update only
            self.btn_update.pack(side="left", padx=8)
        else:
            # No project: show Save only
            self.btn_save.pack(side="left", padx=8)

        # Tooltip style hint label (optional)
        hint = tk.Label(self, text=self.t("hint_autoprompt",
                                          fallback="* All fields required. Missing values will prompt."),
                        bg="#F0F4F8", fg="#555555", font=("Helvetica", 8, "italic"))
        hint.grid(row=row+1, column=0, columnspan=3, sticky="w", padx=10)

    # ------------- Load data -------------
    def load_project(self):
        conn = connect_db()
        if not conn:
            return None
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute('SELECT * FROM project_details ORDER BY id DESC LIMIT 1')
            return cur.fetchone()
        except sqlite3.Error as e:
            logger.error(f"[load_project] {e}")
            return None
        finally:
            cur.close()
            conn.close()

    # ------------- Populate UI fields -------------
    def populate_fields(self):
        if not self.project_data:
            # Defaults
            self.combo_type.set(self.eprep_types_map["By Kits"])
            self.combo_frequency.set(self.freq_map["Once a year"])
            self.spin_lead.delete(0, tk.END); self.spin_lead.insert(0, "1")
            self.spin_cover.delete(0, tk.END); self.spin_cover.insert(0, "6")
            self.spin_buffer.delete(0, tk.END); self.spin_buffer.insert(0, "0")
            return

        pd = self.project_data
        # Safely get columns (may be missing)
        def get_col(r, name, default=None):
            try:
                return r[name]
            except Exception:
                return default

        self.entry_name.insert(0, get_col(pd, 'project_name', ""))
        self.entry_code.insert(0, get_col(pd, 'project_code', ""))

        internal_type = get_col(pd, 'eprep_type', "By Kits")
        self.combo_type.set(self.eprep_types_map.get(internal_type, self.eprep_types_map["By Kits"]))

        internal_freq = get_col(pd, 'replenishment_frequency', "Once a year")
        self.combo_frequency.set(self.freq_map.get(internal_freq, self.freq_map["Once a year"]))

        lead_val = str(get_col(pd, 'lead_time_months', 1) or 1)
        cover_val = str(get_col(pd, 'cover_period_months', 6) or 6)
        buffer_val = str(get_col(pd, 'buffer_months', 0) or 0)

        self.spin_lead.delete(0, tk.END); self.spin_lead.insert(0, lead_val)
        self.spin_cover.delete(0, tk.END); self.spin_cover.insert(0, cover_val)
        self.spin_buffer.delete(0, tk.END); self.spin_buffer.insert(0, buffer_val)

    # ------------- Role Permissions -------------
    def apply_role_permissions(self):
        role = self.role
        # Editable base fields
        if role not in self.NAME_CODE_TYPE_EDIT_ROLES:
            # Restrict name, code, type
            for w in (self.entry_name, self.entry_code, self.combo_type):
                w.config(state="disabled")

        # Update button
        if self.project_data:
            if role not in self.UPDATE_BUTTON_ROLES:
                self.btn_update.config(state="disabled")

        # Save button (initial) allowed for admin/hq/coordinator only
        if not self.project_data:
            if role not in self.NAME_CODE_TYPE_EDIT_ROLES:
                self.btn_save.config(state="disabled")

    # ------------- Validation -------------
    def validate_and_prompt_missing(self):
        """
        If any required parameter missing -> prompt user to fill.
        """
        missing = []
        name = self.entry_name.get().strip()
        code = self.entry_code.get().strip()
        eprep_label = self.combo_type.get().strip()
        freq_label = self.combo_frequency.get().strip()
        lead = self.spin_lead.get().strip()
        cover = self.spin_cover.get().strip()
        buffer_m = self.spin_buffer.get().strip()

        if not name: missing.append(self.t("project_name", fallback="Project Name"))
        if not code: missing.append(self.t("project_code", fallback="Project Code"))
        if not eprep_label: missing.append(self.t("eprep_type", fallback="Management Type"))
        if not freq_label: missing.append(self.t("replenishment_frequency", fallback="Replenishment Frequency"))
        if not lead: missing.append(self.t("lead_time", fallback="Lead Time"))
        if not cover: missing.append(self.t("cover_period", fallback="Cover Period"))
        if not buffer_m: missing.append(self.t("buffer_months", fallback="Buffer Months"))

        if missing:
            custom_popup(self,
                         lang.t("dialog_titles.warning", fallback="Warning"),
                         self.t("missing_fields",
                                fallback="Please complete: ") + ", ".join(missing),
                         "warning")
            # Focus first missing
            if not name: self.entry_name.focus_set()
            elif not code: self.entry_code.focus_set()
            elif not eprep_label: self.combo_type.focus_set()
            elif not freq_label: self.combo_frequency.focus_set()
            elif not lead: self.spin_lead.focus_set()
            elif not cover: self.spin_cover.focus_set()
            else: self.spin_buffer.focus_set()

    # ------------- Data extraction -------------
    def get_form_data(self):
        label_type = self.combo_type.get()
        eprep_type = next((k for k, v in self.eprep_types_map.items() if v == label_type), "By Kits")

        label_freq = self.combo_frequency.get()
        replenishment_frequency = next((k for k, v in self.freq_map.items() if v == label_freq), "Once a year")

        def safe_int(spin_val, default, mn=None, mx=None):
            try:
                val = int(spin_val)
            except Exception:
                val = default
            if mn is not None and val < mn: val = mn
            if mx is not None and val > mx: val = mx
            return val

        lead = safe_int(self.spin_lead.get(), 1, 1, 15)
        cover = safe_int(self.spin_cover.get(), 6, 1, 12)
        buffer_m = safe_int(self.spin_buffer.get(), 0, 0, 9)

        return {
            "project_name": self.entry_name.get().strip(),
            "project_code": self.entry_code.get().strip(),
            "eprep_type": eprep_type,
            "replenishment_frequency": replenishment_frequency,
            "lead_time_months": lead,
            "cover_period_months": cover,
            "buffer_months": buffer_m
        }

    # ------------- Save -------------
    def save_project(self):
        if self.project_data:
            custom_popup(self, lang.t("dialog_titles.info","Info"),
                         self.t("already_exists", fallback="Project already exists. Use Update."),
                         "info")
            return
        data = self.get_form_data()
        if not data["project_name"] or not data["project_code"]:
            custom_popup(self, lang.t("dialog_titles.error","Error"),
                         self.t("required_fields", fallback="Project Name and Code are required."),
                         "error")
            return
        conn = connect_db()
        if not conn:
            custom_popup(self, lang.t("dialog_titles.error","Error"),
                         self.t("db_error", fallback="Database connection failed."),"error")
            return
        cur = conn.cursor()
        try:
            cur.execute("""
                INSERT INTO project_details
                (project_name, project_code, eprep_type, replenishment_frequency,
                 lead_time_months, cover_period_months, buffer_months, created_by)
                VALUES (?,?,?,?,?,?,?,?)
            """, (
                data["project_name"], data["project_code"], data["eprep_type"],
                data["replenishment_frequency"], data["lead_time_months"],
                data["cover_period_months"], data["buffer_months"],
                self.current_user.get("username")
            ))
            conn.commit()
            custom_popup(self, lang.t("dialog_titles.success","Success"),
                         self.t("save_success", fallback="Project saved successfully."),
                         "success")
            self.project_data = self.load_project()
            self.update_parent_title()
            if hasattr(self.parent, "refresh_dashboard_after_project_save"):
                try: self.parent.refresh_dashboard_after_project_save()
                except Exception: pass
            self.destroy()
        except sqlite3.Error as e:
            custom_popup(self, lang.t("dialog_titles.error","Error"),
                         self.t("db_error", fallback="Database error: {}").format(str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # ------------- Update -------------
    def update_project(self):
        if not self.project_data:
            custom_popup(self, lang.t("dialog_titles.info","Info"),
                         self.t("no_project", fallback="No project to update. Use Save first."),
                         "info")
            return
        if self.role not in self.UPDATE_BUTTON_ROLES:
            custom_popup(self, lang.t("dialog_titles.restricted","Restricted"),
                         self.t("no_update_rights", fallback="You are not allowed to update."),
                         "warning")
            return
        data = self.get_form_data()
        if not data["project_name"] or not data["project_code"]:
            custom_popup(self, lang.t("dialog_titles.error","Error"),
                         self.t("required_fields", fallback="Project Name and Code are required."),
                         "error")
            return
        conn = connect_db()
        if not conn:
            custom_popup(self, lang.t("dialog_titles.error","Error"),
                         self.t("db_error", fallback="Database connection failed."),
                         "error")
            return
        cur = conn.cursor()
        try:
            cur.execute("""
                UPDATE project_details SET
                    project_name = ?,
                    project_code = ?,
                    eprep_type = ?,
                    replenishment_frequency = ?,
                    lead_time_months = ?,
                    cover_period_months = ?,
                    buffer_months = ?,
                    updated_by = ?
                WHERE id = ?
            """, (
                data["project_name"], data["project_code"], data["eprep_type"],
                data["replenishment_frequency"], data["lead_time_months"],
                data["cover_period_months"], data["buffer_months"],
                self.current_user.get("username"),
                self.project_data["id"]
            ))
            conn.commit()
            custom_popup(self, lang.t("dialog_titles.success","Success"),
                         self.t("update_success", fallback="Project updated successfully."),
                         "success")
            self.project_data = self.load_project()
            self.update_parent_title()
            if hasattr(self.parent, "refresh_dashboard_after_project_save"):
                try: self.parent.refresh_dashboard_after_project_save()
                except Exception: pass
            self.destroy()
        except sqlite3.Error as e:
            custom_popup(self, lang.t("dialog_titles.error","Error"),
                         self.t("db_error", fallback="Database error: {}").format(str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # ------------- Parent title -------------
    def update_parent_title(self):
        conn = connect_db()
        if not conn:
            return
        cur = conn.cursor()
        try:
            cur.execute('SELECT project_name, project_code FROM project_details ORDER BY id DESC LIMIT 1')
            proj = cur.fetchone()
            if proj:
                name = proj[0] if isinstance(proj, tuple) else proj['project_name']
                code = proj[1] if isinstance(proj, tuple) else proj['project_code']
                new_title = f"{name} - {code}"
                try:
                    self.parent.title(new_title)
                    self.parent.project_title = new_title
                except Exception:
                    pass
            else:
                base = lang.t("app.title", fallback="IsEPREP") + " - " + self.t("setup_required", fallback="Setup Required")
                try:
                    self.parent.title(base)
                    self.parent.project_title = base
                except Exception:
                    pass
        except sqlite3.Error:
            try:
                self.parent.title(lang.t("app.title", fallback="IsEPREP") + " - DB Error")
            except Exception:
                pass
        finally:
            cur.close()
            conn.close()

    # ------------- Close -------------
    def on_close(self):
        # Optional: block close if required fields empty
        data = self.get_form_data()
        if (not self.project_data) and (not data["project_name"] or not data["project_code"]):
            if tk.messagebox.askyesno(
                lang.t("dialog_titles.confirm","Confirm"),
                self.t("confirm_close_incomplete", fallback="Required fields missing. Close anyway?")
            ):
                self.destroy()
        else:
            self.destroy()


if __name__ == "__main__":
    # Simple manual test harness
    from db import connect_db
    root = tk.Tk()
    root.title("Test Host")
    dummy_user = {"username": "tester", "role": "admin"}
    ProjectDetailsWindow(root, dummy_user)
    root.mainloop()