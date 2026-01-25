import tkinter as tk
from tkinter import ttk
import sqlite3
import logging
from db import connect_db
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno

# ============================================================
# IMPORT CENTRALIZED THEME (NEW)
# ============================================================
from theme_config import AppTheme, configure_tree_tags

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ============================================================
# REMOVED OLD COLOR CONSTANTS - Now using AppTheme
# ============================================================
# OLD (REMOVED):
# BG_MAIN        = "#F0F4F8"
# BG_PANEL       = "#FFFFFF"
# COLOR_PRIMARY  = "#2C3E50"
# COLOR_ACCENT   = "#2563EB"
# COLOR_BORDER   = "#D0D7DE"
# ROW_ALT_COLOR  = "#F7FAFC"
# ROW_NORM_COLOR = "#FFFFFF"
# BTN_ADD        = "#27AE60"
# BTN_EDIT       = "#2980B9"
# BTN_DELETE     = "#C0392B"
# BTN_DISABLED   = "#94A3B8"

# Canonical roles allowed to modify. Supervisor must NOT be able to edit.
ALLOWED_CANONICAL_ROLES = {"admin", "manager"}
BLOCKED_ROLES_OR_SYMBOLS = {"supervisor", "$"}

# Canonical English user types (stored in database)
USER_TYPE_OPTIONS_CANONICAL = [
    "Emergency Coordination",
    "Regular Coordination",
    "Emergency Project",
    "Regular Project",
    "Prepositioned Stock",
    "Staff Health"
]


def _center_toplevel(win: tk.Toplevel, parent: tk.Widget = None):
    win.update_idletasks()
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
    if x < 0: x = 0
    if y < 0: y = 0
    win.geometry(f"+{x}+{y}")


class ManageEndUsers(tk.Frame):
    """
    End Users management with: 
      - Role-symbol restriction (admin/manager can edit, supervisor/$ cannot)
      - Translatable user type dropdown (English in DB, translated in UI)
    """
    def __init__(self, parent, app):
        super().__init__(parent, bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        self.app = app
        self.role = getattr(app, "role", "supervisor")
        self.tree = None
        self.status_var = tk.StringVar(value=self.t("ready", fallback="Ready"))
        self.id_column_available = False
        self._configure_styles()
        self._build_ui()
        self.detect_id_column()
        self.load_end_users()

    # ---------------- Permission helpers ----------------
    def _can_modify(self) -> bool:
        raw = (self.role or "").strip().lower()
        if raw in BLOCKED_ROLES_OR_SYMBOLS:
            return False
        return raw in ALLOWED_CANONICAL_ROLES

    def _show_restricted(self):
        custom_popup(
            self,
            lang.t("dialog_titles.restricted", "Restricted"),
            self.t("access_denied", fallback="You don't have permission to manage end users."),
            "warning"
        )

    # ---------------- Translation helpers ----------------
    def t(self, key, **kwargs):
        return lang.t(f"end_users.{key}", **kwargs)

    def _get_user_type_map(self):
        """Return canonical_english -> translated_display mapping."""
        section = lang.get_section("end_users.user_types")
        if not section or not isinstance(section, dict):
            # Fallback: return identity mapping
            return {utype: utype for utype in USER_TYPE_OPTIONS_CANONICAL}
        mapping = {}
        for canonical in USER_TYPE_OPTIONS_CANONICAL:
            mapping[canonical] = section.get(canonical, canonical)
        return mapping

    def _get_user_type_reverse_map(self):
        """Return translated_display -> canonical_english mapping."""
        forward = self._get_user_type_map()
        reverse = {}
        for canonical, display in forward.items():
            # Normalize for case-insensitive matching
            reverse[display.strip().lower()] = canonical
        return reverse

    def _user_type_to_display(self, canonical: str) -> str:
        """Convert canonical English user type to translated display."""
        return self._get_user_type_map().get(canonical, canonical)

    def _user_type_to_canonical(self, display: str) -> str:
        """Convert translated display back to canonical English."""
        rev = self._get_user_type_reverse_map()
        return rev.get(display.strip().lower(), display)

    def _get_translated_user_types(self):
        """Return list of translated user type options for dropdown."""
        return [self._user_type_to_display(ut) for ut in USER_TYPE_OPTIONS_CANONICAL]

    # ---------------- Styles (UPDATED: Removed theme_use, using AppTheme) ----------------
    def _configure_styles(self):
        # Global theme already applied by login_gui
        style = ttk.Style()
        # REMOVED: style.theme_use("clam") - already applied globally
        
        style.configure(
            "EndUsers.Treeview",
            background=AppTheme.BG_PANEL,  # UPDATED: Use AppTheme
            fieldbackground=AppTheme.BG_PANEL,  # UPDATED: Use AppTheme
            foreground=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            rowheight=AppTheme.TREE_ROW_HEIGHT,  # UPDATED: Use AppTheme
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),  # UPDATED: Use AppTheme
            bordercolor=AppTheme.COLOR_BORDER,  # UPDATED: Use AppTheme
            relief="flat"
        )
        style.map("EndUsers.Treeview",
                  background=[("selected", AppTheme.COLOR_ACCENT)],  # UPDATED: Use AppTheme
                  foreground=[("selected", AppTheme.TEXT_WHITE)])  # UPDATED: Use AppTheme
        style.configure(
            "EndUsers.Treeview.Heading",
            background="#E5E8EB",
            foreground=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HEADING, "bold"),  # UPDATED: Use AppTheme
            relief="flat",
            bordercolor=AppTheme.COLOR_BORDER  # UPDATED: Use AppTheme
        )

    # ---------------- UI (UPDATED: All color references use AppTheme) ----------------
    def _build_ui(self):
        tk.Label(
            self,
            text=self.t("title", fallback="Manage End Users"),
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HUGE, "bold"),  # UPDATED: Use AppTheme
            bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
            fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(12, 8))

        # Tree container
        outer = tk.Frame(self, bg=AppTheme.COLOR_BORDER, bd=1, relief="solid")  # UPDATED: Use AppTheme
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        display_columns = ["Name", "User Type"]
        self.tree = ttk.Treeview(
            outer,
            columns=display_columns,
            show="headings",
            height=16,
            style="EndUsers.Treeview"
        )
        self.tree.heading("Name", text=self.t("column.name", fallback="Name"))
        self.tree.heading("User Type", text=self.t("column.user_type", fallback="User Type"))
        self.tree.column("Name", width=260, anchor="w")
        self.tree.column("User Type", width=200, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        # Buttons (UPDATED: All button colors use AppTheme)
        btn_frame = tk.Frame(self, bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        btn_frame.pack(fill="x", padx=12, pady=(0, 6))
        can_modify = self._can_modify()

        def mk_btn(label_key, fallback, cmd, color):
            return tk.Button(
                btn_frame,
                text=self.t(label_key, fallback=fallback),
                command=cmd if can_modify else self._show_restricted,
                bg=color if can_modify else AppTheme.BTN_DISABLED,  # UPDATED: Use AppTheme
                fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                activebackground=color if can_modify else AppTheme.BTN_DISABLED,  # UPDATED: Use AppTheme
                relief="flat",
                padx=14, pady=6,
                font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),  # UPDATED: Use AppTheme
                state="normal"
            )

        self.btn_add = mk_btn("add_button", "Add End User", self.add_end_user, AppTheme.BTN_SUCCESS)  # UPDATED: Use AppTheme
        self.btn_add.pack(side="left", padx=4)

        self.btn_edit = mk_btn("edit_button", "Edit End User", self.edit_end_user, AppTheme.BTN_WARNING)  # UPDATED: Use AppTheme
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = mk_btn("delete_button", "Delete End User", self.delete_end_user, AppTheme.BTN_DANGER)  # UPDATED: Use AppTheme
        self.btn_delete.pack(side="left", padx=4)

        if not can_modify:
            custom_popup(
                self,
                lang.t("dialog_titles.restricted", "Restricted"),
                self.t("access_denied", fallback="You don't have permission to manage end users."),
                "warning"
            )

        # Status bar
        tk.Label(
            self,
            textvariable=self.status_var,
            anchor="w",
            bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
            fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            relief="sunken"
        ).pack(fill="x", padx=12, pady=(0, 10))

    # ---------------- Detect ID column presence ----------------
    def detect_id_column(self):
        conn = connect_db()
        if conn is None:
            return
        cur = conn.cursor()
        try:
            cur.execute("PRAGMA table_info(end_users)")
            cols = [r[1].lower() for r in cur.fetchall()]
            self.id_column_available = "end_user_id" in cols
        except sqlite3.Error:
            self.id_column_available = False
        finally:
            cur.close()
            conn.close()

    # ---------------- Data Loading (UPDATED: Row colors use AppTheme) ----------------
    def load_end_users(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        conn = connect_db()
        if conn is None:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", fallback="Database connection failed"),
                         "error")
            return
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            if self.id_column_available:
                cur.execute('SELECT end_user_id, name, user_type FROM "end_users" ORDER BY name')
            else:
                cur.execute('SELECT name, user_type FROM "end_users" ORDER BY name')
            rows = cur.fetchall()
            for idx, row in enumerate(rows):
                name = row["name"] or ""
                utype_canonical = row["user_type"] or ""
                # Translate for display
                utype_display = self._user_type_to_display(utype_canonical)
                
                if self.id_column_available:
                    self.tree.insert(
                        "",
                        "end",
                        values=(name, utype_display),
                        tags=("alt" if idx % 2 else "norm",),
                        iid=f"end_{row['end_user_id']}"
                    )
                else: 
                    self.tree.insert(
                        "",
                        "end",
                        values=(name, utype_display),
                        tags=("alt" if idx % 2 else "norm",)
                    )
            self._tag_rows()
            self.status_var.set(self.t("loaded_records",
                                       fallback="Loaded {n} records").format(n=len(rows)))
        except sqlite3.Error as e:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    def _tag_rows(self):
        self.tree.tag_configure("norm", background=AppTheme.ROW_NORM)  # UPDATED: Use AppTheme
        self.tree.tag_configure("alt", background=AppTheme.ROW_ALT)  # UPDATED: Use AppTheme

    # ---------------- Add (UPDATED: All color references use AppTheme) ----------------
    def add_end_user(self):
        if not self._can_modify():
            self._show_restricted()
            return

        form = tk.Toplevel(self)
        form.title(self.t("add_title", fallback="Add End User"))
        form.configure(bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        form.geometry("420x380")
        form.transient(self)
        form.grab_set()

        def lbl(text):
            return tk.Label(form, text=text, 
                          bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                          fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                          font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),  # UPDATED: Use AppTheme
                          anchor="w")

        lbl(self.t("name", fallback="Name") + " *").pack(fill="x", padx=18, pady=(18, 2))
        name_entry = tk.Entry(form, 
                             font=(AppTheme.FONT_FAMILY, 11),  # UPDATED: Use AppTheme
                             relief="solid", bd=1)
        name_entry.pack(fill="x", padx=18, pady=(0, 8))

        lbl(self.t("user_type", fallback="User Type") + " *").pack(fill="x", padx=18, pady=(0, 2))
        
        # Get translated options for dropdown
        translated_options = self._get_translated_user_types()
        ut_var = tk.StringVar()
        ut_cb = ttk.Combobox(form, textvariable=ut_var, values=translated_options, state="readonly")
        ut_cb.set(translated_options[0])  # Default to first option
        ut_cb.pack(fill="x", padx=18, pady=(0, 12))

        def save():
            name = name_entry.get().strip()
            utype_display = ut_var.get().strip()
            # Convert display back to canonical English for storage
            utype_canonical = self._user_type_to_canonical(utype_display)
            
            if not name or not utype_canonical:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("required_fields", fallback="Name and User Type are required."),
                             "error")
                return
            conn = connect_db()
            if conn is None:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", fallback="Database connection failed"),
                             "error")
                return
            cur = conn.cursor()
            try:
                # Store canonical English in database
                cur.execute('INSERT INTO "end_users" (name, user_type) VALUES (?, ?)', 
                           (name, utype_canonical))
                conn.commit()
                custom_popup(self,
                             lang.t("dialog_titles.success", "Success"),
                             self.t("add_success", fallback="End user added successfully."),
                             "info")
                form.destroy()
                self.load_end_users()
            except sqlite3.Error as e:
                conn.rollback()
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                             "error")
            finally: 
                cur.close()
                conn.close()

        tk.Button(form,
                  text=self.t("save_button", fallback="Save"),
                  command=save,
                  bg=AppTheme.BTN_SUCCESS,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  font=(AppTheme.FONT_FAMILY, 11, "bold"),  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14, pady=8,
                  activebackground="#1E874B").pack(fill="x", padx=18, pady=(8, 18))

        form.after(50, lambda: _center_toplevel(form, self))
        name_entry.focus()

    # ---------------- Edit (UPDATED: All color references use AppTheme) ----------------
    def edit_end_user(self):
        if not self._can_modify():
            self._show_restricted()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("select_record", fallback="Select an end user to edit"),
                         "error")
            return

        iid = sel[0]
        values = self.tree.item(iid)["values"]
        old_name = values[0]
        old_type_display = values[1]  # This is translated display text

        end_user_id = None
        if self.id_column_available and iid.startswith("end_"):
            try:
                end_user_id = int(iid.split("_", 1)[1])
            except Exception:
                end_user_id = None

        form = tk.Toplevel(self)
        form.title(self.t("edit_title", fallback="Edit End User"))
        form.configure(bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        form.geometry("420x380")
        form.transient(self)
        form.grab_set()

        def lbl(text):
            return tk.Label(form, text=text, 
                          bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                          fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                          font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),  # UPDATED: Use AppTheme
                          anchor="w")

        lbl(self.t("name", fallback="Name") + " *").pack(fill="x", padx=18, pady=(18, 2))
        name_entry = tk.Entry(form, 
                             font=(AppTheme.FONT_FAMILY, 11),  # UPDATED: Use AppTheme
                             relief="solid", bd=1)
        name_entry.insert(0, old_name)
        name_entry.pack(fill="x", padx=18, pady=(0, 8))

        lbl(self.t("user_type", fallback="User Type") + " *").pack(fill="x", padx=18, pady=(0, 2))
        
        translated_options = self._get_translated_user_types()
        ut_var = tk.StringVar()
        ut_cb = ttk.Combobox(form, textvariable=ut_var, values=translated_options, state="readonly")
        
        # Set the current value (already translated)
        if old_type_display in translated_options:
            ut_cb.set(old_type_display)
        else:
            ut_cb.set(translated_options[0])
        ut_cb.pack(fill="x", padx=18, pady=(0, 12))

        def save():
            new_name = name_entry.get().strip()
            utype_display = ut_var.get().strip()
            # Convert back to canonical English
            utype_canonical = self._user_type_to_canonical(utype_display)
            
            if not new_name or not utype_canonical:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("required_fields", fallback="Name and User Type are required."),
                             "error")
                return
            conn = connect_db()
            if conn is None:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", fallback="Database connection failed"),
                             "error")
                return
            cur = conn.cursor()
            try:
                # Store canonical English in database
                if self.id_column_available and end_user_id is not None: 
                    cur.execute('UPDATE "end_users" SET name=?, user_type=? WHERE end_user_id=?',
                                (new_name, utype_canonical, end_user_id))
                else:
                    cur.execute('UPDATE "end_users" SET name=?, user_type=? WHERE name=?',
                                (new_name, utype_canonical, old_name))
                conn.commit()
                custom_popup(self,
                             lang.t("dialog_titles.success", "Success"),
                             self.t("edit_success", fallback="End user updated successfully."),
                             "info")
                form.destroy()
                self.load_end_users()
            except sqlite3.Error as e:
                conn.rollback()
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                             "error")
            finally:
                cur.close()
                conn.close()

        tk.Button(form,
                  text=self.t("save_button", fallback="Save"),
                  command=save,
                  bg=AppTheme.BTN_WARNING,  # UPDATED: Use AppTheme
                  fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
                  font=(AppTheme.FONT_FAMILY, 11, "bold"),  # UPDATED: Use AppTheme
                  relief="flat",
                  padx=14, pady=8,
                  activebackground="#1F5D82").pack(fill="x", padx=18, pady=(8, 18))

        form.after(50, lambda: _center_toplevel(form, self))
        name_entry.focus()

    # ---------------- Delete ----------------
    def delete_end_user(self):
        if not self._can_modify():
            self._show_restricted()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("select_record", fallback="Select an end user to delete"),
                         "error")
            return
        iid = sel[0]
        values = self.tree.item(iid)["values"]
        name = values[0]
        end_user_id = None
        if self.id_column_available and iid.startswith("end_"):
            try:
                end_user_id = int(iid.split("_", 1)[1])
            except Exception:
                end_user_id = None

        ans = custom_askyesno(
            self,
            lang.t("dialog_titles.confirm", "Confirm"),
            self.t("confirm_delete", fallback="Delete end user '{name}'?").format(name=name)
        )
        if ans != "yes":
            return

        conn = connect_db()
        if conn is None:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", fallback="Database connection failed"),
                         "error")
            return
        cur = conn.cursor()
        try:
            if self.id_column_available and end_user_id is not None:
                cur.execute('DELETE FROM "end_users" WHERE end_user_id = ?', (end_user_id,))
            else:
                cur.execute('DELETE FROM "end_users" WHERE name = ?', (name,))
            conn.commit()
            custom_popup(self,
                         lang.t("dialog_titles.success", "Success"),
                         self.t("delete_success", fallback="End user deleted successfully."),
                         "info")
            self.load_end_users()
        except sqlite3.Error as e:
            conn.rollback()
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # ---------------- External refresh ----------------
    def refresh(self):
        self.load_end_users()


# Standalone test
if __name__ == "__main__":
    root = tk.Tk()
    class DummyApp:
        role = "admin"  # Try: "admin", "manager", "$", "supervisor"
        project_title = "IsEPREP"
    app = DummyApp()
    root.title("Manage End Users - Test")
    ManageEndUsers(root, app)
    root.geometry("900x600")
    root.mainloop()