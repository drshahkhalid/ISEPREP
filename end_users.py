import tkinter as tk
from tkinter import ttk
import sqlite3
import logging
from db import connect_db
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ============================================================
# THEME
# ============================================================
BG_MAIN        = "#F0F4F8"
BG_PANEL       = "#FFFFFF"
COLOR_PRIMARY  = "#2C3E50"
COLOR_ACCENT   = "#2563EB"
COLOR_BORDER   = "#D0D7DE"
ROW_ALT_COLOR  = "#F7FAFC"
ROW_NORM_COLOR = "#FFFFFF"
BTN_ADD        = "#27AE60"
BTN_EDIT       = "#2980B9"
BTN_DELETE     = "#C0392B"
BTN_DISABLED   = "#94A3B8"

# Canonical roles allowed to modify. Supervisor must NOT be able to edit.
# We also block the supervisor symbol "$".
ALLOWED_CANONICAL_ROLES = {"admin", "manager"}
BLOCKED_ROLES_OR_SYMBOLS = {"supervisor", "$"}

USER_TYPE_OPTIONS = [
    "E-Coordination",
    "Regular Coordination",
    "E-Project",
    "Regular Project",
    "Prepositioned stock",
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
    End Users management with role-symbol restriction:
      - admin / manager can add/edit/delete
      - supervisor (canonical) or symbol "$" are read-only
      - Others (hq, coordinator, etc.) are currently treated as read-only (adjust logic if needed)
    """
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG_MAIN)
        self.app = app
        self.role = getattr(app, "role", "supervisor")  # may be canonical or symbol
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
        # If symbols used for admin/manager in future you can map here; for now we just match canonical names.
        return raw in ALLOWED_CANONICAL_ROLES

    def _show_restricted(self):
        custom_popup(
            self,
            lang.t("dialog_titles.restricted", "Restricted"),
            self.t("access_denied", fallback="You don't have permission to manage end users."),
            "warning"
        )

    # ---------------- Translation shortcut ----------------
    def t(self, key, **kwargs):
        return lang.t(f"end_users.{key}", **kwargs)

    # ---------------- Styles ----------------
    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "EndUsers.Treeview",
            background=BG_PANEL,
            fieldbackground=BG_PANEL,
            foreground=COLOR_PRIMARY,
            rowheight=26,
            font=("Helvetica", 10),
            bordercolor=COLOR_BORDER,
            relief="flat"
        )
        style.map("EndUsers.Treeview",
                  background=[("selected", COLOR_ACCENT)],
                  foreground=[("selected", "#FFFFFF")])
        style.configure(
            "EndUsers.Treeview.Heading",
            background="#E5E8EB",
            foreground=COLOR_PRIMARY,
            font=("Helvetica", 11, "bold"),
            relief="flat",
            bordercolor=COLOR_BORDER
        )

    # ---------------- UI ----------------
    def _build_ui(self):
        tk.Label(
            self,
            text=self.t("title", fallback="Manage End Users"),
            font=("Helvetica", 20, "bold"),
            bg=BG_MAIN,
            fg=COLOR_PRIMARY,
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(12, 8))

        # Tree container
        outer = tk.Frame(self, bg=COLOR_BORDER, bd=1, relief="solid")
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

        # Buttons
        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill="x", padx=12, pady=(0, 6))
        can_modify = self._can_modify()

        def mk_btn(label_key, fallback, cmd, color):
            return tk.Button(
                btn_frame,
                text=self.t(label_key, fallback=fallback),
                command=cmd if can_modify else self._show_restricted,
                bg=color if can_modify else BTN_DISABLED,
                fg="#FFFFFF",
                activebackground=color if can_modify else BTN_DISABLED,
                relief="flat",
                padx=14, pady=6,
                font=("Helvetica", 10, "bold"),
                state="normal"
            )

        self.btn_add = mk_btn("add_button", "Add End User", self.add_end_user, BTN_ADD)
        self.btn_add.pack(side="left", padx=4)

        self.btn_edit = mk_btn("edit_button", "Edit End User", self.edit_end_user, BTN_EDIT)
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = mk_btn("delete_button", "Delete End User", self.delete_end_user, BTN_DELETE)
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
            bg=BG_MAIN,
            fg=COLOR_PRIMARY,
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

    # ---------------- Data Loading ----------------
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
                if self.id_column_available:
                    name = row["name"] or ""
                    utype = row["user_type"] or ""
                    self.tree.insert(
                        "",
                        "end",
                        values=(name, utype),
                        tags=("alt" if idx % 2 else "norm",),
                        iid=f"end_{row['end_user_id']}"
                    )
                else:
                    self.tree.insert(
                        "",
                        "end",
                        values=(row["name"] or "", row["user_type"] or ""),
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
        self.tree.tag_configure("norm", background=ROW_NORM_COLOR)
        self.tree.tag_configure("alt", background=ROW_ALT_COLOR)

    # ---------------- Add ----------------
    def add_end_user(self):
        if not self._can_modify():
            self._show_restricted()
            return

        form = tk.Toplevel(self)
        form.title(self.t("add_title", fallback="Add End User"))
        form.configure(bg=BG_MAIN)
        form.geometry("420x380")
        form.transient(self)
        form.grab_set()

        def lbl(text):
            return tk.Label(form, text=text, bg=BG_MAIN, fg=COLOR_PRIMARY,
                            font=("Helvetica", 10), anchor="w")

        lbl(self.t("name", fallback="Name") + " *").pack(fill="x", padx=18, pady=(18, 2))
        name_entry = tk.Entry(form, font=("Helvetica", 11), relief="solid", bd=1)
        name_entry.pack(fill="x", padx=18, pady=(0, 8))

        lbl(self.t("user_type", fallback="User Type") + " *").pack(fill="x", padx=18, pady=(0, 2))
        ut_var = tk.StringVar()
        ut_cb = ttk.Combobox(form, textvariable=ut_var, values=USER_TYPE_OPTIONS, state="readonly")
        ut_cb.set(USER_TYPE_OPTIONS[0])
        ut_cb.pack(fill="x", padx=18, pady=(0, 12))

        def save():
            name = name_entry.get().strip()
            utype = ut_var.get().strip()
            if not name or not utype:
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
                cur.execute('INSERT INTO "end_users" (name, user_type) VALUES (?, ?)', (name, utype))
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
                  bg=BTN_ADD,
                  fg="#FFFFFF",
                  font=("Helvetica", 11, "bold"),
                  relief="flat",
                  padx=14, pady=8,
                  activebackground="#1E874B").pack(fill="x", padx=18, pady=(8, 18))

        form.after(50, lambda: _center_toplevel(form, self))
        name_entry.focus()

    # ---------------- Edit ----------------
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
        old_type = values[1]

        end_user_id = None
        if self.id_column_available and iid.startswith("end_"):
            try:
                end_user_id = int(iid.split("_", 1)[1])
            except Exception:
                end_user_id = None

        form = tk.Toplevel(self)
        form.title(self.t("edit_title", fallback="Edit End User"))
        form.configure(bg=BG_MAIN)
        form.geometry("420x380")
        form.transient(self)
        form.grab_set()

        def lbl(text):
            return tk.Label(form, text=text, bg=BG_MAIN, fg=COLOR_PRIMARY,
                            font=("Helvetica", 10), anchor="w")

        lbl(self.t("name", fallback="Name") + " *").pack(fill="x", padx=18, pady=(18, 2))
        name_entry = tk.Entry(form, font=("Helvetica", 11), relief="solid", bd=1)
        name_entry.insert(0, old_name)
        name_entry.pack(fill="x", padx=18, pady=(0, 8))

        lbl(self.t("user_type", fallback="User Type") + " *").pack(fill="x", padx=18, pady=(0, 2))
        ut_var = tk.StringVar()
        ut_cb = ttk.Combobox(form, textvariable=ut_var, values=USER_TYPE_OPTIONS, state="readonly")
        if old_type in USER_TYPE_OPTIONS:
            ut_cb.set(old_type)
        else:
            ut_cb.set(USER_TYPE_OPTIONS[0])
        ut_cb.pack(fill="x", padx=18, pady=(0, 12))

        def save():
            new_name = name_entry.get().strip()
            new_type = ut_var.get().strip()
            if not new_name or not new_type:
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
                if self.id_column_available and end_user_id is not None:
                    cur.execute('UPDATE "end_users" SET name=?, user_type=? WHERE end_user_id=?',
                                (new_name, new_type, end_user_id))
                else:
                    cur.execute('UPDATE "end_users" SET name=?, user_type=? WHERE name=?',
                                (new_name, new_type, old_name))
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
                  bg=BTN_EDIT,
                  fg="#FFFFFF",
                  font=("Helvetica", 11, "bold"),
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
        # Try "admin", "manager", "$", "supervisor"
        role = "$"   # Should be read-only
        project_title = "IsEPREP"
    app = DummyApp()
    root.title("Manage End Users - Test")
    ManageEndUsers(root, app)
    root.geometry("900x600")
    root.mainloop()
