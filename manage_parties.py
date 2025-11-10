import tkinter as tk
from tkinter import ttk
import sqlite3
from db import connect_db
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno, custom_dialog

# ============================================================
# THEME (aligned with manage_items / receive_kit / manage_users)
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
BTN_MISC       = "#7F8C8D"

# Canonical roles allowed to modify (you can adjust as needed)
# We explicitly EXCLUDE supervisor. We also handle the symbol "$".
ALLOWED_CANONICAL_ROLES = {"admin", "manager"}   # leave as your original policy
SUPERVISOR_SYMBOLS = {"$", "supervisor"}         # anything here is denied editing


def _center_toplevel(win: tk.Toplevel, parent: tk.Widget = None):
    """Center a toplevel window relative to parent or screen."""
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


class ManageParties(tk.Frame):
    """
    Themed management UI for Third Parties or End Users.
    - NOW symbol-aware: role symbol "$" (supervisor) cannot add/edit/delete.
    - Canonical 'supervisor' also blocked.
    - Original allowed editors: admin, manager.
    """
    def __init__(self, parent, app, party_type: str):
        super().__init__(parent, bg=BG_MAIN)
        self.app = app
        self.role = getattr(app, "role", "supervisor")  # may be canonical or symbol
        self.party_type = party_type  # "third" or "end"
        self.tree = None
        self.entries = {}
        self._define_context()
        self._configure_styles()
        self._build_ui()
        self.load_data()

    # --------------------------------------------------------
    # Permission logic (centralized)
    # --------------------------------------------------------
    def _can_modify(self) -> bool:
        """
        Return True if current role (canonical or symbol) is allowed to modify.
        Blocks supervisor by canonical name or symbol "$".
        """
        raw = (self.role or "").strip().lower()
        if raw in SUPERVISOR_SYMBOLS:
            return False
        # If role was passed as symbol for an allowed one (e.g. admin -> "@"),
        # you can map if you later adopt symbols for others; for now we just
        # match canonical names in ALLOWED_CANONICAL_ROLES.
        return raw in ALLOWED_CANONICAL_ROLES

    # --------------------------------------------------------
    # Context / config
    # --------------------------------------------------------
    def _define_context(self):
        if self.party_type == "third":
            self.table_name = "third_parties"
            self.heading = lang.t("manage.third_parties", fallback="Manage Third Parties")
            self.id_field = "third_party_id"
            self.type_options = [
                "MSF-Same Section", "MSF-Other Section", "Non-MSF", "MOH"
            ]
            self.default_type = "MSF-Same Section"
            self.form_key_prefix = "third"
        else:
            self.table_name = "end_users"
            self.heading = lang.t("manage.end_users", fallback="Manage End Users")
            self.id_field = "end_user_id"
            self.type_options = ["Individual", "Organization"]
            self.default_type = "Individual"
            self.form_key_prefix = "end"

        self.columns = [
            "ID", "Name", "Type", "City", "Address",
            "Contact Person", "Email", "Phone"
        ]

    # --------------------------------------------------------
    # Translation helper
    # --------------------------------------------------------
    def t(self, key, **kwargs):
        return lang.t(f"manage_parties.{key}", **kwargs)

    # --------------------------------------------------------
    # Style
    # --------------------------------------------------------
    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Parties.Treeview",
            background=BG_PANEL,
            fieldbackground=BG_PANEL,
            foreground=COLOR_PRIMARY,
            rowheight=26,
            font=("Helvetica", 10),
            bordercolor=COLOR_BORDER,
            relief="flat"
        )
        style.map("Parties.Treeview",
                  background=[("selected", COLOR_ACCENT)],
                  foreground=[("selected", "#FFFFFF")])

        style.configure(
            "Parties.Treeview.Heading",
            background="#E5E8EB",
            foreground=COLOR_PRIMARY,
            font=("Helvetica", 11, "bold"),
            relief="flat",
            bordercolor=COLOR_BORDER
        )

    # --------------------------------------------------------
    # UI
    # --------------------------------------------------------
    def _build_ui(self):
        # Title
        tk.Label(
            self,
            text=self.heading,
            font=("Helvetica", 20, "bold"),
            bg=BG_MAIN,
            fg=COLOR_PRIMARY,
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(12, 8))

        # Tree frame with border
        outer = tk.Frame(self, bg=COLOR_BORDER, bd=1, relief="solid")
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        self.tree = ttk.Treeview(
            outer,
            columns=self.columns,
            show="headings",
            height=14,
            style="Parties.Treeview"
        )

        width_map = {
            "ID": 60,
            "Name": 180,
            "Type": 140,
            "City": 120,
            "Address": 200,
            "Contact Person": 160,
            "Email": 180,
            "Phone": 120
        }
        anchor_map = {col: "w" for col in self.columns}

        for col in self.columns:
            self.tree.heading(col, text=self.t(f"column.{col.lower()}", fallback=col))
            self.tree.column(col, width=width_map.get(col, 120), anchor=anchor_map.get(col, "w"))

        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        # Buttons
        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill="x", padx=12, pady=(0, 6))

        can_modify = self._can_modify()

        def mk_btn(text, cmd, color, fallback):
            return tk.Button(
                btn_frame,
                text=self.t(text, fallback=fallback),
                command=cmd if can_modify else self._show_restricted,
                bg=color if can_modify else BTN_DISABLED,
                fg="#FFFFFF",
                activebackground=color if can_modify else BTN_DISABLED,
                relief="flat",
                padx=14,
                pady=6,
                font=("Helvetica", 10, "bold"),
                state="normal"
            )

        self.btn_add = mk_btn("add_button", self.add_party, BTN_ADD, "Add")
        self.btn_add.pack(side="left", padx=4)

        self.btn_edit = mk_btn("edit_button", self.edit_party, BTN_EDIT, "Edit")
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = mk_btn("delete_button", self.delete_party, BTN_DELETE, "Delete")
        self.btn_delete.pack(side="left", padx=4)

        if not can_modify:
            custom_popup(
                self,
                lang.t("dialog_titles.restricted", "Restricted"),
                self.t("access_denied", fallback="You don't have permission to manage parties."),
                "warning"
            )

        # Status bar
        self.status_var = tk.StringVar(value=self.t("ready", fallback="Ready"))
        tk.Label(
            self,
            textvariable=self.status_var,
            anchor="w",
            bg=BG_MAIN,
            fg=COLOR_PRIMARY,
            relief="sunken"
        ).pack(fill="x", padx=12, pady=(0, 10))

    def _show_restricted(self):
        custom_popup(
            self,
            lang.t("dialog_titles.restricted", "Restricted"),
            self.t("access_denied", fallback="You don't have permission to manage parties."),
            "warning"
        )

    # --------------------------------------------------------
    # Row tags
    # --------------------------------------------------------
    def _row_tags_config(self):
        self.tree.tag_configure("norm", background=ROW_NORM_COLOR)
        self.tree.tag_configure("alt", background=ROW_ALT_COLOR)

    # --------------------------------------------------------
    # Data load
    # --------------------------------------------------------
    def load_data(self):
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
            cur.execute(f'SELECT * FROM "{self.table_name}" ORDER BY {self.id_field}')
            rows = cur.fetchall()
            for idx, row in enumerate(rows):
                values = (
                    row[self.id_field],
                    (row["name"] or ""),
                    (row["type"] or ""),
                    (row["city"] or ""),
                    (row["address"] or ""),
                    (row["contact_person"] or ""),
                    (row["email"] or ""),
                    (row["phone"] or "")
                )
                tag = "alt" if idx % 2 else "norm"
                self.tree.insert("", "end", values=values, tags=(tag,))
            self._row_tags_config()
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

    # --------------------------------------------------------
    # Form
    # --------------------------------------------------------
    def open_form(self, title, save_callback, initial_data=None):
        form = tk.Toplevel(self)
        form.title(title)
        form.configure(bg=BG_MAIN)
        form.geometry("460x600")
        form.transient(self)
        form.grab_set()

        tk.Label(
            form,
            text=title,
            font=("Helvetica", 16, "bold"),
            fg=COLOR_PRIMARY,
            bg=BG_MAIN,
            anchor="w"
        ).pack(fill="x", padx=16, pady=(16, 8))

        fields = [
            ("name", True),
            ("type", True),
            ("city", False),
            ("address", False),
            ("contact_person", False),
            ("email", False),
            ("phone", False),
        ]

        self.entries = {}

        def add_field(fname, required):
            label_text = self.t(fname, fallback=fname.replace("_", " ").title())
            tk.Label(
                form,
                text=f"{label_text}{' *' if required else ''}:",
                bg=BG_MAIN,
                fg=COLOR_PRIMARY,
                font=("Helvetica", 10),
                anchor="w"
            ).pack(fill="x", padx=18, pady=(6, 0))

            if fname == "type":
                cb = ttk.Combobox(form, state="readonly", values=self.type_options)
                cb.set(self.default_type)
                cb.pack(fill="x", padx=18, pady=(0, 6))
                self.entries[fname] = cb
            else:
                ent = tk.Entry(form, font=("Helvetica", 11), relief="solid", bd=1)
                ent.pack(fill="x", padx=18, pady=(0, 6))
                self.entries[fname] = ent

        for fname, req in fields:
            add_field(fname, req)

        # Prefill
        if initial_data:
            for key, widget in self.entries.items():
                val = initial_data.get(key, "")
                if isinstance(widget, ttk.Combobox):
                    widget.set(val or self.default_type)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, val or "")

        def do_save():
            save_callback(form)

        tk.Button(
            form,
            text=self.t("save_button", fallback="Save"),
            command=do_save,
            bg=COLOR_ACCENT,
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            relief="flat",
            font=("Helvetica", 11, "bold"),
            padx=14, pady=8
        ).pack(fill="x", padx=18, pady=18)

        form.after(50, lambda: _center_toplevel(form, self))

    # --------------------------------------------------------
    # Add
    # --------------------------------------------------------
    def add_party(self):
        if not self._can_modify():
            self._show_restricted()
            return

        def save(form):
            data = {k: (w.get().strip() if hasattr(w, "get") else "") for k, w in self.entries.items()}
            if not data["name"] or not data["type"]:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("required_fields", fallback="Name and Type are required."),
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
                cur.execute(f"""
                    INSERT INTO "{self.table_name}"
                      (name, type, city, address, contact_person, email, phone)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    data["name"], data["type"], data.get("city"),
                    data.get("address"), data.get("contact_person"),
                    data.get("email"), data.get("phone")
                ))
                conn.commit()
                custom_popup(self,
                             lang.t("dialog_titles.success", "Success"),
                             self.t("add_success", fallback="Record added successfully."),
                             "info")
                form.destroy()
                self.load_data()
            except sqlite3.Error as e:
                conn.rollback()
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                             "error")
            finally:
                cur.close()
                conn.close()

        title = self.t("add_title",
                       fallback=f"Add {'Third Party' if self.party_type=='third' else 'End User'}")
        self.open_form(title, save)

    # --------------------------------------------------------
    # Edit
    # --------------------------------------------------------
    def edit_party(self):
        if not self._can_modify():
            self._show_restricted()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("select_record", fallback="Please select a record to edit"),
                         "error")
            return
        values = self.tree.item(sel[0])["values"]
        party_id = values[0]
        initial_data = {
            "name": values[1],
            "type": values[2],
            "city": values[3],
            "address": values[4],
            "contact_person": values[5],
            "email": values[6],
            "phone": values[7]
        }

        def save(form):
            data = {k: (w.get().strip() if hasattr(w, "get") else "") for k, w in self.entries.items()}
            if not data["name"] or not data["type"]:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("required_fields", fallback="Name and Type are required."),
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
                cur.execute(f"""
                    UPDATE "{self.table_name}"
                       SET name=?, type=?, city=?, address=?, contact_person=?, email=?, phone=?
                     WHERE {self.id_field} = ?
                """, (
                    data["name"], data["type"], data.get("city"),
                    data.get("address"), data.get("contact_person"),
                    data.get("email"), data.get("phone"), party_id
                ))
                conn.commit()
                custom_popup(self,
                             lang.t("dialog_titles.success", "Success"),
                             self.t("edit_success", fallback="Record updated successfully."),
                             "info")
                form.destroy()
                self.load_data()
            except sqlite3.Error as e:
                conn.rollback()
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                             "error")
            finally:
                cur.close()
                conn.close()

        title = self.t("edit_title",
                       fallback=f"Edit {'Third Party' if self.party_type=='third' else 'End User'}")
        self.open_form(title, save, initial_data)

    # --------------------------------------------------------
    # Delete
    # --------------------------------------------------------
    def delete_party(self):
        if not self._can_modify():
            self._show_restricted()
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("select_record", fallback="Please select a record to delete"),
                         "error")
            return
        party_id = self.tree.item(sel[0])["values"][0]
        ans = custom_askyesno(
            self,
            lang.t("dialog_titles.confirm", "Confirm"),
            self.t("confirm_delete", fallback="Delete this record?")
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
            cur.execute(f'DELETE FROM "{self.table_name}" WHERE {self.id_field} = ?', (party_id,))
            conn.commit()
            custom_popup(self,
                         lang.t("dialog_titles.success", "Success"),
                         self.t("delete_success", fallback="Record deleted successfully."),
                         "info")
            self.load_data()
        except sqlite3.Error as e:
            conn.rollback()
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # --------------------------------------------------------
    # External refresh
    # --------------------------------------------------------
    def refresh(self):
        self.load_data()


# Standalone test
if __name__ == "__main__":
    root = tk.Tk()
    class DummyApp:
        # Try role values: "admin", "manager", "$", "supervisor"
        role = "$"   # Should be read-only
        project_title = "IsEPREP"
    app = DummyApp()
    root.title("Manage Parties - Test")
    ManageParties(root, app, party_type="third")
    root.geometry("1100x640")
    root.mainloop()
