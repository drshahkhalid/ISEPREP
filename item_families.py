import tkinter as tk
from tkinter import ttk
import sqlite3
from db import connect_db
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno, custom_dialog

# ============================================================
# THEME (consistent with other management screens)
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

ALLOWED_ROLES  = ["admin", "manager"]


def _center_toplevel(win: tk.Toplevel, parent: tk.Widget = None):
    win.update_idletasks()
    if parent and parent.winfo_exists():
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        ww, wh = win.winfo_width(), win.winfo_height()
        x = px + (pw // 2) - (ww // 2)
        y = py + (ph // 2) - (wh // 2)
    else:
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        ww, wh = win.winfo_width(), win.winfo_height()
        x = (sw // 2) - (ww // 2)
        y = (sh // 2) - (wh // 2)
    if x < 0: x = 0
    if y < 0: y = 0
    win.geometry(f"+{x}+{y}")


# ============================================================
# Data Layer
# ============================================================
class ItemFamilyManager:
    def __init__(self):
        self.connection = connect_db()

    def __del__(self):
        try:
            if hasattr(self, 'connection') and self.connection:
                self.connection.close()
        except Exception:
            pass

    def get_remarks_by_item_code(self, item_code):
        if not item_code or len(item_code) < 4 or not self.connection:
            return None
        family_code = item_code[:4]
        cursor = self.connection.cursor()
        try:
            cursor.execute('SELECT remarks FROM "item_families" WHERE item_family = ?', (family_code,))
            row = cursor.fetchone()
            return row['remarks'] if row else None
        except sqlite3.DatabaseError:
            return None
        finally:
            cursor.close()

    def add_item_family(self, family_type, item_family, remarks=None):
        if family_type not in ('log', 'med', 'lib'):
            return False
        if not item_family or len(item_family) != 4:
            return False
        if not self.connection:
            return False
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                'INSERT INTO "item_families" (family_type, item_family, remarks) VALUES (?, ?, ?)',
                (family_type, item_family.upper(), remarks or '')
            )
            self.connection.commit()
            return True
        except sqlite3.DatabaseError:
            return False
        finally:
            cursor.close()

    def update_item_family(self, family_type, item_family, remarks, old_item_family):
        if not self.connection:
            return False
        cursor = self.connection.cursor()
        try:
            cursor.execute(
                'UPDATE "item_families" SET family_type=?, item_family=?, remarks=? WHERE item_family=?',
                (family_type, item_family.upper(), remarks or '', old_item_family)
            )
            self.connection.commit()
            return True
        except sqlite3.DatabaseError:
            return False
        finally:
            cursor.close()

    def delete_item_family(self, item_family):
        if not self.connection:
            return False
        cursor = self.connection.cursor()
        try:
            cursor.execute('DELETE FROM "item_families" WHERE item_family = ?', (item_family,))
            self.connection.commit()
            return True
        except sqlite3.DatabaseError:
            return False
        finally:
            cursor.close()


# ============================================================
# UI Layer
# ============================================================
class ManageItemFamilies(tk.Frame):
    """
    Themed UI for managing item families (log/med/lib).
    Uses custom_popup / custom_askyesno and consistent colors.
    """
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG_MAIN)
        self.app = app
        self.role = getattr(app, "role", "supervisor")
        self.item_family_manager = ItemFamilyManager()
        self.tree = None
        self.entries = {}
        self.status_var = tk.StringVar(value=self.t("ready", fallback="Ready"))
        self._configure_styles()
        self._build_ui()
        self.load_item_families()

    # Translation helper
    def t(self, key, **kwargs):
        return lang.t(f"item_families.{key}", **kwargs)

    # Styles
    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "IF.Treeview",
            background=BG_PANEL,
            fieldbackground=BG_PANEL,
            foreground=COLOR_PRIMARY,
            rowheight=26,
            font=("Helvetica", 10),
            bordercolor=COLOR_BORDER,
            relief="flat"
        )
        style.map("IF.Treeview",
                  background=[("selected", COLOR_ACCENT)],
                  foreground=[("selected", "#FFFFFF")])
        style.configure(
            "IF.Treeview.Heading",
            background="#E5E8EB",
            foreground=COLOR_PRIMARY,
            font=("Helvetica", 11, "bold"),
            relief="flat",
            bordercolor=COLOR_BORDER
        )

    # UI Layout
    def _build_ui(self):
        tk.Label(
            self,
            text=self.t("title", fallback="Manage Item Families"),
            font=("Helvetica", 20, "bold"),
            bg=BG_MAIN,
            fg=COLOR_PRIMARY,
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(12, 8))

        # Table frame
        outer = tk.Frame(self, bg=COLOR_BORDER, bd=1, relief="solid")
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        columns = ("family_type", "item_family", "remarks")
        self.tree = ttk.Treeview(
            outer,
            columns=columns,
            show="headings",
            height=16,
            style="IF.Treeview"
        )

        headers = {
            "family_type": self.t("column.family_type", fallback="Family Type"),
            "item_family": self.t("column.item_family", fallback="Item Family"),
            "remarks": self.t("column.remarks", fallback="Remarks")
        }
        widths = {"family_type": 140, "item_family": 140, "remarks": 480}
        for col in columns:
            self.tree.heading(col, text=headers[col])
            self.tree.column(col, width=widths[col], anchor="w")

        self.tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        # Buttons
        btn_frame = tk.Frame(self, bg=BG_MAIN)
        btn_frame.pack(fill="x", padx=12, pady=(0, 6))

        can_modify = self.role in ALLOWED_ROLES

        def mk_btn(text_key, fallback, cmd, color):
            return tk.Button(
                btn_frame,
                text=self.t(text_key, fallback=fallback),
                command=cmd if can_modify else None,
                bg=color if can_modify else BTN_DISABLED,
                fg="#FFFFFF",
                activebackground=color if can_modify else BTN_DISABLED,
                relief="flat",
                padx=14, pady=6,
                font=("Helvetica", 10, "bold"),
                state="normal" if can_modify else "disabled"
            )

        self.btn_add = mk_btn("add_button", "Add", self.add_item_family, BTN_ADD)
        self.btn_add.pack(side="left", padx=4)

        self.btn_edit = mk_btn("edit_button", "Edit", self.edit_item_family, BTN_EDIT)
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = mk_btn("delete_button", "Delete", self.delete_item_family, BTN_DELETE)
        self.btn_delete.pack(side="left", padx=4)

        if not can_modify:
            custom_popup(
                self,
                lang.t("dialog_titles.restricted", "Restricted"),
                self.t("access_denied", fallback="You don't have permission to manage item families."),
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

    # Data load
    def load_item_families(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
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
            cur.execute('SELECT family_type, item_family, remarks FROM "item_families" ORDER BY item_family')
            rows = cur.fetchall()
            for idx, row in enumerate(rows):
                tag = "alt" if idx % 2 else "norm"
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        row["family_type"] or "",
                        row["item_family"] or "",
                        row["remarks"] or ""
                    ),
                    tags=(tag,)
                )
            self.tree.tag_configure("norm", background=ROW_NORM_COLOR)
            self.tree.tag_configure("alt", background=ROW_ALT_COLOR)
            self.status_var.set(
                self.t("loaded_records", fallback="Loaded {n} records").format(n=len(rows))
            )
        except sqlite3.Error as e:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", fallback="Database error: {err}").format(err=str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # Form helper
    def open_form(self, title, save_callback, initial=None):
        form = tk.Toplevel(self)
        form.title(title)
        form.configure(bg=BG_MAIN)
        form.geometry("480x400")
        form.transient(self)
        form.grab_set()

        tk.Label(
            form,
            text=title,
            font=("Helvetica", 16, "bold"),
            fg=COLOR_PRIMARY,
            bg=BG_MAIN,
            anchor="w"
        ).pack(fill="x", padx=16, pady=(16, 10))

        fields = [
            ("family_type", True),
            ("item_family", True),
            ("remarks", False),
        ]
        self.entries = {}

        def add_field(fname, required):
            label = self.t(fname, fallback=fname.replace("_", " ").title())
            tk.Label(
                form,
                text=f"{label}{' *' if required else ''}:",
                bg=BG_MAIN,
                fg=COLOR_PRIMARY,
                font=("Helvetica", 10),
                anchor="w"
            ).pack(fill="x", padx=18, pady=(6, 0))
            if fname == "family_type":
                cb = ttk.Combobox(form, state="readonly", values=["log", "med", "lib"])
                cb.set((initial and initial.get("family_type")) or "log")
                cb.pack(fill="x", padx=18, pady=(0, 6))
                self.entries[fname] = cb
            else:
                ent = tk.Entry(form, font=("Helvetica", 11), relief="solid", bd=1)
                if initial and fname in initial:
                    ent.insert(0, initial[fname] or "")
                ent.pack(fill="x", padx=18, pady=(0, 6))
                self.entries[fname] = ent

        for fname, req in fields:
            add_field(fname, req)

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
        # Focus first entry
        if isinstance(self.entries.get("item_family"), tk.Entry):
            self.entries["item_family"].focus()

    # Add
    def add_item_family(self):
        if self.role not in ALLOWED_ROLES:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("access_denied", fallback="You don't have permission to manage item families."),
                         "error")
            return

        def save(form):
            family_type = self.entries["family_type"].get().strip()
            item_family = self.entries["item_family"].get().strip().upper()
            remarks = self.entries["remarks"].get().strip() if self.entries["remarks"].get() else ""
            if not family_type or not item_family or len(item_family) != 4:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("required_fields",
                                    fallback="Family Type and Item Family (4 characters) are required."),
                             "error")
                return
            if not self.item_family_manager.add_item_family(family_type, item_family, remarks):
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("add_error", fallback="Failed to add item family."),
                             "error")
                return
            custom_popup(self,
                         lang.t("dialog_titles.success", "Success"),
                         self.t("add_success", fallback="Item family added successfully."),
                         "info")
            form.destroy()
            self.load_item_families()

        title = self.t("add_title", fallback="Add Item Family")
        self.open_form(title, save)

    # Edit
    def edit_item_family(self):
        if self.role not in ALLOWED_ROLES:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("access_denied", fallback="You don't have permission to manage item families."),
                         "error")
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("select_record", fallback="Please select an item family to edit."),
                         "error")
            return
        vals = self.tree.item(sel[0])["values"]
        old_item_family = vals[1]

        initial = {
            "family_type": vals[0],
            "item_family": vals[1],
            "remarks": vals[2]
        }

        def save(form):
            family_type = self.entries["family_type"].get().strip()
            item_family = self.entries["item_family"].get().strip().upper()
            remarks = self.entries["remarks"].get().strip()
            if not family_type or not item_family or len(item_family) != 4:
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("required_fields",
                                    fallback="Family Type and Item Family (4 characters) are required."),
                             "error")
                return
            if not self.item_family_manager.update_item_family(family_type, item_family, remarks, old_item_family):
                custom_popup(form,
                             lang.t("dialog_titles.error", "Error"),
                             self.t("update_error", fallback="Failed to update item family."),
                             "error")
                return
            custom_popup(self,
                         lang.t("dialog_titles.success", "Success"),
                         self.t("update_success", fallback="Item family updated successfully."),
                         "info")
            form.destroy()
            self.load_item_families()

        title = self.t("edit_title", fallback="Edit Item Family")
        self.open_form(title, save, initial)

    # Delete
    def delete_item_family(self):
        if self.role not in ALLOWED_ROLES:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("access_denied", fallback="You don't have permission to manage item families."),
                         "error")
            return
        sel = self.tree.selection()
        if not sel:
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("select_record", fallback="Please select an item family to delete."),
                         "error")
            return
        vals = self.tree.item(sel[0])["values"]
        item_family = vals[1]
        ans = custom_askyesno(
            self,
            lang.t("dialog_titles.confirm", "Confirm"),
            self.t("delete_confirm",
                   fallback="Are you sure you want to delete item family {code}?").format(code=item_family)
        )
        if ans != "yes":
            return
        if not self.item_family_manager.delete_item_family(item_family):
            custom_popup(self,
                         lang.t("dialog_titles.error", "Error"),
                         self.t("delete_error", fallback="Failed to delete item family."),
                         "error")
            return
        custom_popup(self,
                     lang.t("dialog_titles.success", "Success"),
                     self.t("delete_success", fallback="Item family deleted successfully."),
                     "info")
        self.load_item_families()

    # External refresh
    def refresh(self):
        self.load_item_families()


# Standalone test
if __name__ == "__main__":
    root = tk.Tk()
    class DummyApp:
        role = "admin"
        project_title = "IsEPREP"
    dummy = DummyApp()
    root.title("Manage Item Families - Test")
    ManageItemFamilies(root, dummy)
    root.geometry("960x600")
    root.mainloop()