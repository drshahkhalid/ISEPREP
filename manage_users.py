import tkinter as tk
from tkinter import ttk
import sqlite3
import hashlib
import base64
import secrets

from db import connect_db
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno, custom_dialog

# ============================================================
# IMPORT CENTRALIZED THEME (NEW)
# ============================================================
from theme_config import AppTheme, configure_tree_tags

# New: role symbol mapping (can also be imported from role_map if you prefer)
ROLE_TO_SYMBOL = {
    "admin": "@",
    "hq": "&",
    "coordinator": "(",
    "manager": "~",
    "supervisor": "$",
}
SYMBOL_TO_ROLE = {v: k for k, v in ROLE_TO_SYMBOL.items()}

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

ROLES_ALLOWED   = ["admin", "hq", "coordinator", "manager", "supervisor"]

# Note: Your table constraint uses ('EN','FR','SP')
# Earlier code used 'ES' for Spanish. We'll normalize:
# - Display 'ES', but store 'SP' if the table enforces 'SP'
# Adjust LANG_DB_VALUE/LANG_DISPLAY_VALUE maps below if needed.
LANG_DB_VALUES     = ["EN", "FR", "SP"]
DISPLAY_TO_DB_LANG = {"EN": "EN", "FR": "FR", "ES": "SP", "SP": "SP"}
DB_TO_DISPLAY_LANG = {"EN": "EN", "FR": "FR", "SP": "ES"}  # Show 'ES' to users
LANG_OPTIONS_DISPLAY = ["EN", "FR", "ES"]  # what user sees

# ------------------------------------------------------------
# Password Hashing Utilities (PBKDF2-HMAC-SHA256)
# Stored format: pbkdf2_sha256$<iterations>$<base64(salt)>$<base64(hash)>
# ------------------------------------------------------------
DEFAULT_ITERATIONS = 260000

def hash_password(password: str, iterations: int = DEFAULT_ITERATIONS) -> str:
    if not isinstance(password, str):
        raise TypeError("Password must be a string")
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return f"pbkdf2_sha256${iterations}${base64.b64encode(salt).decode()}${base64.b64encode(dk).decode()}"

def verify_password(password: str, stored_hash: str) -> bool:
    """Return True if password matches stored_hash."""
    try:
        algo, iterations, salt_b64, hash_b64 = stored_hash.split("$", 3)
        if algo != "pbkdf2_sha256":
            return False
        iterations = int(iterations)
        salt = base64.b64decode(salt_b64)
        original_hash = base64.b64decode(hash_b64)
        test_hash = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
        return secrets.compare_digest(original_hash, test_hash)
    except Exception:
        return False

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

class ManageUsers(tk.Frame):
    """
    Manage Users screen (updated for 'symbol' column in DB):
      - Adds / migrates preferred_language & symbol columns if needed
      - Shows symbol and role side by side
      - Always stores canonical role in 'role' column, and symbol in 'symbol'
      - Language display shows ES while DB stores SP (if constraint uses SP)
      - Editing: blank password -> keeps existing hash
    """
    def __init__(self, parent, app):
        super().__init__(parent, bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        self.app = app
        self.role = self._decode_role(getattr(app, "role", "supervisor"))
        self.tree = None
        self.pack(fill="both", expand=True)
        self._ensure_preferred_language_column()
        self._ensure_symbol_column()
        self._configure_styles()
        self._build_ui()
        self.load_users()

    def t(self, key, fallback=""):
        return lang.t(f"manage_users.{key}", fallback or key)

    # --------------------------------------------------------
    # Role symbol helpers
    # --------------------------------------------------------
    def _encode_role(self, role: str) -> str:
        return ROLE_TO_SYMBOL.get(role.lower(), role)

    def _decode_role(self, value: str) -> str:
        low = (value or "").lower()
        if low in ROLE_TO_SYMBOL:
            return low
        if value in SYMBOL_TO_ROLE:
            return SYMBOL_TO_ROLE[value]
        return value

    # --------------------------------------------------------
    # DB Migration helpers
    # --------------------------------------------------------
    def _ensure_preferred_language_column(self):
        conn = connect_db()
        if conn is None:
            return
        cur = conn.cursor()
        try:
            cur.execute("PRAGMA table_info(users)")
            cols = [r[1].lower() for r in cur.fetchall()]
            if "preferred_language" not in cols:
                cur.execute("ALTER TABLE users ADD COLUMN preferred_language TEXT")
                conn.commit()
            # Normalize existing null/empty -> EN
            cur.execute("""
                UPDATE users
                   SET preferred_language='EN'
                 WHERE preferred_language IS NULL
                    OR TRIM(preferred_language)=''
            """)
            conn.commit()
        except sqlite3.Error:
            conn.rollback()
        finally:
            cur.close()
            conn.close()

    def _ensure_symbol_column(self):
        """
        Adds symbol column if missing, then backfills using role mapping.
        """
        conn = connect_db()
        if conn is None:
            return
        cur = conn.cursor()
        try:
            cur.execute("PRAGMA table_info(users)")
            cols = [r[1].lower() for r in cur.fetchall()]
            if "symbol" not in cols:
                cur.execute("ALTER TABLE users ADD COLUMN symbol TEXT")
                conn.commit()
            # Backfill any NULL symbol values
            cur.execute('SELECT user_id, role, symbol FROM "users"')
            rows = cur.fetchall()
            for user_id, role, sym in rows:
                if not sym or sym.strip() == "":
                    new_sym = self._encode_role(role)
                    cur.execute('UPDATE "users" SET symbol=? WHERE user_id=?', (new_sym, user_id))
            conn.commit()
        except sqlite3.Error:
            conn.rollback()
        finally:
            cur.close()
            conn.close()

    # --------------------------------------------------------
    # Styles (UPDATED: Removed theme_use call, using AppTheme colors)
    # --------------------------------------------------------
    def _configure_styles(self):
        # Global theme already applied by login_gui
        # Just configure module-specific styles here
        style = ttk.Style()
        # REMOVED: style.theme_use("clam") - already applied globally

        style.configure(
            "Users.Treeview",
            background=AppTheme.BG_PANEL,  # UPDATED: Use AppTheme
            fieldbackground=AppTheme.BG_PANEL,  # UPDATED: Use AppTheme
            foreground=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            rowheight=AppTheme.TREE_ROW_HEIGHT,  # UPDATED: Use AppTheme
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),  # UPDATED: Use AppTheme
            bordercolor=AppTheme.COLOR_BORDER,  # UPDATED: Use AppTheme
            relief="flat"
        )
        style.map("Users.Treeview",
                  background=[("selected", AppTheme.COLOR_ACCENT)],  # UPDATED: Use AppTheme
                  foreground=[("selected", AppTheme.TEXT_WHITE)])  # UPDATED: Use AppTheme

        style.configure(
            "Users.Treeview.Heading",
            background="#E5E8EB",
            foreground=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HEADING, "bold"),  # UPDATED: Use AppTheme
            relief="flat",
            bordercolor=AppTheme.COLOR_BORDER  # UPDATED: Use AppTheme
        )
        style.layout("Users.Treeview", [
            ("Users.Treeview.treearea", {"sticky": "nswe"})
        ])
        style.configure("Users.TEntry", font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL))  # UPDATED: Use AppTheme
        style.configure("Users.TCombobox", font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL))  # UPDATED: Use AppTheme

    # --------------------------------------------------------
    # UI (UPDATED: All color references use AppTheme)
    # --------------------------------------------------------
    def _build_ui(self):
        tk.Label(
            self,
            text=self.t("title", "Manage Users"),
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_HUGE, "bold"),  # UPDATED: Use AppTheme
            bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
            fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            anchor="w",
            justify="left"
        ).pack(fill="x", padx=12, pady=(12, 8))

        outer = tk.Frame(self, bg=AppTheme.COLOR_BORDER, bd=1, relief="solid")  # UPDATED: Use AppTheme
        outer.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        # Added symbol column
        columns = ("user_id", "username", "symbol", "role", "preferred_language")
        self.tree = ttk.Treeview(
            outer,
            columns=columns,
            show="headings",
            height=12,
            style="Users.Treeview"
        )
        self.tree.column("user_id", width=70, anchor="w")
        self.tree.column("username", width=180, anchor="w")
        self.tree.column("symbol", width=70, anchor="center")
        self.tree.column("role", width=140, anchor="w")
        self.tree.column("preferred_language", width=140, anchor="w")

        self.tree.heading("user_id", text=self.t("column.id", "ID"))
        self.tree.heading("username", text=self.t("column.username", "Username"))
        self.tree.heading("symbol", text=self.t("column.symbol", "Symbol"))
        self.tree.heading("role", text=self.t("column.role", "Role"))
        self.tree.heading("preferred_language", text=self.t("column.preferred_language", "Preferred Language"))

        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(outer, orient="vertical", command=self.tree.yview)
        sb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb.set)

        btn_frame = tk.Frame(self, bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        btn_frame.pack(fill="x", padx=12, pady=(0, 6))

        can_modify = self.role in ["admin", "manager", "hq", "coordinator"]

        self.btn_add = tk.Button(
            btn_frame,
            text=self.t("add_button", "Add User"),
            bg=AppTheme.BTN_SUCCESS if can_modify else AppTheme.BTN_DISABLED,  # UPDATED: Use AppTheme
            fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
            activebackground="#1E874B",
            relief="flat",
            padx=12, pady=6,
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),  # UPDATED: Use AppTheme
            command=self.add_user if can_modify else None
        )
        self.btn_add.pack(side="left", padx=4)

        self.btn_edit = tk.Button(
            btn_frame,
            text=self.t("edit_button", "Edit User"),
            bg=AppTheme.BTN_WARNING if can_modify else AppTheme.BTN_DISABLED,  # UPDATED: Use AppTheme
            fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
            activebackground="#1F5D82",
            relief="flat",
            padx=12, pady=6,
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),  # UPDATED: Use AppTheme
            command=self.edit_user if can_modify else None
        )
        self.btn_edit.pack(side="left", padx=4)

        self.btn_delete = tk.Button(
            btn_frame,
            text=self.t("delete_button", "Delete User"),
            bg=AppTheme.BTN_DANGER if can_modify else AppTheme.BTN_DISABLED,  # UPDATED: Use AppTheme
            fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
            activebackground="#962D22",
            relief="flat",
            padx=12, pady=6,
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),  # UPDATED: Use AppTheme
            command=self.delete_user if can_modify else None
        )
        self.btn_delete.pack(side="left", padx=4)

        if not can_modify:
            custom_popup(
                self,
                lang.t("dialog_titles.restricted", "Restricted"),
                self.t("access_denied", "You don't have permission to manage users."),
                "warning"
            )

        self.status_var = tk.StringVar(value=self.t("ready", "Ready"))
        tk.Label(
            self,
            textvariable=self.status_var,
            anchor="w",
            bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
            fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
            relief="sunken"
        ).pack(fill="x", padx=12, pady=(0, 10))

    # --------------------------------------------------------
    # Load users (UPDATED: Row colors use AppTheme)
    # --------------------------------------------------------
    def load_users(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", "Database connection failed"),
                         "error")
            return
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        try:
            cur.execute('SELECT user_id, username, role, preferred_language, symbol FROM "users" ORDER BY user_id')
            rows = cur.fetchall()
            for idx, row in enumerate(rows):
                tag = "alt" if idx % 2 else "norm"
                # Preferred language normalization (DB -> display)
                plang_db = (row["preferred_language"] or "EN").upper()
                plang_disp = DB_TO_DISPLAY_LANG.get(plang_db, "EN")
                role_canon = self._decode_role(row["role"])
                sym = row["symbol"] if row["symbol"] else self._encode_role(role_canon)
                self.tree.insert(
                    "",
                    "end",
                    values=(row["user_id"], row["username"], sym, role_canon, plang_disp),
                    tags=(tag,)
                )
            self.tree.tag_configure("norm", background=AppTheme.ROW_NORM)  # UPDATED: Use AppTheme
            self.tree.tag_configure("alt", background=AppTheme.ROW_ALT)  # UPDATED: Use AppTheme
            self.status_var.set(self.t("loaded", "Loaded users: ") + str(len(rows)))
        except sqlite3.Error as e:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", "Database error: {}").format(str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # --------------------------------------------------------
    # Add user
    # --------------------------------------------------------
    def add_user(self):
        self._open_user_form(edit=False)

    # --------------------------------------------------------
    # Edit user
    # --------------------------------------------------------
    def edit_user(self):
        sel = self.tree.selection()
        if not sel:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         self.t("select_user", "Select a user to edit"),
                         "error")
            return
        user_id = self.tree.item(sel[0])["values"][0]
        self._open_user_form(edit=True, user_id=user_id)

    # --------------------------------------------------------
    # Delete user
    # --------------------------------------------------------
    def delete_user(self):
        sel = self.tree.selection()
        if not sel:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         self.t("select_user", "Select a user to edit or delete"),
                         "error")
            return
        user_id = self.tree.item(sel[0])["values"][0]

        if self.role not in ["admin", "manager", "hq"]:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         self.t("access_denied", "You don't have permission to manage users."),
                         "error")
            return

        ans = custom_askyesno(
            self,
            lang.t("dialog_titles.confirm", "Confirm"),
            self.t("confirm_delete", "Delete this user?")
        )
        if ans != "yes":
            return

        conn = connect_db()
        if conn is None:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", "Database error: {}").format("No connection"),
                         "error")
            return
        cur = conn.cursor()
        try:
            cur.execute('DELETE FROM "users" WHERE user_id = ?', (user_id,))
            conn.commit()
            custom_popup(self, lang.t("dialog_titles.success", "Success"),
                         self.t("delete_success", "User deleted successfully."),
                         "info")
            self.load_users()
        except sqlite3.Error as e:
            custom_popup(self, lang.t("dialog_titles.error", "Error"),
                         self.t("db_error", "Database error: {}").format(str(e)),
                         "error")
        finally:
            cur.close()
            conn.close()

    # --------------------------------------------------------
    # User form (add/edit) (UPDATED: All color references use AppTheme)
    # --------------------------------------------------------
    def _open_user_form(self, edit=False, user_id=None):
        form = tk.Toplevel(self)
        form.configure(bg=AppTheme.BG_MAIN)  # UPDATED: Use AppTheme
        form.title(
            self.t("edit_title", "Edit User") if edit
            else self.t("add_title", "Add User")
        )
        form.geometry("420x470")
        form.transient(self)
        form.grab_set()

        def lbl(text):
            return tk.Label(form, text=text, 
                          bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
                          fg=AppTheme.COLOR_PRIMARY,  # UPDATED: Use AppTheme
                          font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL),  # UPDATED: Use AppTheme
                          anchor="w", justify="left")

        lbl(self.t("username", "Username")).pack(fill="x", padx=18, pady=(18, 2))
        username_entry = ttk.Entry(form, style="Users.TEntry")
        username_entry.pack(fill="x", padx=18, pady=(0, 10))

        lbl(self.t("password", "Password")).pack(fill="x", padx=18, pady=(0, 2))
        password_entry = ttk.Entry(form, show="*", style="Users.TEntry")
        password_entry.pack(fill="x", padx=18, pady=(0, 4))
        tk.Label(
            form,
            text=self.t("password_hint_edit", "Leave blank to keep current password") if edit else "",
            bg=AppTheme.BG_MAIN,  # UPDATED: Use AppTheme
            fg="#6B7280", 
            anchor="w", 
            justify="left", 
            font=(AppTheme.FONT_FAMILY, 8, "italic")  # UPDATED: Use AppTheme
        ).pack(fill="x", padx=18, pady=(0, 6))

        lbl(self.t("role", "Role")).pack(fill="x", padx=18, pady=(0, 2))
        role_cb = ttk.Combobox(form, state="readonly",
                               values=ROLES_ALLOWED,
                               style="Users.TCombobox")
        role_cb.set("supervisor")
        role_cb.pack(fill="x", padx=18, pady=(0, 10))

        lbl(self.t("preferred_language", "Preferred Language")).pack(fill="x", padx=18, pady=(0, 2))
        plang_cb = ttk.Combobox(form, state="readonly",
                                values=LANG_OPTIONS_DISPLAY,
                                style="Users.TCombobox")
        plang_cb.set("EN")
        plang_cb.pack(fill="x", padx=18, pady=(0, 10))

        # Display symbol (read only) for clarity
        lbl(self.t("symbol", "Role Symbol")).pack(fill="x", padx=18, pady=(0, 2))
        symbol_var = tk.StringVar(value=self._encode_role("supervisor"))
        symbol_entry = ttk.Entry(form, textvariable=symbol_var, state="readonly")
        symbol_entry.pack(fill="x", padx=18, pady=(0, 12))

        def update_symbol(*_):
            sel_role = role_cb.get().strip().lower()
            symbol_var.set(self._encode_role(sel_role))

        role_cb.bind("<<ComboboxSelected>>", update_symbol)

        existing_hash = None
        existing_plang_db = "EN"
        existing_symbol = None

        if edit and user_id is not None:
            conn = connect_db()
            if conn:
                conn.row_factory = sqlite3.Row
                cur = conn.cursor()
                try:
                    cur.execute('SELECT username, role, password_hash, preferred_language, symbol FROM "users" WHERE user_id = ?', (user_id,))
                    row = cur.fetchone()
                    if row:
                        username_entry.insert(0, row["username"])
                        existing_hash = row["password_hash"]
                        role_canon = self._decode_role(row["role"])
                        if role_canon in ROLES_ALLOWED:
                            role_cb.set(role_canon)
                            symbol_var.set(self._encode_role(role_canon))
                        plang_db = (row["preferred_language"] or "EN").upper()
                        existing_plang_db = plang_db if plang_db in LANG_DB_VALUES else "EN"
                        plang_cb.set(DB_TO_DISPLAY_LANG.get(existing_plang_db, "EN"))
                        existing_symbol = row["symbol"]
                except sqlite3.Error as e:
                    custom_popup(form, lang.t("dialog_titles.error", "Error"),
                                 self.t("db_error", "Database error: {}").format(str(e)),
                                 "error")
                finally:
                    cur.close()
                    conn.close()

        def save():
            uname = username_entry.get().strip()
            pw = password_entry.get()
            r = role_cb.get().strip().lower()
            plang_display = plang_cb.get().strip().upper()

            if not uname or not r or not plang_display:
                custom_popup(form, lang.t("dialog_titles.error", "Error"),
                             self.t("required_fields", "All fields are required.") if not edit
                             else self.t("required_fields_partial", "Username, Role and Preferred Language are required."),
                             "error")
                return
            if r not in ROLES_ALLOWED:
                custom_popup(form, lang.t("dialog_titles.error", "Error"),
                             self.t("invalid_role", "Role must be admin, hq, coordinator, manager, or supervisor."),
                             "error")
                return
            # Convert display language (ES) to DB (SP) if needed
            plang_db = DISPLAY_TO_DB_LANG.get(plang_display, "EN")
            if plang_db not in LANG_DB_VALUES:
                custom_popup(form, lang.t("dialog_titles.error", "Error"),
                             self.t("invalid_language", "Preferred Language must be EN, FR or ES."),
                             "error")
                return

            # Derive symbol
            role_symbol = self._encode_role(r)

            conn2 = connect_db()
            if conn2 is None:
                custom_popup(form, lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", "Database error: {}").format("No connection"),
                             "error")
                return
            cur2 = conn2.cursor()
            try:
                if edit and user_id is not None:
                    # Decide on hash (keep existing if blank)
                    new_hash = hash_password(pw) if pw.strip() else existing_hash
                    if not new_hash:
                        custom_popup(form, lang.t("dialog_titles.error", "Error"),
                                     self.t("password_missing", "Password hash missing; please enter a password."),
                                     "error")
                        return
                    cur2.execute(
                        'UPDATE "users" SET username=?, password_hash=?, role=?, preferred_language=?, symbol=? WHERE user_id=?',
                        (uname, new_hash, r, plang_db, role_symbol, user_id)
                    )
                    msg_key = "edit_success"
                    default_msg = "User updated successfully."
                else:
                    if not pw.strip():
                        custom_popup(form, lang.t("dialog_titles.error", "Error"),
                                     self.t("password_required", "Password is required for new user."),
                                     "error")
                        return
                    # Check duplicate
                    cur2.execute('SELECT 1 FROM "users" WHERE username=?', (uname,))
                    if cur2.fetchone():
                        custom_popup(form, lang.t("dialog_titles.error", "Error"),
                                     self.t("duplicate_username", "Username already exists."),
                                     "error")
                        return
                    pw_hash = hash_password(pw)
                    cur2.execute(
                        'INSERT INTO "users" (username, password_hash, role, preferred_language, symbol) VALUES (?, ?, ?, ?, ?)',
                        (uname, pw_hash, r, plang_db, role_symbol)
                    )
                    msg_key = "add_success"
                    default_msg = "User added successfully."

                conn2.commit()
                custom_popup(form,
                             lang.t("dialog_titles.success", "Success"),
                             self.t(msg_key, default_msg),
                             "info")
                form.destroy()
                self.load_users()
            except sqlite3.Error as e:
                conn2.rollback()
                custom_popup(form, lang.t("dialog_titles.error", "Error"),
                             self.t("db_error", "Database error: {}").format(str(e)),
                             "error")
            finally:
                cur2.close()
                conn2.close()

        save_btn = tk.Button(
            form,
            text=self.t("save_button", "Save"),
            bg=AppTheme.COLOR_ACCENT,  # UPDATED: Use AppTheme
            fg=AppTheme.TEXT_WHITE,  # UPDATED: Use AppTheme
            activebackground="#1D4ED8",
            relief="flat",
            padx=12,
            pady=6,
            font=(AppTheme.FONT_FAMILY, AppTheme.FONT_SIZE_NORMAL, "bold"),  # UPDATED: Use AppTheme
            command=save
        )
        save_btn.pack(fill="x", padx=18, pady=(4, 12))
        form.after(50, lambda: _center_toplevel(form, self))

    def refresh(self):
        self.load_users()

if __name__ == "__main__":
    root = tk.Tk()
    class DummyApp: pass
    app = DummyApp()
    app.role = "admin"
    root.title("Manage Users")
    ManageUsers(root, app)
    root.geometry("880x560")
    root.mainloop()