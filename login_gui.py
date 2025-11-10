import tkinter as tk
from tkinter import messagebox, ttk, Toplevel
import sqlite3
import os
import logging

from db import connect_db
from manage_users import ManageUsers
from manage_items import ManageItems
from manage_parties import ManageParties
from stock_transactions import StockTransactions
from reports import Reports
from scenarios import Scenarios
from kits_Composition import KitsComposition
from menu_bar import create_menu
from language_manager import lang
from standard_list import StandardList
from stock_inv import StockInventory
from item_families import ManageItemFamilies
from in_ import StockIn
from out import StockOut
from in_kit import StockInKit
from out_kit import StockOutKit
from dispatch_kit import StockDispatchKit
from receive_kit import StockReceiveKit
from stock_card import StockCard
from end_users import ManageEndUsers
from project_details import ProjectDetailsWindow
from dashboard import Dashboard
from stock_availability import StockAvailability
from order import OrderNeeds
from stock_summary import open_stock_summary
from auth_utils import authenticate

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ---------------- Style constants ----------------
BG_LOGIN = "#F0F4F8"
ENTRY_BG = "#FFFFFF"
ENTRY_FG = "#1E3A8A"
BTN_BG = "#5DADE2"
BTN_BG_ACTIVE = "#3498DB"
BOTTOM_BAR_BG = "#1D4ED8"
CONTENT_BG = "#F5F5F5"

DEBUG = False  # Set True for verbose logging


def fetch_eprep_type(default=None):
    conn = connect_db()
    if conn is None:
        return default
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA table_info(project_details)")
        cols = {c[1].lower(): c[1] for c in cur.fetchall()}
        if "eprep_type" not in cols:
            return default
        cur.execute("SELECT eprep_type FROM project_details ORDER BY id DESC LIMIT 1")
        row = cur.fetchone()
        return row[0] if row and row[0] else default
    except sqlite3.Error as e:
        logging.error(f"fetch_eprep_type error: {e}")
        return default
    finally:
        try:
            cur.close()
            conn.close()
        except Exception:
            pass


class LoginGUI(tk.Frame):
    """
    IMPORTANT CHANGE:
      After successful login the dashboard & bottom bar are now created INSIDE this frame
      (self) NOT directly on the master Toplevel. Previously an empty top gap remained
      because the LoginGUI frame (packed to master) still occupied the top area while
      new content frames were siblings of it. Now we destroy login_frame children and
      reuse 'self' as the root container for dashboard content. This removes the blank top band.
    """
    def __init__(self, master):
        super().__init__(master, bg=BG_LOGIN)
        self.master = master
        self.master.title(lang.t("app.title", fallback="IsEPREP"))
        self.master.geometry("1280x1024")
        self.master.minsize(1280, 1024)
        try:
            self.master.state('zoomed')
        except Exception:
            pass

        # Pack this main frame once; everything else happens inside it
        self.pack(fill="both", expand=True)

        # State
        self.current_user = None
        self.role = None
        self.login_frame = None
        self.bottom_bar = None
        self.logged_in = False
        self.has_project = None
        self.project_title = None
        self.windows = []
        self.selected_language = "en"
        self.eprep_type = None

        self.content_frame = None  # (dashboard container when logged in)

        self._fetch_project_meta()
        self.login_screen()

    # ---------------- Project meta ----------------
    def _fetch_project_meta(self):
        conn = connect_db()
        if not conn:
            self.project_title = lang.t("app.title", fallback="IsEPREP")
            self.has_project = False
            self.master.title(self.project_title)
            return
        try:
            cur = conn.cursor()
            cur.execute('SELECT project_name, project_code FROM project_details ORDER BY id DESC LIMIT 1')
            row = cur.fetchone()
            if row:
                self.project_title = f"{row[0]} - {row[1]}"
                self.has_project = True
            else:
                self.project_title = lang.t("app.title", fallback="IsEPREP")
                self.has_project = False
            self.master.title(self.project_title)
        except sqlite3.Error as e:
            logging.error(f"Project meta fetch error: {e}")
            self.project_title = lang.t("app.title", fallback="IsEPREP") + " - DB Error"
            self.has_project = False
            self.master.title(self.project_title)
        finally:
            try:
                cur.close()
                conn.close()
            except Exception:
                pass

    def set_project_title(self):
        self._fetch_project_meta()

    # ---------------- Helpers ----------------
    def get_active_window(self):
        self.windows = [w for w in self.windows if w.winfo_exists()]
        return self.windows[-1] if self.windows else self.master

    def clear_content(self):
        if self.content_frame and self.content_frame.winfo_exists():
            for w in self.content_frame.winfo_children():
                w.destroy()

    # ---------------- Ensure project details ----------------
    def ensure_project_details(self):
        self._fetch_project_meta()
        if self.has_project:
            return True

        messagebox.showinfo(
            lang.t("dialog_titles.setup_required", "Setup Required"),
            lang.t("project_details.setup_required", "Project details setup required."),
            parent=self.get_active_window()
        )
        pd_win = ProjectDetailsWindow(self.master, {"username": self.current_user, "role": self.role})
        self.master.wait_window(pd_win)

        self._fetch_project_meta()
        if not self.has_project:
            messagebox.showwarning(
                lang.t("dialog_titles.warning", "Warning"),
                lang.t("project_details.still_missing", "Project details still missing. Please create them to proceed."),
                parent=self.get_active_window()
            )
            return False
        return True

    # ---------------- Login UI ----------------
    def login_screen(self):
        if self.logged_in:
            return
        # Clean existing children inside this top-level frame
        for child in self.winfo_children():
            child.destroy()

        self.configure(bg=BG_LOGIN)
        self.login_frame = tk.Frame(self, bg=BG_LOGIN)
        self.login_frame.place(relx=0.5, rely=0.5, anchor="center")

        logo_path = "logo.png"
        if os.path.exists(logo_path):
            try:
                self.logo_img = tk.PhotoImage(file=logo_path)
                tk.Label(self.login_frame, image=self.logo_img, bg=BG_LOGIN,
                         borderwidth=0, highlightthickness=0).pack(pady=20)
            except Exception:
                self._fallback_logo()
        else:
            self._fallback_logo()

        tk.Label(
            self.login_frame,
            text=lang.t("login.language_select", fallback="Select Language"),
            font=("Helvetica", 12),
            bg=BG_LOGIN,
            fg="#1E3A8A"
        ).pack(pady=(0, 10))

        self.language_var = tk.StringVar(value=lang.t("menu.language.english", fallback="English"))
        language_options = [
            lang.t("menu.language.english", fallback="English"),
            lang.t("menu.language.french", fallback="Français"),
            lang.t("menu.language.spanish", fallback="Español")
        ]
        self.language_cb = ttk.Combobox(
            self.login_frame, textvariable=self.language_var,
            values=language_options, state="readonly",
            font=("Helvetica", 10), width=16
        )
        self.language_cb.pack(pady=(0, 10))
        self.language_cb.bind("<<ComboboxSelected>>", self.on_language_select)

        self.username_entry = tk.Entry(
            self.login_frame, font=("Helvetica", 14),
            relief="flat", bd=1, bg=ENTRY_BG, fg=ENTRY_FG
        )
        self.username_entry.insert(0, lang.t("login.username", fallback="Username"))
        self.username_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(
            self.username_entry, lang.t("login.username", fallback="Username")))
        self.username_entry.bind("<FocusOut>", lambda e: self._add_placeholder(
            self.username_entry, lang.t("login.username", fallback="Username")))
        self.username_entry.bind("<Return>", lambda e: self.do_login())
        self.username_entry.pack(pady=(10, 10))

        self.password_entry = tk.Entry(
            self.login_frame, show="*",
            font=("Helvetica", 14),
            relief="flat", bd=1, bg=ENTRY_BG, fg=ENTRY_FG
        )
        self.password_entry.insert(0, lang.t("login.password", fallback="Password"))
        self.password_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(
            self.password_entry, lang.t("login.password", fallback="Password")))
        self.password_entry.bind("<FocusOut>", lambda e: self._add_placeholder(
            self.password_entry, lang.t("login.password", fallback="Password")))
        self.password_entry.bind("<Return>", lambda e: self.do_login())
        self.password_entry.pack(pady=(10, 20))

        tk.Button(
            self.login_frame,
            text=lang.t("login.login_button", fallback="Sign In"),
            font=("Helvetica", 14, "bold"),
            bg=BTN_BG, fg="white",
            activebackground=BTN_BG_ACTIVE,
            width=20,
            cursor="hand2",
            command=self.do_login
        ).pack(pady=10)

    def _fallback_logo(self):
        tk.Label(self.login_frame,
                 text=lang.t("app.title", fallback="IsEPREP"),
                 font=("Helvetica", 32, "bold"),
                 fg="#1E3A8A", bg=BG_LOGIN).pack(pady=20)

    # ---------------- Language ----------------
    def on_language_select(self, event=None):
        language_map = {
            lang.t("menu.language.english", fallback="English"): "en",
            lang.t("menu.language.french", fallback="Français"): "fr",
            lang.t("menu.language.spanish", fallback="Español"): "es"
        }
        self.selected_language = language_map.get(self.language_var.get(), "en")
        lang.set_language(self.selected_language)
        self.update_login_labels()

    def update_login_labels(self):
        if not (self.login_frame and self.login_frame.winfo_exists()):
            return
        # Simple refresh; re-render to apply translation
        self.login_screen()

    # ---------------- Placeholders ----------------
    def _clear_placeholder(self, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            if "password" in placeholder.lower():
                entry.config(show="*")

    def _add_placeholder(self, entry, placeholder):
        if entry.get() == "":
            entry.insert(0, placeholder)
            if "password" in placeholder.lower():
                entry.config(show="")

    # ---------------- Login Logic ----------------
    def do_login(self):
        if not (self.login_frame and self.login_frame.winfo_exists()):
            self.login_screen()
            return
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        if (not username or not password or
                username == lang.t("login.username", fallback="Username") or
                password == lang.t("login.password", fallback="Password")):
            messagebox.showerror(
                lang.t("dialog_titles.error", "Error"),
                lang.t("login.error_empty", "Please enter both username and password"),
                parent=self.get_active_window()
            )
            return
        try:
            ok, user_obj, migrated, msg = authenticate(username, password)
        except Exception as e:
            logging.error(f"authenticate() raised: {e}")
            messagebox.showerror(
                lang.t("dialog_titles.error", "Error"),
                f"Authentication error: {e}",
                parent=self.get_active_window()
            )
            return
        if not ok or not user_obj:
            messagebox.showerror(
                lang.t("dialog_titles.error", "Error"),
                lang.t("login.error_invalid", "Invalid username or password"),
                parent=self.get_active_window()
            )
            return
        self.current_user = user_obj.get("username", username)
        self.role = user_obj.get("role", "user")
        pref_lang = user_obj.get("preferred_language", "EN").upper()
        lang_map = {"EN": "en", "FR": "fr", "ES": "es"}
        if pref_lang in lang_map:
            self.selected_language = lang_map[pref_lang]
            lang.set_language(self.selected_language)
        self.logged_in = True
        if not self.ensure_project_details():
            self.logged_in = False
            return
        # Destroy login widgets ONLY (keep self frame)
        if self.login_frame and self.login_frame.winfo_exists():
            self.login_frame.destroy()
        self.dashboard()

    # ---------------- Dashboard ----------------
    def dashboard(self):
        # Remove any existing dashboard content
        for child in self.winfo_children():
            # keep bottom bar if will be recreated
            if child is not self.bottom_bar:
                child.destroy()

        self.set_project_title()
        self.eprep_type = (fetch_eprep_type(default="") or "").strip()
        eprep_norm = self.eprep_type.lower()

        create_menu(
            self,  # menubar still attached to Toplevel
            lambda: self.open_new_window(ManageUsers, lang.t("sidebar.manage_users", "Manage Users")),
            lambda: self.open_new_window(lambda parent, app: ManageParties(parent, app, "third"),
                                         lang.t("sidebar.manage_third_parties", "Manage Third Parties")),
            lambda: self.open_new_window(self.load_kits_composition,
                                         lang.t("sidebar.kits_composition", "Kits Composition")),
            self.change_language,
            self.has_project
        )

        # Content area occupies all vertical space except bottom bar
        self.content_frame = tk.Frame(self, bg=CONTENT_BG)
        self.content_frame.pack(side="top", fill="both", expand=True)

        dash = Dashboard(self.content_frame, self)
        dash.pack(fill="both", expand=True)

        # Bottom bar now inside self (not master)
        if self.bottom_bar and self.bottom_bar.winfo_exists():
            self.bottom_bar.destroy()
        self.bottom_bar = tk.Frame(self, bg=BOTTOM_BAR_BG, height=78)
        self.bottom_bar.pack(side="bottom", fill="x")

        # Determine mode
        is_by_kits = (eprep_norm == "by kits")
        is_by_items = (eprep_norm == "by items")

        # Buttons configuration
        buttons = [
            # (icon, label, command, show)
            ("📦", lang.t("sidebar.manage_items", "Items"),
             lambda: self.open_new_window(ManageItems, lang.t("sidebar.manage_items", "Items")), True),

            ("📑", lang.t("sidebar.standard_list", "Standard List"),
             lambda: self.open_new_window(StandardList, lang.t("sidebar.standard_list", "Standard List")),
             not is_by_kits),

            ("🧰", lang.t("sidebar.kits_composition", "Kits Composition"),
             lambda: self.open_new_window(self.load_kits_composition, lang.t("sidebar.kits_composition", "Kits Composition")),
             not is_by_items),

            ("📥", lang.t("stock_in.title", "Stock In"),
             lambda: self.open_new_window(StockIn, lang.t("stock_in.title", "Stock In")),
             not is_by_items),

            ("📤", lang.t("stock_out.title", "Stock Out"),
             lambda: self.open_new_window(StockOut, lang.t("stock_out.title", "Stock Out")),
             not is_by_items),

            ("🚚", lang.t("menu.stock.receive_kit", "Receive Kit"),
             lambda: self.open_new_window(StockReceiveKit, lang.t("menu.stock.receive_kit", "Receive Kit")),
             not is_by_kits),

            ("🚛", lang.t("menu.stock.dispatch_kit", "Dispatch Kit"),
             lambda: self.open_new_window(StockDispatchKit, lang.t("menu.stock.dispatch_kit", "Dispatch Kit")),
             not is_by_kits),

            ("🗂️", lang.t("sidebar.stock_card", "Stock Card"),
             lambda: self.open_new_window(StockCard, lang.t("sidebar.stock_card", "Stock Card")),
             True),

            ("📊", lang.t("menu.reports.stock_availability", "Stock Availability"),
             lambda: self.open_new_window(StockAvailability, lang.t("menu.reports.stock_availability", "Stock Availability")),
             True),

            ("🧾", lang.t("menu.reports.stock_summary", "Stock Summary"),
             lambda: open_stock_summary(self, role=self.role),
             True),

            ("🛒", lang.t("menu.reports.order_needs", "Order"),
             lambda: self.open_new_window(OrderNeeds, lang.t("menu.reports.order_needs", "Order / Needs")),
             True),

            ("🚪", lang.t("sidebar.logout", "Sign Out"),
             self.logout, True),
        ]

        visible = [b for b in buttons if b[3]]
        total = len(visible)
        for idx, (icon, label, cmd, _) in enumerate(visible):
            state = "normal" if self.has_project or label == lang.t("sidebar.logout", "Sign Out") else "disabled"
            btn = tk.Button(
                self.bottom_bar, text=icon,
                font=("Helvetica", 26),  # Larger icon
                bg=BOTTOM_BAR_BG, fg="white",
                activebackground="#2563EB",
                command=cmd,
                bd=0, relief="flat",
                state=state,
                cursor="hand2" if state == "normal" else "arrow"
            )
            relx = (idx + 0.5) / total
            btn.place(relx=relx, rely=0.30, anchor="center", width=70, height=54)

            tk.Label(self.bottom_bar,
                     text=label,
                     bg=BOTTOM_BAR_BG,
                     fg="white",
                     font=("Helvetica", 9),
                     wraplength=90,
                     justify="center")\
                .place(relx=relx, rely=0.83, anchor="center")

        self.update_idletasks()

    # ---------------- Logout ----------------
    def logout(self):
        for w in self.windows:
            try:
                if w.winfo_exists():
                    w.destroy()
            except Exception:
                pass
        self.windows.clear()
        self.logged_in = False
        self.current_user = None
        self.role = None
        self.eprep_type = None
        # Clear all children
        for child in self.winfo_children():
            child.destroy()
        self.login_screen()

    # ---------------- Window helpers ----------------
    def open_new_window(self, module, title):
        role_modules = [
            StockIn, StockOut, StockInventory, StockInKit, StockOutKit,
            StockDispatchKit, StockReceiveKit, StockCard
        ]
        special_modules = {
            ManageParties: lambda parent, app: ManageParties(parent, app, "third"),
            "ManagePartiesEnd": lambda parent, app: ManageParties(parent, app, "end")
        }

        for window in self.windows[:]:
            if not window.winfo_exists():
                self.windows.remove(window)
            elif window.title() == f"{self.project_title} - {title}":
                self.close_window(window)

        try:
            new_window = Toplevel(self.master)
            new_window.title(f"{self.project_title} - {title}")
            new_window.geometry("1280x1024")
            new_window.minsize(1280, 1024)
            try:
                new_window.state('zoomed')
            except Exception:
                pass
            new_window.configure(bg=CONTENT_BG)
            new_window.protocol("WM_DELETE_WINDOW",
                                lambda: self.close_window(new_window))
            self.windows.append(new_window)
            offset = len(self.windows) * 30
            new_window.geometry(f"+{50 + offset}+{50 + offset}")

            if module in role_modules:
                frame = module(new_window, self, role=self.role)
            elif module == StandardList:
                frame = module(new_window, self)
            elif module in special_modules:
                frame = special_modules[module](new_window, self)
            elif module == self.load_kits_composition:
                frame = module(new_window, self)
            elif callable(module) and module.__name__ == "<lambda>":
                frame = module(new_window, self)
            else:
                frame = module(new_window, self)

            frame.pack(fill="both", expand=True)
            if hasattr(frame, 'initialize_ui'):
                frame.initialize_ui()
            new_window.lift()
            return new_window
        except Exception as e:
            logging.error(f"Error opening {title}: {e}")
            try:
                if 'new_window' in locals() and new_window.winfo_exists():
                    new_window.destroy()
            except Exception:
                pass
            messagebox.showerror(
                lang.t("dialog_titles.error", "Error"),
                lang.t("error.open_window", f"Failed to open {title}: {str(e)}"),
                parent=self.get_active_window()
            )
            return None

    def close_window(self, window):
        if window in self.windows:
            self.windows.remove(window)
        try:
            window.destroy()
        except tk.TclError as e:
            logging.error(f"Error closing window: {e}")

    def change_language(self, lang_code):
        self.selected_language = lang_code
        lang.set_language(lang_code)
        self.set_project_title()
        # rebuild dashboard in current layout
        if self.logged_in:
            self.dashboard()
        else:
            self.login_screen()

    def load_kits_composition(self, parent, app):
        return KitsComposition(parent, app)

    def refresh_dashboard_after_project_save(self):
        self.has_project = True
        self.set_project_title()
        self.dashboard()

    def load_stock_in_out(self):
        self.clear_content()
        self.master.title(self.project_title or lang.t("app.title", fallback="IsEPREP"))
        tab_frame = tk.Frame(self.content_frame, bg=BG_LOGIN)
        tab_frame.pack(fill="x", pady=5)

        def show_in():
            self.open_new_window(StockIn, lang.t("stock_in.title", fallback="Stock In"))

        def show_out():
            self.open_new_window(StockOut, lang.t("stock_out.title", fallback="Stock Out"))

        def show_inventory():
            self.open_new_window(StockInventory, lang.t("stock_inv.title", fallback="Stock Inventory Adjustment"))

        tk.Button(tab_frame, text=lang.t("stock_in.title", fallback="Stock In"),
                  bg=BOTTOM_BAR_BG, fg="white", command=show_in).pack(side="left", padx=5, pady=5)
        tk.Button(tab_frame, text=lang.t("stock_out.title", fallback="Stock Out"),
                  bg=BOTTOM_BAR_BG, fg="white", command=show_out).pack(side="left", padx=5, pady=5)
        tk.Button(tab_frame, text=lang.t("stock_inv.title", fallback="Stock Inventory Adjustment"),
                  bg=BOTTOM_BAR_BG, fg="white", command=show_inventory).pack(side="left", padx=5, pady=5)

        show_in()

    def export_to_excel(self):
        messagebox.showinfo(
            lang.t("dialog_titles.info", fallback="Info"),
            lang.t("standard_list.export_excel", fallback="Export to Excel") + " " +
            lang.t("reports.export_success", fallback="Export successful."),
            parent=self.get_active_window()
        )


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    mainwin = Toplevel(root)
    app = LoginGUI(mainwin)
    mainwin.protocol("WM_DELETE_WINDOW", root.quit)
    root.mainloop()
