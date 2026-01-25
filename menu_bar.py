import tkinter as tk
from tkinter import filedialog
import sqlite3
from db import connect_db

from language_manager import lang
from project_details import ProjectDetailsWindow
from backup_restore import create_backup_zip, restore_backup

# Module windows / frames
from in_kit import StockInKit
from out_kit import StockOutKit
from inv_kit import InventoryKit
from in_ import StockIn
from out import StockOut
from stock_inv import StockInventory
from item_families import ManageItemFamilies
from manage_items import ManageItems
from end_users import ManageEndUsers
from reports import Reports
from stock_transactions import StockTransactions
from standard_list import StandardList
from scenarios import Scenarios
from stock_card import StockCard
from dispatch_kit import StockDispatchKit
from receive_kit import StockReceiveKit

# Additional report / info modules
from expiry_data import StockExpiry
from stock_availability import StockAvailability
from consumption import Consumption
from loans import Loans
from donations import Donations
from losses import Losses
from order import OrderNeeds
from stock_summary import open_stock_summary
from info import AppInfo

from popup_utils import custom_popup, custom_askyesno, custom_dialog  # retained if needed

# Role symbol utilities
from role_map import encode_role, decode_role

import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# ---------------- Helper: Fetch eprep_type ----------------
def fetch_eprep_type(default=None):
    """
    Reads project_details.eprep_type.
    Returns the string (e.g., 'By Kits', 'By Items') or default if missing.
    """
    conn = connect_db()
    if conn is None:
        return default
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA table_info(project_details)")
        cols = {c[1].lower(): c[1] for c in cur.fetchall()}
        if "eprep_type" not in cols:
            return default
        cur.execute("SELECT eprep_type FROM project_details LIMIT 1")
        row = cur.fetchone()
        if not row:
            return default
        val = row[0]
        return val if val else default
    except sqlite3.Error:
        return default
    finally:
        try:
            cur.close()
            conn.close()
        except Exception:
            pass


def create_menu(parent, user_mgmt_cmd, parties_cmd, kits_cmd, change_language, has_project, app=None):
    """
    Menubar with symbol-based role gating plus dynamic hiding based on project_details.eprep_type.

    Args:
        parent: The Toplevel window to attach the menu to
        user_mgmt_cmd: Command for user management
        parties_cmd: Command for parties management
        kits_cmd: Command for kits composition
        change_language: Language change callback
        has_project: Boolean indicating if project exists
        app: The LoginGUI instance (REQUIRED for open_new_window calls)

    eprep_type logic:
      If eprep_type == 'By Kits':
          HIDE: Stock In, Stock Out,
                Generate New Kit/Module from item's stock,
                Break Kit/Module to item's stock.
      If eprep_type == 'By Items':
          HIDE: Generate New Kit/Module from item's stock,
                Break Kit/Module to item's stock,
                Receive Kit,
                Dispatch Kit,
                Inventory Kits/Modules.
      Independent rule:
          Hide "Inventory Kits/Modules" for all users except admin (even if visible under type logic).

    Role Gating (original):
      - Project Details: allowed symbols {"@", "&", "("}
      - User Management: allowed symbols {"@", "&"}
    """
    # CRITICAL FIX: Use app parameter if provided, otherwise fall back to parent
    if app is None:
        app = parent
    
    # Get role from app (LoginGUI instance)
    raw_role = getattr(app, "role", "") or ""
    role_canonical = decode_role(raw_role)
    role_symbol = encode_role(role_canonical)

    eprep_type = fetch_eprep_type(default=None)  # Could be 'By Kits', 'By Items', or None/other

    ALLOW_PROJECT_DETAILS = {"@", "&", "("}
    ALLOW_USER_MGMT = {"@", "&"}

    allow_project_details = role_symbol in ALLOW_PROJECT_DETAILS
    allow_user_management = role_symbol in ALLOW_USER_MGMT

    window = parent if isinstance(parent, (tk.Tk, tk.Toplevel)) else parent.master
    menubar = tk.Menu(window)

    # ---------------- FILE ----------------
    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(
        label=lang.t("menu.file.backup", fallback="Create Backup"),
        command=create_backup_zip
    )
    file_menu.add_command(
        label=lang.t("menu.file.restore", fallback="Restore Backup"),
        command=lambda: restore_backup(
            filedialog.askopenfilename(filetypes=[("Zip Files", "*.zip")])
        )
    )
    if allow_project_details:
        file_menu.add_command(
            label=lang.t("menu.file.project_details", fallback="Project Details"),
            command=lambda: ProjectDetailsWindow(
                window, {"username": app.current_user, "role": role_canonical}
            )
        )
    file_menu.add_separator()
    file_menu.add_command(
        label=lang.t("menu.file.exit", fallback="Exit"),
        command=window.quit
    )
    menubar.add_cascade(
        label=lang.t("menu.file.menu", fallback="File"),
        menu=file_menu
    )

    # ---------------- MANAGE ----------------
    manage_menu = tk.Menu(menubar, tearoff=0)
    if allow_user_management:
        manage_menu.add_command(
            label=lang.t("menu.manage.users", fallback="User Management"),
            command=user_mgmt_cmd
        )
    manage_menu.add_command(
        label=lang.t("menu.manage.scenarios", fallback="Scenario Management"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            Scenarios, lang.t("sidebar.scenarios", "Scenarios")
        )
    )
    manage_menu.add_command(
        label=lang.t("menu.manage.items", fallback="Item Management"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            ManageItems, lang.t("sidebar.manage_items", "Manage Items")
        )
    )
    manage_menu.add_command(
        label=lang.t("menu.manage.third_parties", fallback="Third Party Management"),
        command=parties_cmd
    )
    manage_menu.add_command(
        label=lang.t("menu.manage.end_users", fallback="End User Management"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            ManageEndUsers, lang.t("sidebar.manage_end_users", "Manage End Users")
        )
    )
    manage_menu.add_command(
        label=lang.t("menu.manage.item_families", fallback="Item Family Management"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            ManageItemFamilies, lang.t("sidebar.manage_item_families", "Manage Item Families")
        )
    )
    manage_menu.add_command(
        label=lang.t("menu.manage.standard_list", fallback="Standard Item List"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            StandardList, lang.t("sidebar.standard_list", "Standard List")
        )
    )
    manage_menu.add_command(
        label=lang.t("menu.manage.kits", fallback="Kit Composition Management"),
        command=kits_cmd
    )
    menubar.add_cascade(
        label=lang.t("menu.manage.menu", fallback="Manage"),
        menu=manage_menu
    )

    # ---------------- REPORTS ----------------
    reports_menu = tk.Menu(menubar, tearoff=0)
    reports_menu.add_command(
        label=lang.t("menu.reports.stock_statement", fallback="Stock Statement"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            Reports, lang.t("menu.reports.stock_statement", "Stock Statement")
        )
    )
    reports_menu.add_command(
        label=lang.t("menu.reports.stock_summary", fallback="Stock Summary"),
        command=lambda: open_stock_summary(app, role=role_canonical)  # FIXED: Use app
    )
    reports_menu.add_separator()
    report_defs = [
        ("menu.reports.stock_expiry", StockExpiry, "Stock Expiry"),
        ("menu.reports.stock_availability", StockAvailability, "Stock Availability"),
        ("menu.reports.consumption", Consumption, "Consumption/Receptions"),
        ("menu.reports.loans", Loans, "Loans"),
        ("menu.reports.donations", Donations, "Donations"),
        ("menu.reports.losses", Losses, "Losses"),
        ("menu.reports.order_needs", OrderNeeds, "Order/Needs"),
    ]
    for key, cls, fallback in report_defs:
        reports_menu.add_command(
            label=lang.t(key, fallback=fallback),
            command=lambda c=cls, k=key, fb=fallback: app.open_new_window(  # FIXED: Use app
                c, lang.t(k, fb)
            )
        )
    menubar.add_cascade(
        label=lang.t("menu.reports.menu", fallback="Reports"),
        menu=reports_menu
    )

    # ---------------- TOOLS ----------------
    tools_menu = tk.Menu(menubar, tearoff=0)
    tools_menu.add_command(
        label=lang.t("menu.tools.stock_transactions", fallback="Stock Transactions"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            StockTransactions, lang.t("sidebar.stock_movements", "Stock Movements")
        )
    )
    tools_menu.add_command(
        label=lang.t("menu.tools.stock_card", fallback="Stock Card"),
        command=lambda: app.open_new_window(  # FIXED: Use app instead of parent
            StockCard, lang.t("sidebar.stock_card", "Stock Card")
        )
    )
    menubar.add_cascade(
        label=lang.t("menu.tools.menu", fallback="Tools"),
        menu=tools_menu
    )

    # ---------------- STOCK MOVEMENTS ----------------
    stock_menu = tk.Menu(menubar, tearoff=0)

    # Determine visibility based on eprep_type
    show_stock_in = True
    show_stock_out = True
    show_generate_new = True  # StockInKit
    show_break = True         # StockOutKit
    show_receive_kit = True
    show_dispatch_kit = True
    show_inventory_kits_modules = True  # InventoryKit (kit/module inventory)

    if eprep_type == "By Kits":
        # Hide Stock In/Out & generate/break kit/module from items
        show_stock_in = False
        show_stock_out = False
        show_generate_new = False
        show_break = False
    elif eprep_type == "By Items":
        # Hide kit-oriented flows
        show_generate_new = False
        show_break = False
        show_receive_kit = False
        show_dispatch_kit = False
        show_inventory_kits_modules = False  # regardless of admin in this mode

    # Independent rule: only admin can see Inventory Kits/Modules if still allowed
    if role_canonical.lower() != "admin":
        show_inventory_kits_modules = False

    # Helper to add commands conditionally while controlling separators
    any_added = False
    def add_cmd(show, *args, **kwargs):
        nonlocal any_added
        if show:
            stock_menu.add_command(*args, **kwargs)
            any_added = True

    # Group 1: Stock In / Out
    add_cmd(show_stock_in,
            label=lang.t("menu.stock.stock_in", fallback="Stock In"),
            command=lambda: app.open_new_window(  # FIXED: Use app
                StockIn, lang.t("stock_in.title", "Stock In")
            ))
    add_cmd(show_stock_out,
            label=lang.t("menu.stock.stock_out", fallback="Stock Out"),
            command=lambda: app.open_new_window(  # FIXED: Use app
                StockOut, lang.t("stock_out.title", "Stock Out")
            ))
    # Separator if next group will appear
    if any([show_generate_new, show_break]) and any([show_stock_in, show_stock_out]):
        stock_menu.add_separator()

    # Group 2: Generate / Break kit/module
    group2_added = False
    if show_generate_new:
        stock_menu.add_command(
            label=lang.t("menu.stock.in_to_kit", fallback="Generate New Kit/Module from item's stock"),
            command=lambda: app.open_new_window(  # FIXED: Use app
                StockInKit, lang.t("in_kit.title", "Generate New Kit/Module from item's stock")
            )
        )
        group2_added = True
    if show_break:
        stock_menu.add_command(
            label=lang.t("menu.stock.out_from_kit", fallback="Break Kit/Module to item's stock"),
            command=lambda: app.open_new_window(  # FIXED: Use app
                StockOutKit, lang.t("menu.stock.out_from_kit", "Break Kit/Module to item's stock")
            )
        )
        group2_added = True

    if group2_added and any([show_receive_kit, show_dispatch_kit]):
        stock_menu.add_separator()

    # Group 3: Receive / Dispatch Kit
    group3_added = False
    if show_receive_kit:
        stock_menu.add_command(
            label=lang.t("menu.stock.receive_kit", fallback="Receive Kit"),
            command=lambda: app.open_new_window(  # FIXED: Use app
                StockReceiveKit, lang.t("menu.stock.receive_kit", "Receive Kit")
            )
        )
        group3_added = True
    if show_dispatch_kit:
        stock_menu.add_command(
            label=lang.t("menu.stock.dispatch_kit", fallback="Dispatch Kit"),
            command=lambda: app.open_new_window(  # FIXED: Use app
                StockDispatchKit, lang.t("menu.stock.dispatch_kit", "Dispatch Kit")
            )
        )
        group3_added = True

    # Group 4: Inventory Adjustments (always show Stock Inventory; kit/module inventory conditional)
    if any([show_stock_in, show_stock_out, group2_added, group3_added]):
        stock_menu.add_separator()

    stock_menu.add_command(
        label=lang.t("menu.stock.stock_inventory", fallback="Stock Inventory"),
        command=lambda: app.open_new_window(  # FIXED: Use app
            StockInventory, lang.t("stock_inv.title", "Stock Inventory Adjustment")
        )
    )
    if show_inventory_kits_modules:
        stock_menu.add_command(
            label=lang.t("menu.stock.inventory_kit", fallback="Inventory Kits/Modules"),
            command=lambda: app.open_new_window(  # FIXED: Use app
                InventoryKit, lang.t("menu.stock.inventory_kit", "Inventory Kits/Modules")
            )
        )

    menubar.add_cascade(
        label=lang.t("menu.stock.menu", fallback="Stock Movements"),
        menu=stock_menu
    )

    # ---------------- LANGUAGE ----------------
    if change_language:
        language_menu = tk.Menu(menubar, tearoff=0)
        for code, label in [
            ("en", lang.t("menu.language.english", fallback="English")),
            ("fr", lang.t("menu.language.french", fallback="Français")),
            ("es", lang.t("menu.language.spanish", fallback="Español")),
        ]:
            language_menu.add_command(
                label=label,
                command=lambda c=code: change_language(c)
            )
        menubar.add_cascade(
            label=lang.t("menu.language.menu", fallback="Language"),
            menu=language_menu
        )

    # ---------------- INFO ----------------
    info_menu = tk.Menu(menubar, tearoff=0)
    info_menu.add_command(
        label=lang.t("menu.file.info", fallback="Info"),
        command=lambda: app.open_new_window(AppInfo, lang.t("menu.file.info", "Info"))  # FIXED: Use app
    )
    menubar.add_cascade(
        label=lang.t("menu.info.menu", fallback="Info"),
        menu=info_menu
    )

    window.config(menu=menubar)