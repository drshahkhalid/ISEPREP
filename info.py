import tkinter as tk
from tkinter import ttk, messagebox
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno, custom_dialog

# NOTE FOR TRANSLATORS:
# The large INFO_TEXT block below is intentionally structured with clear
# UPPERCASE SECTION HEADINGS and simple punctuation so it can be copied
# into translation JSON files (FR / ES) or processed by a translation
# workflow. Avoid altering the heading tokens when translating.
#
# If you add new sections, keep the same style:
# === SECTION NAME ===
#
# You may also choose to externalize this text into language files
# later (e.g. lang.t("info.manual.introduction")); for now it is inline
# for simplicity.

INFO_TEXT = """
=== APPLICATION MANUAL (IsEPREP) ===
Version: 1.0
Author: Shah Khalid (e-pool Pharmacy Coordinator, OCG)
Purpose: Structured preparedness & deployment logistics (kits, modules, stock, reporting) with multilingual UI.

This manual provides:
- High‑level overview
- Core concepts & roles
- End‑to‑end workflow
- Data model snapshot
- Functional module guide
- Common tasks
- Security & integrity notes
- Troubleshooting & FAQ
- ASCII flow & architecture diagrams
- Future enhancements & glossary

============================================================
INTRODUCTION
============================================================
IsEPREP helps humanitarian / medical operations manage:
- Item master data (multi‑language designations)
- Kit & Module structural composition
- Scenario-based planning (up to 15)
- Stock movements (IN / OUT / Kit receive / Kit dispatch / adjustments)
- Tracking of kit_number and module_number instances
- Visibility via consolidated reports (expiry risk, availability, consumption preparation)
- Role & language aware UI (EN, FR, ES)
Goal: Consistent, auditable, scalable data to support readiness & rapid deployment.

============================================================
CORE CONCEPTS
============================================================
1. ITEM: Base supply element (medicine / device / logistic component).
2. MODULE: Logical grouping of Items (e.g. Dressing Module).
3. KIT: Higher-level grouping (e.g. Trauma Kit) composed of MODULES and standalone ITEMS.
4. SCENARIO: Contextual operational environment (e.g. Cholera Outbreak, Field Hospital).
5. COMPOSITION LAYERS:
   - STRUCTURAL: Kit -> Module -> Item blueprint (no quantities in stock yet).
   - INSTANCE / STOCK: Real physical stock with transaction history and expiry.
6. UNIQUE IDENTIFIERS:
   - code: Item code or Kit/Module structural code.
   - treecode: 11-digit hierarchical code (SS PPP MMM III).
   - unique_id / unique_id_2: Encodes scenario and path for traceability.
7. ROLES (Canonical -> Symbol):
   - admin (@)
   - hq (&)
   - coordinator (
   - manager (~)  (restricted editing in many modules)
   - supervisor ($) (read-only in many management contexts)
8. MULTI-LINGUAL: designation_en / designation_fr / designation_sp prioritized by user language.

============================================================
ROLE PERMISSION SUMMARY (OBFUSCATION VIA SYMBOLS)
============================================================
(Exact permissions may evolve; always enforced server-side / logic-side.)

Symbol  Canonical    Typical Capabilities (Illustrative)
@       admin        Full access (users, scenarios, items, kits, reports, backups)
&       hq           Nearly full; strategic oversight, manage users/project
(       coordinator  Build compositions, manage stock & reports
~       manager      Operative review; restricted from modifying sensitive compositions
$       supervisor   Read-only for management screens; can view reports and perform limited lookups

Note: Symbols are an obfuscation layer only, not a security boundary.

============================================================
HIGH-LEVEL WORKFLOW (RECOMMENDED SEQUENCE)
============================================================
1. LOGIN & LANGUAGE → Choose interface language.
2. PROJECT SETUP → Enter Project Details (name/code) if missing.
3. MASTER DATA → Import or add Items & Item Families.
4. SCENARIOS → Define scenario contexts (max 15).
5. KIT / MODULE COMPOSITION → Build structural relationships (Kits -> Modules -> Items).
6. STANDARD LIST (Compositions) → View quantities per scenario; adjust planning.
7. STOCK RECEIPT → Record initial inbound stock (Stock In or Receive Kit).
8. KIT/MODULE INSTANCE → Assign kit_number/module_number when assembling or dispatching.
9. STOCK OUT / DISPATCH → Track consumption or deployment.
10. ADJUSTMENTS → Inventory corrections (physical vs system counts).
11. REPORTING → Use Statement / Expiry / Availability / Required Qty / etc.
12. EXPORT & ARCHIVE → Generate Excel snapshots; perform backups.

============================================================
ASCII FLOW DIAGRAM (LOGICAL)
============================================================
          +-------------+
          |  LOGIN / UI |
          +------+------+ 
                 |
                 v
        +--------+---------+
        |  PROJECT DETAILS |
        +--------+---------+
                 |
                 v
        +--------+---------+
        |  MASTER ITEMS    |
        +--------+---------+
                 |
                 v
        +--------+---------+
        |  SCENARIOS       |
        +--------+---------+
                 |
                 v
        +--------+---------+
        | KITS / MODULES   |  (STRUCTURAL)
        +--------+---------+
                 |
                 v
        +--------+---------+
        | STANDARD LIST    |  (Scenario Qty Planning)
        +--------+---------+
                 |
   +-------------+----------------------------+
   |                                          |
   v                                          v
+--+------------------+        +--------------+-------------+
| STOCK IN / RECEIVE  |        | INVENTORY ADJUST / CORRECT |
+--+------------------+        +--------------+-------------+
            |                                 |
            +---------------+-----------------+
                            |
                            v
                      +-----+------+
                      |  REPORTS   |
                      +-----+------+
                            |
                            v
                      +-----+------+
                      |  EXPORTS   |
                      +------------+

============================================================
DATA MODEL SNAPSHOT (KEY TABLES)
============================================================
users                (username, password_hash, role, symbol, preferred_language)
project_details      (project_name, code, metadata)
items_list           (code, type, designations, pack, prices, shelf_life, remarks)
kit_items            (scenario_id, kit, module, item, code, std_qty, level, treecode)
compositions         (code, scenario_id, quantity, unique_id_2[a..o], remarks, updated_at)
stock_transactions   (date, time, unique_id, code, qty_in, qty_out, types, third_party, end_user, expiry_date, remarks, movement_type, kit, module, scenario)
stock_data           (unique_id, code, qty_in, qty_out, exp_date, kit_number, module_number)
third_parties        (...)
end_users            (...)
item_families        (...)
views (dynamic)      (vw_report_detailed, vw_report_summary, vw_report_expiry, vw_report_required_qty)

============================================================
STANDARD LIST VS KIT COMPOSITION
============================================================
- KIT COMPOSITION (kits_Composition): Structural relationships.
- STANDARD LIST (standard_list): Scenario-specific planned quantities (editable per scenario if role allows).
- Changes in structural side do not automatically adjust scenario quantities; planning is distinct.

============================================================
SECURITY & INTEGRITY
============================================================
- Password hashing: PBKDF2 (salted).
- Pepper (recommended) can be added via environment variable.
- Symbols DO NOT equal security; authorization always checks canonical role.
- Backups: Use encrypted ZIP (recommended; future integration).
- Avoid manual DB edits; use application flows to preserve derived fields (treecode, unique ids).
- For large deployments: Consider adding indexes (code, scenario_id, date, expiry_date).

============================================================
COMMON TASK RECIPES
============================================================
A) ADD NEW ITEM
   1. Manage Items → Add
   2. Provide code (>=8–9 chars recommended) & English designation.
   3. Optional price, shelf life, remarks ('exp' triggers expiry tracking).
B) BUILD KIT
   1. Open Kits Composition
   2. Select Scenario
   3. Right-click Scenario → Add Kit
   4. Right-click Kit → Add Module / Item
C) RECEIVE STOCK (WITHOUT KIT STRUCTURE)
   1. Stock In → Enter item code(s), quantities, expiry.
D) ASSEMBLE KIT INSTANCES
   1. Receive Kit OR In to Kit (generate kit_number)
   2. Out from Kit / Dispatch Kit when sending.
E) UPDATE SCENARIO QUANTITIES
   1. Standard List → Double-click scenario column cell (if permitted) → Enter numeric qty.
F) CHECK EXPIRY RISK
   1. Reports → Expiry → Filter / Export.

============================================================
REPORTS OVERVIEW
============================================================
1. Stock Statement     → Transaction-level detail.
2. Stock Summary       → Aggregated per code.
3. Expiry              → Buckets: EXPIRED, ALERT_30, ALERT_60, ALERT_90, OK.
4. Required Qty        → Planned (scenario composition) vs current coverage (future metrics).
5. Availability        → (Planned future) net available / allocated.
6. Consumption         → (Planned future) outflow analytics.
7. Donations / Loans / Losses → Based on movement_type or remarks.
8. Order / Needs       → Gap analysis (future enhancement).

============================================================
KEY UI SCREENS
============================================================
- Manage Users: Create / edit user credentials & roles (admin/hq only).
- Manage Items: CRUD item master data; import / export Excel.
- Item Families: Group items; append family remarks.
- Scenarios: Add up to 15 scenario contexts.
- Kits Composition: Hierarchical tree editing via context menu.
- Standard List: Scenario quantity matrix; remarks column.
- Stock In / Out: Direct movements.
- In to Kit / Out from Kit: Structural assignment / breakdown.
- Receive Kit / Dispatch Kit: Instance-based operational flow.
- Inventory Adjustment: Correct physical vs system variance.
- Reports: Analytical outputs with export.
- Info: This manual & system synopsis.

============================================================
ASCII STATE TRANSITIONS (SIMPLIFIED)
============================================================
[Item Master] -> [Kit Composition] -> [Scenario Quantities] -> [Stock In] -> [Inventory / Dispatch / Reports]

============================================================
TROUBLESHOOTING & FAQ
============================================================
Q: Items not appearing in Kits Composition search?
A: Ensure type field matches (Item / Module / Kit) and correct language designation.

Q: Scenario columns missing in Standard List?
A: Confirm scenarios exist and not exceeding 15; reload screen.

Q: Expiry not visible?
A: Include 'exp' or interpreted remark convention; ensure reporting views initialized.

Q: Report empty?
A: Verify stock_transactions inserted. Perform Stock In first.

Q: Duplicated kit_number or module_number errors?
A: Numbers must be unique within scenario (or combined context); adjust naming.

Q: Can't edit (greyed out)?
A: Role is restricted (manager '~' or supervisor '$'); escalate to admin.

============================================================
PERFORMANCE TIPS
============================================================
- Limit giant imports: chunk large Excel loads.
- Use search filters in Reports to narrow data sets.
- Periodically archive old transactions if DB grows large (future tool).

============================================================
SECURITY BEST PRACTICES
============================================================
- Enforce strong passwords (≥10 chars).
- Add environment PEPPER for hashing.
- Restrict filesystem permissions on DB & backups.
- Consider packaging (PyInstaller) to reduce casual code edits.
- Maintain audit log (future planned) for regulatory traceability.

============================================================
FUTURE ENHANCEMENTS (ROADMAP IDEAS)
============================================================
- Dashboard with coverage gauges & expiry heat maps.
- Role granular permissions (view vs edit per module).
- Integrated audit trail & action log.
- Batch CSV import wizard with validation preview.
- Offline / remote replication module.
- Embedded charts (matplotlib / ttkbootstrap).
- Automated reorder / needs calculation.

============================================================
GLOSSARY (SHORT)
============================================================
Scenario: Operational context container.
Kit / Module: Structured reusable blueprints.
Treecode: 11-digit hierarchical locator.
Coverage %: Planned vs actual readiness (future).
Instance: Physical kit/module with assigned number.
Movement Type: Categorization of IN/OUT (donation, consumption, loan).

============================================================
KEYBOARD / USABILITY TIPS
============================================================
- Double-click cells (where permitted) to edit scenario qty or remarks.
- Enter confirms popups; Esc cancels most popups.
- Search boxes accept partial code or designation.
- Use consistent naming for kit_number/module_number (e.g. KIT001, MOD001).

============================================================
BACKUP & EXPORT STRATEGY
============================================================
- Perform database backup before large imports.
- Export Standard List & Reports periodically (Excel) for audit.
- Keep backups encrypted and off primary workstation for redundancy.

============================================================
MINIMAL QUICK START (CHEAT SHEET)
============================================================
1. Add Items.
2. Add Scenarios.
3. Build Kit/Module structure.
4. Set scenario quantities (Standard List).
5. Receive stock (Stock In).
6. Assign kit instances (Receive Kit).
7. Dispatch / Outflow (Dispatch Kit / Stock Out).
8. Monitor Expiry & Statement Reports.
9. Export & Backup regularly.

============================================================
SUPPORT & MAINTENANCE
============================================================
If a report view is missing:
- Rerun reporting initialization routine.
- Check schema migrations / PRAGMA statements.
For structural anomalies:
- Validate no duplicate treecodes (rebuild if necessary).
- Export current kit_items before manual intervention.

Enjoy using IsEPREP to streamline preparedness logistics!

=== END OF MANUAL ===
"""

class AppInfo(tk.Frame):
    def __init__(self, parent, app, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.app = app
        self.configure(bg="#FFFFFF")
        self._build()

    def _build(self):
        header = tk.Label(
            self,
            text=lang.t("info.title", "Application Information"),
            font=("Helvetica", 18, "bold"),
            bg="#FFFFFF"
        )
        header.pack(pady=(20, 10))

        toolbar = tk.Frame(self, bg="#FFFFFF")
        toolbar.pack(fill="x", padx=20, pady=(0, 6))

        copy_btn = tk.Button(
            toolbar,
            text=lang.t("info.copy", "Copy All"),
            command=self.copy_all,
            bg="#2563EB",
            fg="white",
            relief="flat",
            padx=14,
            pady=4,
            font=("Helvetica", 10, "bold")
        )
        copy_btn.pack(side="left")

        save_btn = tk.Button(
            toolbar,
            text=lang.t("info.save", "Save to File"),
            command=self.save_to_file,
            bg="#374151",
            fg="white",
            relief="flat",
            padx=14,
            pady=4,
            font=("Helvetica", 10)
        )
        save_btn.pack(side="left", padx=(8, 0))

        container = tk.Frame(self, bg="#FFFFFF")
        container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.text = tk.Text(
            container,
            wrap="word",
            font=("Helvetica", 10),
            bg="#FFFFFF",
            fg="#111111",
            relief="solid",
            bd=1
        )
        self.text.pack(side="left", fill="both", expand=True)
        self.text.insert("1.0", INFO_TEXT.strip())
        self.text.config(state="disabled")

        scrollbar = tk.Scrollbar(container, command=self.text.yview)
        scrollbar.pack(side="right", fill="y")
        self.text.config(yscrollcommand=scrollbar.set)

    def copy_all(self):
        self.clipboard_clear()
        self.clipboard_append(INFO_TEXT.strip())
        messagebox.showinfo(
            lang.t("dialog_titles.info", "Info"),
            lang.t("info.copied", "Information copied to clipboard."),
            parent=self
        )

    def save_to_file(self):
        import datetime, os
        filename = f"IsEPREP_User_Manual_{datetime.date.today().isoformat()}.txt"
        try:
            with open(filename, "w", encoding="utf-8") as f:
                f.write(INFO_TEXT.strip() + "\n")
            messagebox.showinfo(
                lang.t("dialog_titles.info", "Info"),
                lang.t("info.saved", f"Saved as {filename}"),
                parent=self
            )
        except Exception as e:
            messagebox.showerror(
                lang.t("dialog_titles.error", "Error"),
                lang.t("info.save_error", f"Failed to save: {e}"),
                parent=self
            )

if __name__ == "__main__":
    root = tk.Tk()
    root.title("IsEPREP - Application Info")
    class Dummy: pass
    d = Dummy()
    d.role = "admin"
    AppInfo(root, d).pack(fill="both", expand=True)
    root.geometry("900x700")
    root.mainloop()