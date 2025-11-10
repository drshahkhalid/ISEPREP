import tkinter as tk
from tkinter import ttk, messagebox
from language_manager import lang
from popup_utils import custom_popup, custom_askyesno, custom_dialog

def get_info_text():
    """
    Generate the INFO_TEXT from translation keys to support multilingual manual.
    """
    sections = []
    
    # Header
    sections.append(f"=== {lang.t('info.manual_title', 'APPLICATION MANUAL (IsEPREP)')} ===")
    sections.append(lang.t('info.version', 'Version: 1.0'))
    sections.append(lang.t('info.author', 'Author: Shah Khalid (e-pool Pharmacy Coordinator, OCG)'))
    sections.append(lang.t('info.purpose', 'Purpose: Structured preparedness & deployment logistics (kits, modules, stock, reporting) with multilingual UI.'))

    sections.append(lang.t('info.purpose', 'Purpose: Structured preparedness & deployment logistics (kits, modules, stock, reporting) with multilingual UI.'))
    sections.append("")
    
    # Introduction
    sections.append("=" * 60)
    sections.append(lang.t('info.intro_title', 'INTRODUCTION'))
    sections.append("=" * 60)
    sections.append(lang.t('info.intro_text', '''IsEPREP helps humanitarian / medical operations manage:
- Item master data (multi-language designations)
- Kit & Module structural composition
- Scenario-based planning (up to 15)
- Stock movements (IN / OUT / Kit receive / Kit dispatch / adjustments)
- Tracking of kit_number and module_number instances
- Visibility via consolidated reports (expiry risk, availability, consumption preparation)
- Role & language aware UI (EN, FR, ES)
Goal: Consistent, auditable, scalable data to support readiness & rapid deployment.'''))
    sections.append("")
    
    # Core Concepts
    sections.append("=" * 60)
    sections.append(lang.t('info.concepts_title', 'CORE CONCEPTS'))
    sections.append("=" * 60)
    sections.append(lang.t('info.concepts_text', '''1. ITEM: Base supply element (medicine / device / logistic component).
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
7. ROLES: admin (@), hq (&), coordinator ((, manager (~), supervisor ($)
8. MULTI-LINGUAL: designation_en / designation_fr / designation_sp prioritized by user language.'''))
    sections.append("")
    
    # Workflow
    sections.append("=" * 60)
    sections.append(lang.t('info.workflow_title', 'RECOMMENDED WORKFLOW'))
    sections.append("=" * 60)
    sections.append(lang.t('info.workflow_text', '''1. LOGIN & LANGUAGE → Choose interface language.
2. PROJECT SETUP → Enter Project Details (name/code) if missing.
3. MASTER DATA → Import or add Items & Item Families.
4. SCENARIOS → Define scenario contexts (max 15).
5. KIT / MODULE COMPOSITION → Build structural relationships.
6. STANDARD LIST → View quantities per scenario; adjust planning.
7. STOCK RECEIPT → Record initial inbound stock.
8. KIT/MODULE INSTANCE → Assign kit_number/module_number.
9. STOCK OUT / DISPATCH → Track consumption or deployment.
10. ADJUSTMENTS → Inventory corrections.
11. REPORTING → Use Statement / Expiry / Availability reports.
12. EXPORT & ARCHIVE → Generate Excel snapshots; perform backups.'''))
    sections.append("")
    
    # Data Model
    sections.append("=" * 60)
    sections.append(lang.t('info.data_model_title', 'DATA MODEL SNAPSHOT (KEY TABLES)'))
    sections.append("=" * 60)
    sections.append(lang.t('info.data_model_text', '''users, project_details, items_list, kit_items, compositions, stock_transactions, stock_data, third_parties, end_users, item_families, views (vw_report_detailed, vw_report_summary, vw_report_expiry, vw_report_required_qty)'''))
    sections.append("")
    
    # Security
    sections.append("=" * 60)
    sections.append(lang.t('info.security_title', 'SECURITY & INTEGRITY'))
    sections.append("=" * 60)
    sections.append(lang.t('info.security_text', '''- Password hashing: PBKDF2 (salted).
- Backups: Use encrypted ZIP (recommended).
- Avoid manual DB edits; use application flows.
- For large deployments: Consider adding indexes.'''))
    sections.append("")
    
    # Reports
    sections.append("=" * 60)
    sections.append(lang.t('info.reports_title', 'REPORTS OVERVIEW'))
    sections.append("=" * 60)
    sections.append(lang.t('info.reports_text', '''1. Stock Statement → Transaction-level detail.
2. Stock Summary → Aggregated per code.
3. Expiry → Buckets: EXPIRED, ALERT_30, ALERT_60, ALERT_90, OK.
4. Required Qty → Planned vs current coverage.
5. Availability → Net available / allocated (planned).
6. Consumption → Outflow analytics (planned).
7. Donations / Loans / Losses → Based on movement_type.
8. Order / Needs → Gap analysis.'''))
    sections.append("")
    
    # Troubleshooting
    sections.append("=" * 60)
    sections.append(lang.t('info.troubleshooting_title', 'TROUBLESHOOTING & FAQ'))
    sections.append("=" * 60)
    sections.append(lang.t('info.troubleshooting_text', '''Q: Items not appearing in Kits Composition search?
A: Ensure type field matches and correct language designation.

Q: Scenario columns missing in Standard List?
A: Confirm scenarios exist and not exceeding 15; reload screen.

Q: Report empty?
A: Verify stock_transactions inserted. Perform Stock In first.

Q: Duplicated kit_number errors?
A: Numbers must be unique within scenario; adjust naming.

Q: Can't edit (greyed out)?
A: Role is restricted; escalate to admin.'''))
    sections.append("")
    
    # Quick Start
    sections.append("=" * 60)
    sections.append(lang.t('info.quick_start_title', 'QUICK START (CHEAT SHEET)'))
    sections.append("=" * 60)
    sections.append(lang.t('info.quick_start_text', '''1. Add Items.
2. Add Scenarios.
3. Build Kit/Module structure.
4. Set scenario quantities (Standard List).
5. Receive stock (Stock In).
6. Assign kit instances (Receive Kit).
7. Dispatch / Outflow (Dispatch Kit / Stock Out).
8. Monitor Expiry & Statement Reports.
9. Export & Backup regularly.'''))
    sections.append("")
    
    # Support
    sections.append("=" * 60)
    sections.append(lang.t('info.support_title', 'SUPPORT & MAINTENANCE'))
    sections.append("=" * 60)
    sections.append(lang.t('info.support_text', '''If a report view is missing:
- Rerun reporting initialization routine.
- Check schema migrations / PRAGMA statements.

For structural anomalies:
- Validate no duplicate treecodes (rebuild if necessary).
- Export current kit_items before manual intervention.

Enjoy using IsEPREP to streamline preparedness logistics!'''))
    sections.append("")
    sections.append("=== END OF MANUAL ===")
    
    return "\n".join(sections)

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
        
        # Generate translatable content
        info_content = get_info_text()
        self.text.insert("1.0", info_content)
        self.text.config(state="disabled")

        scrollbar = tk.Scrollbar(container, command=self.text.yview)
        scrollbar.pack(side="right", fill="y")
        self.text.config(yscrollcommand=scrollbar.set)

    def copy_all(self):
        self.clipboard_clear()
        info_content = get_info_text()
        self.clipboard_append(info_content)
        messagebox.showinfo(
            lang.t("dialog_titles.info", "Info"),
            lang.t("info.copied", "Information copied to clipboard."),
            parent=self
        )

    def save_to_file(self):
        import datetime, os
        filename = f"IsEPREP_User_Manual_{datetime.date.today().isoformat()}.txt"
        try:
            info_content = get_info_text()
            with open(filename, "w", encoding="utf-8") as f:
                f.write(info_content + "\n")
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