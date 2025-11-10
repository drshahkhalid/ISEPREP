"""
This version = v2.4 + Initial pane sizing ratios patch.

Adds:
  * Automatic initial sizing of PanedWindow panes:
        - Project (left): 16% of total window width (PROJECT_PANE_RATIO)
        - Actions (right): 18% of total window width (ACTION_PANE_RATIO)
        - Center (scenarios): remaining width
  * User can still drag sashes; once user drags (B1-Motion / release) we never auto-resize again.
  * Safe retries until the toplevel has a realized (non‑trivial) width.
  * All functionality from v2.4 retained (action gaps, reduced font size, item-level expiry, etc.).

Configuration additions:
  ACTION_PANE_RATIO (default 0.18)
  PROJECT_PANE_RATIO (default 0.16)

If you want to re-apply automatic sizing after a manual resize, set:
    self._user_resized_panes = False
    self._set_initial_pane_sizes()
manually (e.g. from a menu command).

"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk
from datetime import datetime, date, timedelta
import sqlite3
from collections import defaultdict
from db import connect_db
from language_manager import lang
from manage_items import get_item_description

# ---------------- Configuration ----------------
SHORT_EXPIRY_DAYS = 20
EXPIRY_ACTION_WINDOW_DAYS = 180
STANDARD_LIST_STALE_DAYS = 730
SCENARIO_STALE_DAYS = 1095
AUTO_REFRESH_MS = 5 * 60 * 1000

MIN_STOCK_FOR_ACTION = 1
ALWAYS_INCLUDE_ORPHAN_STOCK = True
SHOW_SCENARIO_EXPIRY_LINES = False
SHOW_DEBUG_PANEL = True
DEBUG = False

# UI / Layout
ENABLE_PANED_WINDOW = True
USE_TEXT_ACTIONS = True
ACTION_WRAP = "word"
WRAP_WIDTH = 58
EXPIRED_COLOR = "#B91C1C"
SOON_COLOR = "#92400E"
ACTIONS_BG = "#FDF5F5"
PANEL_BG = "#FFFFFF"
LEFT_BG  = "#E6F2FA"
SCENARIO_HEADER_BG = "#D9EEF8"
BORDER_COLOR = "#1992D4"

ACTION_FONT_SIZE = 9                 # reduced (was 10)
ACTION_EXTRA_BLANK_LINES = 1         # blank lines after each action line for spacing

# Pane ratio configuration (new)
ACTION_PANE_RATIO = 0.18             # 18% of window width
PROJECT_PANE_RATIO = 0.16            # 16% of window width

FONT_TITLE = ("Helvetica", 18, "bold")
FONT_SCENARIO = ("Helvetica", 16, "bold")
FONT_LABEL = ("Helvetica", 10, "bold")
FONT_VALUE = ("Helvetica", 10)
FONT_ACTION = ("Helvetica", ACTION_FONT_SIZE)

# Date formats accepted
_DATE_PATTERNS = [
    ("%Y-%m-%d", None),
    ("%d/%m/%Y", "/"),
    ("%Y/%m/%d", "/"),
    ("%d-%m-%Y", "-"),
]


def parse_flexible_date(s: str):
    if not s or s in ("None", ""):
        return None
    s = s.strip()
    for fmt, _sep in _DATE_PATTERNS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


class Dashboard(tk.Frame):
    def __init__(self, parent, app, *args, **kwargs):
        super().__init__(parent, bg=PANEL_BG, *args, **kwargs)
        self.app = app
        self._building = False
        self._expiring_items = []
        self._diag = {}
        self._paned = None
        self._user_resized_panes = False
        self.actions_text = None
        self.debug_panel = None
        self._build_layout()
        self.refresh()
        self.after(AUTO_REFRESH_MS, self._auto_refresh)

    # ---------- Initial Pane Sizing Helpers (NEW) ----------
    def _set_initial_pane_sizes(self, retry=0):
        """
        Compute initial sash positions based on configured ratios.
        Only runs if:
          * PanedWindow exists
          * User has not manually resized yet
        Retries a few times until a meaningful width is available.
        """
        if not getattr(self, "_paned", None):
            return
        if self._user_resized_panes:
            return

        total = self.winfo_width()
        if total <= 1 and retry < 25:
            # window not realized yet
            self.after(120, lambda: self._set_initial_pane_sizes(retry + 1))
            return
        elif total <= 1:
            return  # give up quietly

        action_ratio = max(0.05, min(0.40, ACTION_PANE_RATIO))
        project_ratio = max(0.05, min(0.40, PROJECT_PANE_RATIO))
        remaining = 1.0 - (action_ratio + project_ratio)
        if remaining < 0.20:  # ensure center retains at least 20%
            scale = (1.0 - 0.20) / (action_ratio + project_ratio)
            action_ratio *= scale
            project_ratio *= scale
            remaining = 0.20

        project_width = int(total * project_ratio)
        action_width = int(total * action_ratio)
        center_width = max(50, total - project_width - action_width)

        # Place sashes
        try:
            self._paned.sash_place(0, project_width, 0)
            self._paned.sash_place(1, project_width + center_width, 0)
            # Optional min sizes
            self._paned.paneconfig(self.project_frame, minsize=int(project_width * 0.5))
            self._paned.paneconfig(self.actions_frame, minsize=int(action_width * 0.5))
            self._paned.paneconfig(self.center_shell, minsize=200)
        except Exception:
            pass

    def _install_paned_resize_monitor(self):
        """Mark when user moves a sash to prevent future auto-resize."""
        if not getattr(self, "_paned", None):
            return

        def on_user_drag(_event):
            self._user_resized_panes = True

        self._paned.bind("<B1-Motion>", on_user_drag, add="+")
        self._paned.bind("<ButtonRelease-1>", on_user_drag, add="+")

    # ---------- Layout ----------
    def _build_layout(self):
        if ENABLE_PANED_WINDOW:
            self._paned = tk.PanedWindow(self, orient="horizontal", sashwidth=6, bg=PANEL_BG)
            self._paned.pack(fill="both", expand=True)
            self.project_frame = tk.Frame(self._paned, bg=LEFT_BG, bd=1, relief="solid")
            self.center_shell = tk.Frame(self._paned, bg=PANEL_BG, bd=0)
            self.actions_frame = tk.Frame(self._paned, bg=ACTIONS_BG, bd=1, relief="solid")
            self._paned.add(self.project_frame, minsize=220)
            self._paned.add(self.center_shell, minsize=400)
            self._paned.add(self.actions_frame, minsize=240)
            # Schedule initial sizing & install monitor
            self.after(150, self._set_initial_pane_sizes)
            self._install_paned_resize_monitor()
        else:
            self.columnconfigure(1, weight=1)
            self.rowconfigure(0, weight=1)
            self.project_frame = tk.Frame(self, bg=LEFT_BG, bd=1, relief="solid")
            self.project_frame.grid(row=0, column=0, sticky="nsw")
            self.center_shell = tk.Frame(self, bg=PANEL_BG)
            self.center_shell.grid(row=0, column=1, sticky="nsew")
            self.actions_frame = tk.Frame(self, bg=ACTIONS_BG, bd=1, relief="solid")
            self.actions_frame.grid(row=0, column=2, sticky="nse")

        # Project panel
        self.project_frame.columnconfigure(0, weight=1)
        tk.Label(self.project_frame,
                 text=lang.t("dashboard.project_details", fallback="Project Details"),
                 bg=LEFT_BG, fg="#093A54", font=FONT_TITLE, anchor="w")\
            .pack(fill="x", padx=8, pady=(8,4))
        self.project_details_text = tk.Text(self.project_frame, height=16, width=40,
                                            bg=LEFT_BG, bd=0, wrap="word", font=("Helvetica",10))
        self.project_details_text.pack(fill="both", expand=True, padx=8, pady=(0,6))
        self.project_details_text.config(state="disabled")
        tk.Button(self.project_frame,
                  text=lang.t("dashboard.refresh", fallback="Refresh"),
                  bg="#1992D4", fg="white", relief="flat",
                  command=self.refresh).pack(padx=8, pady=(0,8), anchor="e")

        # Actions panel
        self.actions_frame.columnconfigure(0, weight=1)
        header = tk.Frame(self.actions_frame, bg=ACTIONS_BG)
        header.pack(fill="x", padx=8, pady=(8,4))
        tk.Label(header,
                 text=lang.t("dashboard.actions_needed", fallback="Actions Needed"),
                 bg=ACTIONS_BG, fg="#7A0019", font=FONT_TITLE, anchor="w")\
            .pack(side="left", fill="x", expand=True)

        if USE_TEXT_ACTIONS:
            text_container = tk.Frame(self.actions_frame, bg=ACTIONS_BG)
            text_container.pack(fill="both", expand=True, padx=6, pady=(0,4))
            yscroll = ttk.Scrollbar(text_container, orient="vertical")
            self.actions_text = tk.Text(
                text_container,
                wrap=ACTION_WRAP,
                width=WRAP_WIDTH,
                bg=ACTIONS_BG,
                bd=0,
                padx=4,
                pady=4,
                font=FONT_ACTION,
                relief="flat"
            )
            self.actions_text.pack(side="left", fill="both", expand=True)
            yscroll.config(command=self.actions_text.yview)
            yscroll.pack(side="right", fill="y")
            self.actions_text.configure(yscrollcommand=yscroll.set, state="disabled")
            self.actions_text.tag_configure("expired",
                                            foreground=EXPIRED_COLOR,
                                            font=("Helvetica", ACTION_FONT_SIZE, "bold"))
            self.actions_text.tag_configure("soon",
                                            foreground=SOON_COLOR,
                                            font=("Helvetica", ACTION_FONT_SIZE))
            self.actions_text.tag_configure("normal",
                                            foreground="#111827",
                                            font=("Helvetica", ACTION_FONT_SIZE))
        else:
            self.actions_list = tk.Listbox(self.actions_frame, bg=ACTIONS_BG,
                                           fg="#7A0019", activestyle="none",
                                           highlightthickness=0, borderwidth=0,
                                           font=FONT_ACTION, width=46)
            self.actions_list.pack(fill="both", expand=True, padx=8, pady=(0,4))

        if SHOW_DEBUG_PANEL:
            self.debug_panel = tk.Label(self.actions_frame, bg=ACTIONS_BG,
                                        fg="#374151", font=("Helvetica",8,"italic"),
                                        anchor="w", justify="left")
            self.debug_panel.pack(fill="x", padx=8, pady=(0,6))

        # Center scrollable scenarios
        self.center_shell.rowconfigure(0, weight=1)
        self.center_shell.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(self.center_shell, bg=PANEL_BG, highlightthickness=0)
        self.scenario_scroll = ttk.Scrollbar(self.center_shell, orient="vertical",
                                             command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scenario_scroll.set)
        self.inner_frame = tk.Frame(self.canvas, bg=PANEL_BG)
        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas_window = self.canvas.create_window((0,0), window=self.inner_frame, anchor="nw")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scenario_scroll.grid(row=0, column=1, sticky="ns")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.inner_frame.bind("<Enter>", lambda e: self._bind_mousewheel())
        self.inner_frame.bind("<Leave>", lambda e: self._unbind_mousewheel())

    # ---------- Mousewheel ----------
    def _bind_mousewheel(self):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_mousewheel(self):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        delta = -1 * (event.delta // 120) if event.delta else (1 if event.num == 5 else -1)
        self.canvas.yview_scroll(delta, "units")

    def _clear_scenarios_ui(self):
        for w in self.inner_frame.winfo_children():
            w.destroy()

    def _reset_scroll(self):
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0)

    # ---------- Refresh ----------
    def _auto_refresh(self):
        if self.winfo_exists():
            self.refresh()
            self.after(AUTO_REFRESH_MS, self._auto_refresh)

    def refresh(self):
        if self._building:
            return
        self._building = True
        try:
            proj_info = self._fetch_project_info()
            self._update_project_panel(proj_info)
            scenarios = self._fetch_scenarios()
            metrics = self._compute_metrics(scenarios)
            self._render_scenarios(metrics)
            actions = self._compute_actions(metrics, proj_info)
            self._render_actions(actions)
        finally:
            self._building = False

    # ---------- Project Info ----------
    def _fetch_project_info(self):
        info = {
            "project_name": lang.t("dashboard.unknown", fallback="Unknown"),
            "project_code": "",
            "updated_at": None,
            "scenario_count": 0,
            "total_kits": 0,
            "total_modules": 0,
            "total_items": 0,
            "standard_last_update": None
        }
        conn = connect_db()
        if not conn: return info
        c = conn.cursor()
        try:
            c.execute("PRAGMA table_info(project_details)")
            pc = {r[1].lower(): r[1] for r in c.fetchall()}
            if {"project_name","project_code"}.issubset(pc.keys()):
                c.execute("SELECT project_name, project_code, updated_at FROM project_details ORDER BY id DESC LIMIT 1")
                r = c.fetchone()
                if r:
                    info["project_name"] = r[0] or info["project_name"]
                    info["project_code"] = r[1] or ""
                    info["updated_at"] = r[2]

            c.execute("PRAGMA table_info(scenarios)")
            sc = {r[1].lower(): r[1] for r in c.fetchall()}
            if "scenario_id" in sc:
                c.execute("SELECT COUNT(*) FROM scenarios")
                info["scenario_count"] = c.fetchone()[0]

            c.execute("PRAGMA table_info(kit_items)")
            kc = {r[1].lower(): r[1] for r in c.fetchall()}
            if {"kit","module","item"}.issubset(kc.keys()):
                c.execute("SELECT kit,module,item FROM kit_items")
                for kit, module, item in c.fetchall():
                    if kit and not module and not item:
                        info["total_kits"] += 1
                    elif kit and module and not item:
                        info["total_modules"] += 1
                    elif kit and module and item:
                        info["total_items"] += 1
                if "updated_at" in kc:
                    c.execute("SELECT MAX(updated_at) FROM kit_items")
                    info["standard_last_update"] = c.fetchone()[0]
            if not info["standard_last_update"]:
                c.execute("PRAGMA table_info(std_qty_helper)")
                sq = {r[1].lower(): r[1] for r in c.fetchall()}
                if "updated_at" in sq:
                    c.execute("SELECT MAX(updated_at) FROM std_qty_helper")
                    info["standard_last_update"] = c.fetchone()[0]
        finally:
            c.close()
            conn.close()
        return info

    # ---------- Scenarios ----------
    def _fetch_scenarios(self):
        res = []
        conn = connect_db()
        if not conn: return res
        c = conn.cursor()
        try:
            c.execute("PRAGMA table_info(scenarios)")
            cols = {r[1].lower(): r[1] for r in c.fetchall()}
            if "scenario_id" not in cols or "name" not in cols:
                return res
            opt = [col for col in ["activity_type","target_population","stock_location",
                                   "responsible_person","remarks","updated_at"] if col in cols]
            if opt:
                c.execute(f"SELECT scenario_id,name,{','.join(opt)} FROM scenarios ORDER BY name")
            else:
                c.execute("SELECT scenario_id,name FROM scenarios ORDER BY name")
            for row in c.fetchall():
                base = {
                    "scenario_id": str(row[0]).strip(),
                    "name": row[1],
                    "activity_type": "",
                    "target_population": "",
                    "stock_location": "",
                    "responsible_person": "",
                    "remarks": "",
                    "scenario_updated_at": None
                }
                for i,col in enumerate(opt, start=2):
                    if col == "updated_at":
                        base["scenario_updated_at"] = row[i]
                    else:
                        base[col] = row[i] or ""
                res.append(base)
        finally:
            c.close()
            conn.close()
        return res

    # ---------- Metrics & Expiring Items ----------
    def _compute_metrics(self, scenarios):
        self._expiring_items = []
        self._diag = {"total_stock_rows":0,"candidate_item_rows":0,"qualifying_action_rows":0,
                      "expired_count":0,"soon_count":0,"bad_date_format":0,"orphan_rows_included":0}

        if not scenarios: return []
        scenario_ids = {s["scenario_id"] for s in scenarios}

        items_std = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
        kit_std_sum = defaultdict(lambda: defaultdict(int))
        kit_stock_sum = defaultdict(lambda: defaultdict(int))
        module_std_sum = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
        module_stock_sum = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))

        std_qty_col = None

        conn = connect_db()
        if not conn: return scenarios
        c = conn.cursor()
        try:
            # kit_items
            c.execute("PRAGMA table_info(kit_items)")
            kit_cols = {r[1].lower(): r[1] for r in c.fetchall()}
            if {"scenario_id","kit","module","item"}.issubset(kit_cols.keys()):
                for cand in ["std_qty","standard_qty","quantity","qty"]:
                    if cand in kit_cols:
                        std_qty_col = kit_cols[cand]; break
                cols = ["scenario_id","kit","module","item"]
                if std_qty_col: cols.append(std_qty_col)
                c.execute(f"SELECT {', '.join(cols)} FROM kit_items")
                for row in c.fetchall():
                    sid, kit, module, item = row[0:4]
                    if sid is None: continue
                    sid = str(sid).strip()
                    if sid not in scenario_ids: continue
                    if kit and module and item:
                        try: std_v = int(row[4]) if std_qty_col else 0
                        except: std_v = 0
                        items_std[sid][kit][(module,item)] += std_v

            # stock_data
            c.execute("PRAGMA table_info(stock_data)")
            sd_cols = {r[1].lower(): r[1] for r in c.fetchall()}
            sc_col = "scenario" if "scenario" in sd_cols else None
            fields_needed = {"kit","module","item","final_qty"}
            stock_rows = []
            if fields_needed.issubset(sd_cols.keys()):
                fields = ["kit","module","item","final_qty"]
                if sc_col: fields.append(sc_col)
                if "exp_date" in sd_cols: fields.append("exp_date")
                if "unique_id" in sd_cols: fields.append("unique_id")
                if "management_type" in sd_cols: fields.append("management_type")
                c.execute(f"SELECT {', '.join(fields)} FROM stock_data")
                idx = {f:i for i,f in enumerate(fields)}
                for r in c.fetchall():
                    self._diag["total_stock_rows"] += 1
                    kit = (r[idx["kit"]] or "").strip() or None
                    module = (r[idx["module"]] or "").strip() or None
                    item = (r[idx["item"]] or "").strip() or None
                    try:
                        fq = int(str(r[idx["final_qty"]]).strip() or 0)
                    except:
                        fq = 0
                    scenario_id = r[idx[sc_col]] if sc_col else None
                    if (not scenario_id) and "unique_id" in idx and r[idx["unique_id"]]:
                        scenario_id = r[idx["unique_id"]].split("/")[0]
                    scenario_id = str(scenario_id).strip() if scenario_id else None
                    exp_raw = r[idx["exp_date"]] if "exp_date" in idx else None
                    mtype = r[idx["management_type"]].lower().strip() if "management_type" in idx and r[idx["management_type"]] else ""
                    stock_rows.append((scenario_id, kit, module, item, fq, exp_raw, mtype))
        finally:
            c.close()
            conn.close()

        today = date.today()
        short_limit = today + timedelta(days=SHORT_EXPIRY_DAYS)
        action_limit = today + timedelta(days=EXPIRY_ACTION_WINDOW_DAYS)
        scenario_exp_dates = defaultdict(list)
        scenario_short_exp_qty = defaultdict(int)

        for sid, kit, module, item, fq, exp_raw, mtype in stock_rows:
            is_on_shelf = (mtype == "on-shelf")
            is_item_row = (item and ((kit and module) or (item and not module) or is_on_shelf))
            if not is_item_row or fq < MIN_STOCK_FOR_ACTION:
                continue
            self._diag["candidate_item_rows"] += 1
            scenario_known = sid in scenario_ids
            if not scenario_known and not ALWAYS_INCLUDE_ORPHAN_STOCK:
                continue

            if scenario_known and kit and module and item:
                std_qty = items_std[sid][kit].get((module,item),0)
                capped = min(fq, std_qty) if std_qty > 0 else 0
                module_std_sum[sid][kit][module] += std_qty
                module_stock_sum[sid][kit][module] += capped
                kit_std_sum[sid][kit] += std_qty
                kit_stock_sum[sid][kit] += capped

            d = parse_flexible_date(exp_raw)
            if d:
                if scenario_known:
                    scenario_exp_dates[sid].append(d)
                    if d <= short_limit:
                        scenario_short_exp_qty[sid] += fq
                if d <= action_limit:
                    days_left = (d - today).days
                    desc = get_item_description(item) or item
                    orphan_prefix = "* " if not scenario_known else ""
                    if not scenario_known:
                        self._diag["orphan_rows_included"] += 1
                    self._expiring_items.append({
                        "scenario_id": sid if scenario_known else None,
                        "code": item,
                        "description": desc,
                        "expiry": d,
                        "days_left": days_left,
                        "qty": fq,
                        "expired": days_left < 0,
                        "orphan": not scenario_known,
                        "prefix": orphan_prefix
                    })
            else:
                if exp_raw not in (None,"","None"):
                    self._diag["bad_date_format"] += 1

        # Completeness
        complete_kits = defaultdict(int)
        incomplete_kits = defaultdict(int)
        total_kits = defaultdict(int)
        for sid in kit_std_sum:
            for kit in kit_std_sum[sid]:
                std_total = kit_std_sum[sid][kit]
                stock_total = kit_stock_sum[sid][kit]
                if std_total > 0:
                    if stock_total >= std_total:
                        complete_kits[sid] += 1
                    else:
                        incomplete_kits[sid] += 1
                else:
                    incomplete_kits[sid] += 1
                total_kits[sid] += 1

        expired = [e for e in self._expiring_items if e["expired"]]
        soon = [e for e in self._expiring_items if not e["expired"]]
        self._diag["expired_count"] = len(expired)
        self._diag["soon_count"] = len(soon)
        self._diag["qualifying_action_rows"] = len(self._expiring_items)
        if DEBUG:
            print("[Dashboard] diagnostics:", self._diag)

        enriched = []
        for s in scenarios:
            sid = s["scenario_id"]
            exp_list = scenario_exp_dates.get(sid, [])
            shortest = min(exp_list).strftime("%Y-%m-%d") if exp_list else "—"
            enriched.append({
                **s,
                "total_kits": total_kits.get(sid,0),
                "complete_kits": complete_kits.get(sid,0),
                "incomplete_kits": incomplete_kits.get(sid,0),
                "short_expiring_qty": scenario_short_exp_qty.get(sid,0),
                "items_missing": "—",
                "shortest_expiry": shortest,
                "last_inventory": "—",
                "last_movement": "—",
                "kits_count": total_kits.get(sid,0),
                "modules_count": sum(len(module_std_sum[sid][k]) for k in module_std_sum[sid]),
                "items_count": sum(len(items_std[sid][k]) for k in items_std[sid]),
                "on_shelf_items": 0
            })
        return enriched

    # ---------- Rendering Scenarios ----------
    def _render_scenarios(self, metrics):
        self._clear_scenarios_ui()
        if not metrics:
            tk.Label(self.inner_frame,
                     text=lang.t("dashboard.no_scenarios", fallback="No scenarios found."),
                     bg=PANEL_BG, fg="#444", font=("Helvetica",12,"italic")).pack(pady=20)
            self.after_idle(self._reset_scroll)
            return
        for scen in metrics:
            self._render_single_scenario(scen)
        self.after_idle(self._reset_scroll)

    def _render_single_scenario(self, scen):
        outer = tk.Frame(self.inner_frame, bg=PANEL_BG, highlightthickness=1,
                         highlightbackground=BORDER_COLOR, pady=2, padx=4)
        outer.pack(fill="x", expand=True, pady=3, padx=4)

        header = tk.Frame(outer, bg=SCENARIO_HEADER_BG)
        header.pack(fill="x", padx=0, pady=(0,2))
        tk.Label(header, text=lang.t("dashboard.scenario_name", fallback="Scenario Name"),
                 bg=SCENARIO_HEADER_BG, font=FONT_LABEL, width=18, anchor="w").grid(row=0,column=0, sticky="w")
        tk.Label(header, text=scen["name"], bg=SCENARIO_HEADER_BG,
                 font=FONT_SCENARIO, anchor="w").grid(row=1,column=0, sticky="w", padx=(0,8))

        meta_frame = tk.Frame(outer, bg=PANEL_BG)
        meta_frame.pack(fill="x", padx=0, pady=(0,2))

        meta_items = [
            (lang.t("dashboard.activity_type", "Activity Type"), scen["activity_type"] or "—"),
            (lang.t("dashboard.target_population", "Target population"), scen["target_population"] or "—"),
            (lang.t("dashboard.stock_location", "Stock Location"), scen["stock_location"] or "—"),
            (lang.t("dashboard.responsible_person", "Responsible person"), scen["responsible_person"] or "—"),
            (lang.t("dashboard.remarks", "Remarks"), scen["remarks"] or "—"),
        ]

        def adjust_wrap(event, label_widgets):
            width_px = event.width / max(1, len(label_widgets))
            for lw in label_widgets:
                lw.config(wraplength=int(width_px - 14))

        value_labels = []
        for idx,(lbl,val) in enumerate(meta_items):
            bloc = tk.Frame(meta_frame, bg=PANEL_BG)
            bloc.grid(row=0, column=idx, sticky="nsew", padx=4)
            tk.Label(bloc, text=lbl, font=FONT_LABEL, bg=PANEL_BG, anchor="w").pack(anchor="w")
            vlabel = tk.Label(bloc, text=val, font=FONT_VALUE, bg=PANEL_BG,
                              anchor="w", justify="left", wraplength=140)
            vlabel.pack(anchor="w", fill="x")
            value_labels.append(vlabel)
        meta_frame.bind("<Configure>", lambda e: adjust_wrap(e, value_labels))

        metrics_frame = tk.Frame(outer, bg=PANEL_BG)
        metrics_frame.pack(fill="x", padx=0, pady=(2,2))
        pairs = [
            (lang.t("dashboard.total_number_kits","Total Number of Kits"), scen["total_kits"]),
            (lang.t("dashboard.complete","Complete"), scen["complete_kits"]),
            (lang.t("dashboard.incomplete","In-Complete"), scen["incomplete_kits"]),
            (lang.t("dashboard.short_expiring","Number of items with Short-Expiry"), scen["short_expiring_qty"]),
            (lang.t("dashboard.items_missing","Number of items Missing"), scen["items_missing"]),
            (lang.t("dashboard.shortest_expiry","Shortest Expiry date"), scen["shortest_expiry"]),
            (lang.t("dashboard.last_inventory","Date of Last Inventory"), scen["last_inventory"]),
            (lang.t("dashboard.last_movement","Date of last Stock Movement"), scen["last_movement"]),
        ]
        for idx,(lbl,val) in enumerate(pairs):
            cell = tk.Frame(metrics_frame, bg=PANEL_BG, highlightthickness=1,
                            highlightbackground=BORDER_COLOR, padx=4, pady=1)
            cell.grid(row=0, column=idx, sticky="nsew", padx=2, pady=2)
            tk.Label(cell, text=lbl, font=("Helvetica",9,"bold"),
                     bg=PANEL_BG, anchor="w", wraplength=130).pack(fill="x")
            tk.Label(cell, text=val, font=("Helvetica",10),
                     bg=PANEL_BG, anchor="w").pack(fill="x")

        counts_frame = tk.Frame(outer, bg=PANEL_BG)
        counts_frame.pack(fill="x", padx=0, pady=(2,0))
        counts_txt = (
            f"{lang.t('dashboard.kits','Kits')}: {scen['kits_count']}    "
            f"{lang.t('dashboard.modules','Modules')}: {scen['modules_count']}    "
            f"{lang.t('dashboard.items','Items')}: {scen['items_count']}"
        )
        tk.Label(counts_frame, text=counts_txt, bg=PANEL_BG,
                 font=("Helvetica",10,"italic"), anchor="w").pack(fill="x")

    # ---------- Project Panel Update ----------
    def _update_project_panel(self, info):
        self.project_details_text.config(state="normal")
        self.project_details_text.delete("1.0","end")
        lines = [
            f"{lang.t('dashboard.project_name','Project Name')}: {info['project_name']}",
            f"{lang.t('dashboard.project_code','Project Code')}: {info['project_code'] or '—'}",
            f"{lang.t('dashboard.project_updated','Project Updated')}: {info['updated_at'] or '—'}",
            f"{lang.t('dashboard.total_scenarios','Total Scenarios')}: {info['scenario_count']}",
            f"{lang.t('dashboard.total_kits','Total Kits')}: {info['total_kits']}",
            f"{lang.t('dashboard.total_modules','Total Modules')}: {info['total_modules']}",
            f"{lang.t('dashboard.total_items','Total Items')}: {info['total_items']}",
            f"{lang.t('dashboard.standard_last_update','Standard List Last Update')}: {info['standard_last_update'] or '—'}"
        ]
        self.project_details_text.insert("1.0", "\n".join(lines))
        self.project_details_text.config(state="disabled")

    # ---------- Actions ----------
    def _compute_actions(self, metrics, project_info):
        actions = []
        expired = [r for r in self._expiring_items if r["expired"]]
        soon = [r for r in self._expiring_items if not r["expired"]]
        expired.sort(key=lambda r: r["days_left"])
        soon.sort(key=lambda r: r["days_left"])

        for r in expired:
            actions.append(("expired",
                f"{r['prefix']}{r['code']} {r['description']} "
                f"{lang.t('dashboard.action.item_expired', fallback='has expired')} "
                f"({lang.t('dashboard.action.expired_days_ago', fallback='expired {d} days ago').format(d=abs(r['days_left']))}). "
                f"{lang.t('dashboard.action.remove_stock', fallback='Please remove from stock.')}"
            ))

        for r in soon:
            actions.append(("soon",
                f"{r['prefix']}{r['code']} {r['description']} "
                f"{lang.t('dashboard.action.item_expiring', fallback='is expiring in')} {r['days_left']} "
                f"{lang.t('dashboard.action.days', fallback='days')} "
                f"({lang.t('dashboard.action.expiry', fallback='Expiry')}: {r['expiry'].strftime('%Y-%m-%d')}). "
                f"{lang.t('dashboard.action.take_action', fallback='Please take appropriate action.')}"
            ))

        if not actions and SHOW_SCENARIO_EXPIRY_LINES:
            for scen in metrics:
                qty = scen["short_expiring_qty"]
                if qty:
                    actions.append(("normal",
                        lang.t("dashboard.action.expiring",
                               fallback="In scenario {name}, {n} item(s) expiring in {d} days.")
                        .format(name=scen["name"], n=qty, d=SHORT_EXPIRY_DAYS)
                    ))

        if not actions:
            today = date.today()
            std_last = project_info.get("standard_last_update")
            d = parse_flexible_date(std_last) if std_last else None
            if d and (today - d).days > STANDARD_LIST_STALE_DAYS:
                actions.append(("normal", lang.t("dashboard.action.standard_stale",
                                                 fallback="Standard list has not been revised for over 2 years.")))
            for scen in metrics:
                upd = scen.get("scenario_updated_at")
                ds = parse_flexible_date(upd) if upd else None
                if ds and (today - ds).days > SCENARIO_STALE_DAYS:
                    actions.append(("normal",
                        lang.t("dashboard.action.scenario_stale",
                               fallback="Scenario '{name}' not revised for over 3 years.")
                        .format(name=scen["name"])
                    ))
        if not actions:
            actions.append(("normal", lang.t("dashboard.action.none", fallback="No immediate actions required.")))
        return actions

    def _render_actions(self, actions):
        # Common spacing logic: append extra blank lines
        if USE_TEXT_ACTIONS and self.actions_text:
            self.actions_text.config(state="normal")
            self.actions_text.delete("1.0","end")
            last_index = len(actions) - 1
            for i,(tag,line) in enumerate(actions):
                self.actions_text.insert("end", line + "\n", (tag,))
                extra = ACTION_EXTRA_BLANK_LINES
                if i != last_index:
                    self.actions_text.insert("end", ("\n" * extra), ("spacing",))
            self.actions_text.config(state="disabled")
        else:
            self.actions_list.delete(0,"end")
            for i,(tag,line) in enumerate(actions):
                self.actions_list.insert("end", line)
                if ACTION_EXTRA_BLANK_LINES and i != len(actions)-1:
                    self.actions_list.insert("end", "")  # blank line

        if SHOW_DEBUG_PANEL and self.debug_panel:
            diag_text = lang.t(
                "dashboard.debug_panel",
                fallback=(
                    "Stock rows: {total_stock_rows} | "
                    "Item candidates: {candidate_item_rows} | "
                    "Action rows: {qualifying_action_rows} "
                    "(expired {expired_count}, soon {soon_count}) | "
                    "Bad date fmt: {bad_date_format} | "
                    "Orphans incl: {orphan_rows_included}"
                ),
                total_stock_rows=self._diag.get('total_stock_rows',0),
                candidate_item_rows=self._diag.get('candidate_item_rows',0),
                qualifying_action_rows=self._diag.get('qualifying_action_rows',0),
                expired_count=self._diag.get('expired_count',0),
                soon_count=self._diag.get('soon_count',0),
                bad_date_format=self._diag.get('bad_date_format',0),
                orphan_rows_included=self._diag.get('orphan_rows_included',0),
            )
            self.debug_panel.config(text=diag_text)


__all__ = ["Dashboard"]