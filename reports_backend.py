"""
reports_backend.py
Backend reporting/data-access utilities for the Reports UI.

Creates SQL views required by the GUI (reports.py) and provides helper
functions to query them if you later want to bypass direct SQL in the UI.

Views created:
    vw_report_detailed
    vw_report_summary
    vw_report_expiry
    vw_report_required_qty

Call initialize_reporting() once at application startup (or before opening Reports window).
"""

from db import connect_db
import sqlite3

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _execute_ddl(cursor, stmts):
    for s in stmts:
        s = s.strip()
        if not s:
            continue
        try:
            cursor.execute(s)
        except sqlite3.Error as e:
            print(f"[reports_backend] DDL warning: {e} | {s[:70]}...")

# ---------------------------------------------------------------------------
# View creation
# ---------------------------------------------------------------------------

def ensure_report_views():
    """
    Drop + recreate all reporting views.
    Safe to call repeatedly (idempotent enough).
    """
    conn = connect_db()
    cur = conn.cursor()
    try:
        drops = [
            "DROP VIEW IF EXISTS vw_report_detailed",
            "DROP VIEW IF EXISTS vw_report_summary",
            "DROP VIEW IF EXISTS vw_report_expiry",
            "DROP VIEW IF EXISTS vw_report_required_qty"
        ]
        _execute_ddl(cur, drops)

        vw_report_detailed = """
        CREATE VIEW vw_report_detailed AS
        SELECT
            ROW_NUMBER() OVER (ORDER BY Date, Time, code) AS row_id,
            Date || ' ' || COALESCE(Time,'00:00:00') AS timestamp,
            Date AS transaction_date,
            Time AS transaction_time,
            code,
            Description AS description,
            Scenario AS scenario,
            Kit AS kit,
            Module AS module,
            COALESCE(Third_Party,'') AS third_party,
            COALESCE(End_User,'') AS end_user,
            IN_Type,
            Out_Type,
            Movement_Type AS movement_type,
            Qty_IN,
            Qty_Out,
            Expiry_date,
            unique_id,
            Remarks AS remarks,
            CASE
                WHEN IFNULL(Qty_IN,0) > 0 THEN 'IN'
                WHEN IFNULL(Qty_Out,0) > 0 THEN 'OUT'
                ELSE 'NA'
            END AS direction,
            CASE
                WHEN IFNULL(Qty_IN,0) > 0 THEN Qty_IN
                WHEN IFNULL(Qty_Out,0) > 0 THEN Qty_Out
                ELSE 0
            END AS quantity
        FROM stock_transactions
        """

        vw_report_summary = """
        CREATE VIEW vw_report_summary AS
        SELECT
            Date AS transaction_date,
            code,
            MAX(Description) AS description,
            Scenario AS scenario,
            SUM(COALESCE(Qty_IN,0)) AS total_in,
            SUM(COALESCE(Qty_Out,0)) AS total_out,
            (SUM(COALESCE(Qty_IN,0)) - SUM(COALESCE(Qty_Out,0))) AS net_movement,
            CASE
                WHEN (SUM(COALESCE(Qty_IN,0)) - SUM(COALESCE(Qty_Out,0))) > 0 THEN 'IN_DOMINANT'
                WHEN (SUM(COALESCE(Qty_IN,0)) - SUM(COALESCE(Qty_Out,0))) < 0 THEN 'OUT_DOMINANT'
                ELSE 'BALANCED'
            END AS movement_balance_flag
        FROM stock_transactions
        GROUP BY Date, code, Scenario
        """

        vw_report_expiry = """
        CREATE VIEW vw_report_expiry AS
        SELECT
            st.code,
            st.Description AS description,
            st.Scenario AS scenario,
            st.Kit AS kit,
            st.Module AS module,
            st.Expiry_date,
            COALESCE(st.Qty_IN,0) AS qty_in,
            CAST((julianday(st.Expiry_date) - julianday('now')) AS INTEGER) AS days_left,
            CASE
                WHEN st.Expiry_date IS NULL THEN 'NO_DATE'
                WHEN (julianday(st.Expiry_date) - julianday('now')) < 0 THEN 'EXPIRED'
                WHEN (julianday(st.Expiry_date) - julianday('now')) <= 30 THEN 'ALERT_30'
                WHEN (julianday(st.Expiry_date) - julianday('now')) <= 60 THEN 'ALERT_60'
                WHEN (julianday(st.Expiry_date) - julianday('now')) <= 90 THEN 'ALERT_90'
                ELSE 'OK'
            END AS expiry_bucket
        FROM stock_transactions st
        JOIN items_list il ON il.code = st.code
        WHERE st.Qty_IN > 0
          AND st.Expiry_date IS NOT NULL
          AND il.remarks LIKE '%exp%'
        """

        vw_report_required_qty = """
        CREATE VIEW vw_report_required_qty AS
        WITH required AS (
            SELECT
                ki.scenario_id AS scenario,
                ki.kit,
                ki.module,
                ki.item AS code,
                SUM(COALESCE(ki.std_qty,0)) AS required_qty
            FROM kit_items ki
            WHERE (ki.level IN ('tertiary','item') OR ki.level IS NULL)
            GROUP BY ki.scenario_id, ki.kit, ki.module, ki.item
        ),
        current_stock AS (
            SELECT
                sd.code,
                SUBSTR(sd.unique_id, 1, INSTR(sd.unique_id,'/')-1) AS scenario,
                SUM(COALESCE(sd.qty_in,0) - COALESCE(sd.qty_out,0)) AS net_qty
            FROM stock_data sd
            GROUP BY sd.code, scenario
        )
        SELECT
            r.scenario,
            r.kit,
            r.module,
            r.code,
            r.required_qty,
            COALESCE(cs.net_qty,0) AS current_qty,
            CASE
                WHEN r.required_qty > 0
                 THEN ROUND((COALESCE(cs.net_qty,0)/r.required_qty)*100,1)
                ELSE NULL
            END AS coverage_pct
        FROM required r
        LEFT JOIN current_stock cs
               ON cs.scenario = r.scenario
              AND cs.code = r.code
        ORDER BY r.scenario, r.kit, r.module, r.code
        """

        creates = [
            vw_report_detailed,
            vw_report_summary,
            vw_report_expiry,
            vw_report_required_qty
        ]
        _execute_ddl(cur, creates)
        conn.commit()
    finally:
        cur.close()
        conn.close()

# ---------------------------------------------------------------------------
# Optional fetch helpers
# ---------------------------------------------------------------------------

def _fetch_dicts(cursor):
    cols = [c[0] for c in cursor.description]
    return [dict(zip(cols, r)) for r in cursor.fetchall()]

def fetch_view(view_name, where="", params=None, order=""):
    params = params or []
    conn = connect_db()
    cur = conn.cursor()
    try:
        sql = f"SELECT * FROM {view_name}"
        if where:
            sql += f" WHERE {where}"
        if order:
            sql += f" ORDER BY {order}"
        cur.execute(sql, params)
        return _fetch_dicts(cur)
    finally:
        cur.close()
        conn.close()

def initialize_reporting():
    ensure_report_views()

if __name__ == "__main__":
    print("Creating views...")
    initialize_reporting()
    print("Done.")