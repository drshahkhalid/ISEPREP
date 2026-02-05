import sqlite3
from datetime import datetime

DB_PATH = "your_database.sqlite"  # change to your actual DB file path

REFRESH_STD_LIST_COMBINED_SQL = """
DELETE FROM std_list_combined;
INSERT INTO std_list_combined (code, description, type, std_qty_collective)
SELECT
    all_codes.code,
    il.designation,
    il.type,
    SUM(all_codes.qty_component) AS std_qty_collective
FROM (
    SELECT code, quantity AS qty_component FROM compositions
    UNION ALL
    SELECT code, std_qty AS qty_component FROM kit_items
) AS all_codes
LEFT JOIN items_list il ON il.code = all_codes.code
GROUP BY all_codes.code;
"""

REFRESH_STD_QTY_HELPER_SQL = """
DELETE FROM std_qty_helper;
INSERT INTO std_qty_helper (id, code, description, type, scenario_id, scenario, kit, module, std_qty)
SELECT
  c.std_id,
  c.code,
  il.designation,
  il.type,
  c.scenario_id,
  s.name,
  NULL,
  NULL,
  c.quantity
FROM compositions c
LEFT JOIN items_list il ON il.code = c.code
LEFT JOIN scenarios s ON s.scenario_id = c.scenario_id
UNION ALL
SELECT
  CAST(k.id AS TEXT),
  k.code,
  il.designation,
  il.type,
  k.scenario_id,
  s.name,
  k.kit,
  k.module,
  k.std_qty
FROM kit_items k
LEFT JOIN items_list il ON il.code = k.code
LEFT JOIN scenarios s ON s.scenario_id = k.scenario_id;
"""

def refresh_snapshots(db_path: str = DB_PATH) -> dict:
    """
    Rebuilds snapshot tables. Returns summary dict.
    """
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    summary = {"std_list_combined_rows": 0, "std_qty_helper_rows": 0, "timestamp": datetime.now().isoformat(timespec="seconds")}
    try:
        # std_list_combined
        cur.executescript(REFRESH_STD_LIST_COMBINED_SQL)
        cur.execute("SELECT COUNT(*) FROM std_list_combined")
        summary["std_list_combined_rows"] = cur.fetchone()[0]

        # std_qty_helper
        cur.executescript(REFRESH_STD_QTY_HELPER_SQL)
        cur.execute("SELECT COUNT(*) FROM std_qty_helper")
        summary["std_qty_helper_rows"] = cur.fetchone()[0]

        conn.commit()
        return summary
    except Exception as e:
        conn.rollback()
        raise RuntimeError(f"Snapshot refresh failed: {e}") from e
    finally:
        cur.close()
        conn.close()

if __name__ == "__main__":
    info = refresh_snapshots()
    print("Refreshed:", info)