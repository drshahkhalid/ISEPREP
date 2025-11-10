from db import connect_db
from datetime import datetime
import logging
import sqlite3

# Configure logging (you can adjust level to INFO while testing)
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

# Cache for columns so we don't PRAGMA on every insert
_STOCK_TX_COLUMNS_CACHE = None

def _get_stock_transaction_columns(conn) -> set:
    """
    Return a lowercase set of column names for stock_transactions.
    Uses a module-level cache to avoid repeated PRAGMA calls.
    """
    global _STOCK_TX_COLUMNS_CACHE
    if _STOCK_TX_COLUMNS_CACHE is not None:
        return _STOCK_TX_COLUMNS_CACHE
    try:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(stock_transactions)")
        _STOCK_TX_COLUMNS_CACHE = {row[1].lower() for row in cur.fetchall()}
        cur.close()
    except Exception as e:
        logging.error(f"[log_transaction] Failed to inspect table schema: {e}")
        _STOCK_TX_COLUMNS_CACHE = set()
    return _STOCK_TX_COLUMNS_CACHE

def log_transaction(
    *,
    unique_id=None,
    code=None,
    Description=None,
    Expiry_date=None,
    Batch_Number=None,
    Scenario=None,
    Kit=None,
    Module=None,
    Qty_IN=None,
    IN_Type=None,
    Qty_Out=None,
    Out_Type=None,
    Third_Party=None,
    End_User=None,
    Discrepancy=None,
    Remarks=None,
    Movement_Type=None,
    document_number=None,        # NEW OPTIONAL PARAM
    conn=None
):
    """
    Insert a new transaction row into stock_transactions.
    - Auto-generates Date and Time.
    - Reuses provided connection if supplied; otherwise creates & closes one.
    - Automatically includes document_number if the column exists; otherwise logs a warning once.
    - Backward compatible: callers not passing document_number are unaffected.

    Expected columns in table (superset):
      Date, Time, unique_id, code, Description, Expiry_date, Batch_Number,
      Scenario, Kit, Module, Qty_IN, IN_Type, Qty_Out, Out_Type,
      Third_Party, End_User, Discrepancy, Remarks, Movement_Type, document_number (optional)
    """
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")

    created_locally = False
    if conn is None:
        conn = connect_db()
        if conn is None:
            raise ValueError("Database connection failed in log_transaction")
        created_locally = True

    try:
        cols = _get_stock_transaction_columns(conn)
        has_doc_col = "document_number" in cols

        base_fields = [
            "Date", "Time", "unique_id", "code", "Description", "Expiry_date",
            "Batch_Number", "Scenario", "Kit", "Module",
            "Qty_IN", "IN_Type", "Qty_Out", "Out_Type",
            "Third_Party", "End_User", "Discrepancy", "Remarks", "Movement_Type"
        ]
        base_values = [
            date_str, time_str, unique_id, code, Description, Expiry_date,
            Batch_Number, Scenario, Kit, Module,
            Qty_IN, IN_Type, Qty_Out, Out_Type,
            Third_Party, End_User, Discrepancy, Remarks, Movement_Type
        ]

        if has_doc_col:
            field_list = base_fields + ["document_number"]
            values = base_values + [document_number]
        else:
            field_list = base_fields
            values = base_values
            if document_number is not None:
                logging.warning(
                    "[log_transaction] document_number provided but column missing. "
                    "Run: ALTER TABLE stock_transactions ADD COLUMN document_number TEXT;"
                )

        placeholders = ", ".join(["?"] * len(field_list))
        cols_sql = ", ".join(field_list)

        sql = f"INSERT INTO stock_transactions ({cols_sql}) VALUES ({placeholders})"

        cur = conn.cursor()
        cur.execute(sql, values)
        conn.commit()
        logging.info(f"Logged transaction unique_id={unique_id} doc={document_number}")
        cur.close()

    except sqlite3.Error as e:
        try:
            conn.rollback()
        except Exception:
            pass
        logging.error(f"Error logging transaction for unique_id {unique_id}: {e}")
        raise
    finally:
        if created_locally:
            try:
                conn.close()
            except:
                pass

def format_decimal(value):
    try:
        return float(value) if value not in (None, "", " ") else None
    except ValueError:
        return None

def safe_upper(value):
    return value.upper() if value else None
