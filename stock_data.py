import sqlite3
from datetime import datetime, timedelta
from language_manager import lang
from db import connect_db
from dateutil import parser
from calendar import monthrange
from popup_utils import custom_popup, custom_askyesno, custom_dialog
import logging

# Configure logging (only errors)
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def parse_expiry(text):
    """Parse date and return a datetime.date object."""
    if not text or text.lower() in ('none', ''):
        return None
    try:
        dt = parser.parse(text, dayfirst=True, fuzzy=True)
        if '/' in text or '-' in text:
            parts = text.replace('-', '/').split('/')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                month = int(parts[0])
                year = int(parts[1])
                if year < 100:
                    year += 2000 if year < 50 else 1900
                last_day = monthrange(year, month)[1]
                dt = datetime(year, month, last_day)
        return dt.date()
    except Exception as e:
        logging.error(f"Error parsing expiry date {text}: {str(e)}")
        return None

class StockData:
    @staticmethod
    def parse_unique_id(unique_id):
        """
        Unique ID structure: scenario/kit/module/item/std_qty/exp_date
        Returns dict with all 6 layers. Expiry normalized to datetime.date.
        """
        parts = unique_id.split("/")
        raw_exp_date = parts[5] if len(parts) > 5 else None
        formatted_exp_date = parse_expiry(raw_exp_date) if raw_exp_date and raw_exp_date.lower() != 'none' else None

        return {
            "scenario": parts[0] if len(parts) > 0 else None,
            "kit": parts[1] if len(parts) > 1 and parts[1] != "None" else None,
            "module": parts[2] if len(parts) > 2 and parts[2] != "None" else None,
            "item": parts[3] if len(parts) > 3 else None,
            "std_qty": int(parts[4]) if len(parts) > 4 else 0,
            "exp_date": formatted_exp_date
        }

    @staticmethod
    def add_or_update(unique_id, qty_in=0, qty_out=0, exp_date=None):
        """
        Insert or update stock_data table.
        If unique_id exists, update qty_in, qty_out, exp_date, and updated_at.
        If exp_date is provided, use it; otherwise, parse from unique_id.
        """
        parsed = StockData.parse_unique_id(unique_id)
        effective_exp_date = exp_date if exp_date else parsed['exp_date']
        current_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = connect_db()
        cursor = conn.cursor()

        try:
            cursor.execute("SELECT qty_in, qty_out FROM stock_data WHERE unique_id=?", (unique_id,))
            row = cursor.fetchone()

            if row:
                new_qty_in = (row[0] or 0) + qty_in
                new_qty_out = (row[1] or 0) + qty_out
                cursor.execute("""
                    UPDATE stock_data
                    SET qty_in=?, qty_out=?, exp_date=?, updated_at=?
                    WHERE unique_id=?
                """, (new_qty_in, new_qty_out, effective_exp_date, current_timestamp, unique_id))
            else:
                cursor.execute("""
                    INSERT INTO stock_data
                    (unique_id, scenario, kit, module, item, std_qty, qty_in, qty_out, exp_date, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    unique_id,
                    parsed['scenario'], parsed['kit'], parsed['module'], parsed['item'],
                    parsed['std_qty'], qty_in, qty_out, effective_exp_date, current_timestamp
                ))

            conn.commit()
        except Exception as e:
            conn.rollback()
            logging.error(f"Error in add_or_update for unique_id {unique_id}: {str(e)}")
            raise
        finally:
            cursor.close()
            conn.close()

        StockData.recalculate_for_item(parsed['item'])

    @staticmethod
    def recalculate_for_item(item_code):
        """
        Recalculate qty_to_order, qty_overstock, qty_to_order_per_scenario, qt_expiring
        for all rows with this item_code.
        """
        conn = connect_db()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT COALESCE(SUM(final_qty), 0) AS total_final_qty, COALESCE(SUM(std_qty), 0) AS total_std_qty
            FROM stock_data WHERE item=?
        """, (item_code,))
        totals = cursor.fetchone()
        total_final = totals[0] or 0
        total_std = totals[1] or 0

        cursor.execute("SELECT lead_time_months, cover_period_months FROM project_details LIMIT 1")
        project = cursor.fetchone() or (0, 0)
        days_window = (project[0] + project[1]) * 30

        cursor.execute("SELECT unique_id, std_qty, scenario, exp_date FROM stock_data WHERE item=?", (item_code,))
        rows = cursor.fetchall()

        for row in rows:
            unique_id, std_qty, scenario, exp_date = row
            qty_to_order = max(std_qty - total_final, 0)
            qty_overstock = max(total_final - std_qty, 0)

            cursor.execute("""
                SELECT COALESCE(SUM(final_qty), 0) FROM stock_data
                WHERE item=? AND scenario=?
            """, (item_code, scenario))
            scenario_final = cursor.fetchone()[0] or 0
            qty_to_order_per_scenario = max(std_qty - scenario_final, 0)

            qt_expiring = 0
            if exp_date:
                try:
                    exp_date = datetime.strptime(str(exp_date), "%Y-%m-%d").date()
                    if exp_date <= (datetime.today().date() + timedelta(days=days_window)):
                        qt_expiring = std_qty
                except Exception:
                    qt_expiring = 0

            cursor.execute("""
                UPDATE stock_data
                SET qty_to_order=?,
                    qty_overstock=?,
                    qty_to_order_per_scenario=?,
                    qt_expiring=?
                WHERE unique_id=?
            """, (qty_to_order, qty_overstock, qty_to_order_per_scenario, qt_expiring, unique_id))

        conn.commit()
        cursor.close()
        conn.close()

    @staticmethod
    def cleanup_zero_final_qty():
        """
        Delete rows where final_qty = 0 (no stock remaining).
        """
        conn = connect_db()
        cursor = conn.cursor()

        cursor.execute("DELETE FROM stock_data WHERE final_qty = 0")
        conn.commit()

        cursor.close()
        conn.close()