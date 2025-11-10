import sqlite3
from sqlite3 import Error

# SQLite database file (in the same folder as your app)
DB_FILE = 'iseprep.db'

def connect_db():
    """
    Central function to connect to the SQLite database.
    Creates the file if it doesn't exist.
    """
    try:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row  # This makes it return dict-like rows
        return conn
    except Error as e:
        raise
    return None