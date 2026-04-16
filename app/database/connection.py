"""app/database/connection.py — SQLite connection factory."""
import sqlite3
from ..config import DB_FILE

def get_conn():
    conn = sqlite3.connect(DB_FILE)
    try:
        conn.execute("PRAGMA foreign_keys = ON;")
    except Exception:
        pass
    return conn


# -------------------------
# Database helpers
# -------------------------
