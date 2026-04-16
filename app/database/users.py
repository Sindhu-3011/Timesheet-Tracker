"""app/database/users.py — User CRUD and ensure_default_admin."""
import sqlite3
import hashlib
import logging
from .connection import get_conn
from ..config import (
    DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_EMAIL, DEFAULT_ADMIN_PASSWORD,
    utc_now_str, normalize_screen_access,
)

def create_user_db_return_id(username, email, password, role, status, screen_access="BOTH", can_email=0, can_export=0, can_help=1):
    """Same as create_user_db, but returns new user id on success (or None)."""
    username = (username or "").strip()
    email = (email or "").strip().lower()
    role = (role or "Employee").strip()
    status = (status or "Active").strip()
    if not username or not email:
        return None
    if username_exists(username) or email_exists(email):
        return None
    screen_access = normalize_screen_access(screen_access)

    def _to01(v, default=0):
        try:
            return 1 if str(v).strip().lower() in ("1", "true", "on", "yes") else 0
        except Exception:
            return default

    can_email = _to01(can_email, 0)
    can_export = _to01(can_export, 0)
    can_help = _to01(can_help, 1)

    if role == "Employee":
        can_help = 0
    if role == "Admin":
        screen_access = "BOTH"
        can_email, can_export, can_help = 1, 1, 1

    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO users (username, email, password, role, status, screen_access, can_email, can_export, can_help) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (username, email, password or "", role, status, screen_access, int(can_email), int(can_export), int(can_help)),
        )
        conn.commit()
        new_id = cur.lastrowid
        return int(new_id) if new_id else None
    except sqlite3.IntegrityError:
        return None
    finally:
        try:
            conn.close()
        except Exception:
            pass
def ensure_default_admin():
    """Seed the first Admin account when the DB is empty (first run only).

    Credentials are configured via DEFAULT_ADMIN_* variables in .env.
    """
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM users")
    count = cur.fetchone()[0]
    if count == 0:
        cur.execute(
            "INSERT INTO users (username, email, password, role, status, screen_access, can_email, can_export, can_help) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_EMAIL, DEFAULT_ADMIN_PASSWORD, "Admin", "Active", "BOTH", 1, 1, 1),
        )
        conn.commit()
        logging.info("Default admin created: %s", DEFAULT_ADMIN_EMAIL)
    conn.close()

# -------------------------
# Help Content (DB-backed, Admin editable)
# -------------------------

