"""app/database/migrations.py — Schema bootstrap.

Call init_db() once at startup to create / migrate all tables.
"""
import sqlite3
import logging
from .connection import get_conn
from .custom_fields import ensure_user_custom_fields_tables
from .help_content import ensure_help_table_and_seed
from .users import ensure_default_admin
from ..config import (
    DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_EMAIL, DEFAULT_ADMIN_PASSWORD,
    utc_now_str,
)

def ensure_timesheets_has_draft_column():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(timesheets)")
    cols = [r[1] for r in cur.fetchall()]
    if 'has_draft' not in cols:
        cur.execute("ALTER TABLE timesheets ADD COLUMN has_draft INTEGER DEFAULT 0")
    conn.commit()
    conn.close()


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            email TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            status TEXT NOT NULL,
            screen_access TEXT NOT NULL DEFAULT 'BOTH'
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            stored_filename TEXT NOT NULL,
            original_filename TEXT NOT NULL,
            uploaded_at TEXT NOT NULL
            -- module column added by ensure_uploads_module_column()
            -- uploaded_by column added by ensure_uploads_uploaded_by_column()
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS email_reminders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            recipient TEXT NOT NULL,
            subject TEXT NOT NULL,
            message TEXT NOT NULL,
            scheduled_date TEXT NOT NULL,
            scheduled_time TEXT NOT NULL DEFAULT '09:00',
            sent INTEGER NOT NULL DEFAULT 0,
            sent_at TEXT,
            last_error TEXT,
            created_by TEXT,
            created_at TEXT NOT NULL
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS timesheets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id INTEGER NOT NULL UNIQUE,
            week_start TEXT NOT NULL,
            mon REAL NOT NULL DEFAULT 0,
            tue REAL NOT NULL DEFAULT 0,
            wed REAL NOT NULL DEFAULT 0,
            thu REAL NOT NULL DEFAULT 0,
            fri REAL NOT NULL DEFAULT 0,
            sat REAL NOT NULL DEFAULT 0,
            sun REAL NOT NULL DEFAULT 0,
            total REAL NOT NULL DEFAULT 0,
            submitted INTEGER NOT NULL DEFAULT 0,
            submitted_at TEXT,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(upload_id) REFERENCES uploads(id) ON DELETE CASCADE
        )
    """)

    conn.commit()
    conn.close()

    ensure_uploads_module_column()
    ensure_uploads_uploaded_by_column()
    ensure_timesheets_submit_columns()
    ensure_email_reminders_columns()
    ensure_users_screen_access_column()
    ensure_timesheets_remaining_planned_column()
    ensure_timesheets_has_draft_column()
    ensure_users_feature_flags()
    ensure_user_custom_fields_tables()
    ensure_admin_default_access()
    ensure_employee_help_locked_off()
    ensure_help_table_and_seed()
    ensure_default_admin()


def ensure_uploads_module_column():
    """Add 'module' column to uploads if missing. Defaults to 'Generic'."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(uploads)")
    cols = [r[1] for r in cur.fetchall()]
    if "module" not in cols:
        logging.info("Adding 'module' column to uploads table...")
        cur.execute("ALTER TABLE uploads ADD COLUMN module TEXT NOT NULL DEFAULT 'Generic'")
        conn.commit()
    conn.close()


def ensure_uploads_uploaded_by_column():
    """Add 'uploaded_by' column to uploads if missing."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(uploads)")
    cols = [r[1] for r in cur.fetchall()]
    if "uploaded_by" not in cols:
        logging.info("Adding 'uploaded_by' column to uploads table...")
        cur.execute("ALTER TABLE uploads ADD COLUMN uploaded_by TEXT")
        conn.commit()
    conn.close()


def ensure_timesheets_submit_columns():
    """Add submit/lock columns to timesheets if missing (migration-safe)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(timesheets)")
    cols = [r[1] for r in cur.fetchall()]

    if "submitted" not in cols:
        logging.info("Adding 'submitted' column to timesheets table...")
        cur.execute("ALTER TABLE timesheets ADD COLUMN submitted INTEGER NOT NULL DEFAULT 0")

    if "submitted_at" not in cols:
        logging.info("Adding 'submitted_at' column to timesheets table...")
        cur.execute("ALTER TABLE timesheets ADD COLUMN submitted_at TEXT")

    conn.commit()
    conn.close()


def ensure_timesheets_remaining_planned_column():
    """Add 'remaining_planned' column to timesheets if missing."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(timesheets)")
    cols = [r[1] for r in cur.fetchall()]
    if "remaining_planned" not in cols:
        logging.info("Adding 'remaining_planned' column to timesheets table...")
        cur.execute("ALTER TABLE timesheets ADD COLUMN remaining_planned REAL")
        conn.commit()
    conn.close()




def ensure_email_reminders_columns():
    """Add scheduled_time/sent tracking columns to email_reminders if missing (migration-safe)."""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(email_reminders)")
        cols = [r[1] for r in cur.fetchall()]
        changed = False
        if 'scheduled_time' not in cols:
            cur.execute("ALTER TABLE email_reminders ADD COLUMN scheduled_time TEXT NOT NULL DEFAULT '09:00'")
            changed = True
        if 'sent' not in cols:
            cur.execute("ALTER TABLE email_reminders ADD COLUMN sent INTEGER NOT NULL DEFAULT 0")
            changed = True
        if 'sent_at' not in cols:
            cur.execute("ALTER TABLE email_reminders ADD COLUMN sent_at TEXT")
            changed = True
        if 'last_error' not in cols:
            cur.execute("ALTER TABLE email_reminders ADD COLUMN last_error TEXT")
            changed = True
        if changed:
            conn.commit()
    finally:
        conn.close()


def ensure_users_screen_access_column():
    """Add 'screen_access' column to users if missing. Defaults to 'BOTH'."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(users)")
    cols = [r[1] for r in cur.fetchall()]
    if "screen_access" not in cols:
        logging.info("Adding 'screen_access' column to users table...")
        cur.execute("ALTER TABLE users ADD COLUMN screen_access TEXT NOT NULL DEFAULT 'BOTH'")
        conn.commit()
    conn.close()



def ensure_users_feature_flags():
    """Add per-user feature permission flags if missing.

    Flags:
      - can_email  : access to Email Settings
      - can_export : access to Export Data
      - can_help   : access to Help

    Defaults:
      - can_help = 1 (enabled)
      - can_email/can_export = 0 (disabled)
    """
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(users)")
    cols = [r[1] for r in cur.fetchall()]

    changed = False
    if "can_email" not in cols:
        cur.execute("ALTER TABLE users ADD COLUMN can_email INTEGER NOT NULL DEFAULT 0")
        changed = True
    if "can_export" not in cols:
        cur.execute("ALTER TABLE users ADD COLUMN can_export INTEGER NOT NULL DEFAULT 0")
        changed = True
    if "can_help" not in cols:
        cur.execute("ALTER TABLE users ADD COLUMN can_help INTEGER NOT NULL DEFAULT 1")
        changed = True
    if changed:
        conn.commit()
    conn.close()



def ensure_admin_default_access():
    """Ensure Admin users always have Email/Export/Help enabled and those flags stay locked."""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(users)")
        cols = [r[1] for r in cur.fetchall()]
        sets = []
        if "screen_access" in cols:
            sets.append("screen_access='BOTH'")
        if "can_email" in cols:
            sets.append("can_email=1")
        if "can_export" in cols:
            sets.append("can_export=1")
        if "can_help" in cols:
            sets.append("can_help=1")
        if sets:
            cur.execute("UPDATE users SET " + ", ".join(sets) + " WHERE role='Admin'")
            conn.commit()
    finally:
        conn.close()


def ensure_employee_help_locked_off():
    """Ensure Employee users always have Help disabled (can_help=0)."""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(users)")
        cols = [r[1] for r in cur.fetchall()]
        if "can_help" in cols:
            cur.execute("UPDATE users SET can_help=0 WHERE role='Employee'")
            conn.commit()
    finally:
        conn.close()


# -------------------------
# Custom User Fields (Admin-defined)
# -------------------------

