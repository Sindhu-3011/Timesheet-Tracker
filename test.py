
# ===================== EARLY DB MIGRATION (AUTO-GENERATED) =====================
# Ensures feature-flag columns exist even for old databases.
# Safe to run multiple times.
import sqlite3 as _sqlite3

def _early_migrate_users_feature_flags(db_file='users.db'):
    try:
        conn = _sqlite3.connect(db_file)
        cur = conn.cursor()
        cur.execute("PRAGMA table_info(users)")
        cols = [r[1] for r in cur.fetchall()]
        changed = False
        if 'can_email' not in cols:
            cur.execute("ALTER TABLE users ADD COLUMN can_email INTEGER NOT NULL DEFAULT 0")
            changed = True
        if 'can_export' not in cols:
            cur.execute("ALTER TABLE users ADD COLUMN can_export INTEGER NOT NULL DEFAULT 0")
            changed = True
        if 'can_help' not in cols:
            cur.execute("ALTER TABLE users ADD COLUMN can_help INTEGER NOT NULL DEFAULT 1")
            changed = True
        if changed:
            conn.commit()
    except Exception:
        pass
    finally:
        try:
            conn.close()
        except Exception:
            pass

_early_migrate_users_feature_flags()
# =================== END EARLY DB MIGRATION ===================

#!/usr/bin/env python3
"""
# server_ppm_ntt_only.py
VERSION = "1.0.12-FIX-OCR-DEDUPE"

Single-file HTTP server with:
- Login by email + session cookie
- User management (Admin only)
- Screen Capture module with separate PPM and NTT pages ONLY
- Saved captures show thumbnails; clicking opens same-page modal preview
- Timesheet per saved screenprint (upload):
  * Week-wise (Sun–Sat) hours + week_start (Sunday)
  * Total computed in UI, recalculated/validated server-side
  * Submit per timesheet
  * Save Hours is mandatory for Submit (Submit requires total > 0)
  * Lock prevents edits on server and disables UI inputs
  * Submit-only symmetric enforcement:
      - For same user + same week_start, PPM and NTT hours must match
      - Enforced at Submit (not at Save)
      - Submit fails if any other entered/submitted hours in that week mismatch
      - On successful submit, hours are applied to ALL uploads for that user+week (PPM+NTT) and locked
- Email reminder settings (Admin/Manager)
- Export placeholder (Admin/Manager)
- Upload delete (Admin only)

UPDATED per requirements:
- Employee access restricted to Screen Print modules only (no delete)
- Admin can set Employee screen module access: PPM / NTT / BOTH
- Managers/Admin always BOTH
- Employees see only their own uploads
- Employees can access only their own /uploads/* files (URL protection)
- Download button shown only for saved prints (modal), not for unsaved capture preview
- Delete button shown on RIGHT side (Admin only)
- Saved filenames generated as: {PPM|NTT}_{username}_{YYYYMMDD}_{HHMMSS}.png

NEW UPDATES:
- User Management list title: "User List"
- User list Number column is row count (1..N)
- User list filter: Active/Inactive/All
- Manager role has NO access to User Management (hard blocked in GET and POST)
- Admin on PPM/NTT pages:
  * Filter by Month (YYYY-MM)
  * Filter by Name/Email (matches uploader username/email)
- PPM/NTT UI:
  * Swap positions: Submitted(Locked) badge now in upload header (top-right)
  * Delete button now appears in timesheet header (right side)

UI UPDATE (THIS CHANGE):
- Updated Admin filter + button style on PPM/NTT pages (modern filter bar)

Default admin credentials:
  Email: admin@example.com
  Password: admin
"""

from http.server import SimpleHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlparse, quote
import os
import shutil
import sqlite3
import html
import datetime
import json  # required for DROPDOWN custom field options
import time  # used by email scheduler loop
import threading  # used by email scheduler thread startup
from io import BytesIO  # used by Excel export


# ---- Timezone helpers (IST cutoff) ----
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    IST = ZoneInfo("Asia/Kolkata")
except Exception:
    # Fallback: fixed offset if zoneinfo isn't available
    IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

UTC = datetime.timezone.utc


# --- UTC "now" helpers (timezone-aware) ---
def utc_now() -> datetime.datetime:
    return datetime.datetime.now(tz=UTC)

def utc_now_str(sep: str = ' ', timespec: str = 'seconds') -> str:
    return utc_now().isoformat(sep=sep, timespec=timespec)

# ---- IST display helpers ----
def parse_dt_assume_utc(dt_str: str):
    """Parse ISO-like datetime string. If timezone is missing, assume UTC."""
    if not dt_str:
        return None
    try:
        dt = datetime.datetime.fromisoformat(str(dt_str))
    except Exception:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt

def format_dt_ist(dt_str: str) -> str:
    """Return a friendly IST timestamp for display (YYYY-MM-DD HH:MM:SS IST)."""
    dt = parse_dt_assume_utc(dt_str)
    if not dt:
        return (dt_str or '')
    dt_ist = dt.astimezone(IST)
    return dt_ist.strftime('%Y-%m-%d %H:%M:%S IST')

import http.cookies
import uuid
import re
import logging
import traceback

from PIL import Image, ImageOps, ImageEnhance
# Excel export support
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except Exception:
    Workbook = None
    Font = None
    Alignment = None
    Border = None
    Side = None
    PatternFill = None
try:
    import pytesseract
    from pytesseract import Output
except Exception:
    pytesseract = None
    Output = None

# Point pytesseract to the local Tesseract executable on Windows
# (configured later after TESSERACT_DIR is defined)
# -------------------------
# Configuration
# -------------------------
DEBUG = True  # Set False in production
HOST = "127.0.0.1"
PORT = 8000
UPLOAD_DIR = "uploads"
DB_FILE = "users.db"


# -------------------------
# Feature flags / behavior toggles
# -------------------------
# NTT: If True, timesheet is auto-locked after OCR autofill (Submitted).
AUTO_LOCK_NTT_ON_OCR = True
# -------------------------
# Tesseract (Windows) runtime setup
# -------------------------
# Fix for Windows error: "libtesseract-5.dll was not found".
# Override by setting environment variable: TESSERACT_DIR
TESSERACT_DIR = os.environ.get(
    "TESSERACT_DIR",
    r"C:\\Users\\sindhu.sundara\\AppData\\Local\\Programs\\Tesseract-OCR"
)

if os.name == "nt":
    try:
        os.environ["PATH"] = TESSERACT_DIR + os.pathsep + os.environ.get("PATH", "")
        os.environ.setdefault("TESSDATA_PREFIX", os.path.join(TESSERACT_DIR, "tessdata"))
    except Exception:
        pass



# --- Ensure pytesseract can find Tesseract after TESSERACT_DIR/PATH setup (Windows-safe) ---
def _ensure_tesseract_cmd_ready():
    if pytesseract is None:
        return
    # On Windows, prefer the configured directory
    try:
        if os.name == 'nt':
            candidate = os.path.join(TESSERACT_DIR, 'tesseract.exe')
            if os.path.exists(candidate):
                pytesseract.pytesseract.tesseract_cmd = candidate
                return
    except Exception:
        pass
    # Fallback to PATH
    try:
        found = shutil.which('tesseract')
        if found and os.path.exists(found):
            pytesseract.pytesseract.tesseract_cmd = found
            return
    except Exception:
        pass

_ensure_tesseract_cmd_ready()
os.makedirs(UPLOAD_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler("server.log"),
        logging.StreamHandler()
    ]
)


# -------------------------
# Helpers
# -------------------------
def _safe_name_part(s: str, default="user"):
    s = (s or "").strip()
    if not s:
        return default
    s = re.sub(r"[^a-zA-Z0-9._-]+", "_", s)
    return s[:60] or default


def normalize_screen_access(val: str) -> str:
    val = (val or "BOTH").strip().upper()
    return val if val in ("PPM", "NTT", "EMAIL", "BOTH") else "BOTH"


# -------------------------
# DB connection helper (enables FK)
# -------------------------
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

def ensure_user_custom_fields_tables():
    """Create tables to support admin-defined custom user fields (migration-safe)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS user_custom_fields (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        field_key TEXT NOT NULL UNIQUE,
        label TEXT NOT NULL,
        field_type TEXT NOT NULL,
        required INTEGER NOT NULL DEFAULT 0,
        options_json TEXT,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS user_custom_field_values (
        user_id INTEGER NOT NULL,
        field_id INTEGER NOT NULL,
        value_text TEXT,
        PRIMARY KEY (user_id, field_id),
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
        FOREIGN KEY(field_id) REFERENCES user_custom_fields(id) ON DELETE CASCADE
    )
    """)
    conn.commit()
    conn.close()


def list_custom_fields(active_only: bool = True):
    conn = get_conn()
    cur = conn.cursor()
    if active_only:
        cur.execute("""
            SELECT id, field_key, label, field_type, required, COALESCE(options_json,''), active
            FROM user_custom_fields
            WHERE COALESCE(active,1)=1
            ORDER BY id
        """)
    else:
        cur.execute("""
            SELECT id, field_key, label, field_type, required, COALESCE(options_json,''), active
            FROM user_custom_fields
            ORDER BY id
        """)
    rows = cur.fetchall()
    conn.close()
    return rows


def create_custom_field(field_key: str, label: str, field_type: str, required: int = 0, options_list=None):
    field_key = (field_key or '').strip().lower()
    label = (label or '').strip()
    field_type = (field_type or 'TEXT').strip().upper()
    if field_type not in ('TEXT','NUMBER','DATE','BOOLEAN','DROPDOWN'):
        field_type = 'TEXT'
    if not field_key or not re.match(r'^[a-z][a-z0-9_]{1,50}$', field_key):
        raise ValueError('Field key must be snake_case, start with a letter, 2–50 chars.')
    if not label:
        raise ValueError('Label is required.')

    options_json = None
    if field_type == 'DROPDOWN':
        opts = []
        if options_list:
            for x in options_list:
                s = str(x).strip()
                if s:
                    opts.append(s)
        if not opts:
            raise ValueError('Dropdown requires at least one option.')
        options_json = json.dumps(opts)

    created_at = utc_now_str(sep=' ', timespec='seconds')
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO user_custom_fields (field_key, label, field_type, required, options_json, active, created_at)
        VALUES (?, ?, ?, ?, ?, 1, ?)
    """, (field_key, label, field_type, 1 if required else 0, options_json, created_at))
    conn.commit()
    conn.close()


def deactivate_custom_field(field_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE user_custom_fields SET active=0 WHERE id=?", (int(field_id),))
    conn.commit()
    conn.close()

def get_custom_field_by_id(field_id: int):
    """Return a single custom field row or None."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, field_key, label, field_type, required, COALESCE(options_json,''), COALESCE(active,1) "
        "FROM user_custom_fields WHERE id = ?",
        (int(field_id),),
    )
    row = cur.fetchone()
    conn.close()
    return row


def update_custom_field(field_id: int, label: str, required: int = 0, options_list=None, active: int = 1):
    """Update an existing field (label/required/options/active). Field key & type are immutable."""
    row = get_custom_field_by_id(field_id)
    if not row:
        raise ValueError('Field not found.')
    _id, _fkey, _lbl, ftype, _req, opt_json, _active = row
    ftype = (ftype or 'TEXT').strip().upper()
    label = (label or '').strip()
    if not label:
        raise ValueError('Label is required.')

    options_json = None
    if ftype == 'DROPDOWN':
        opts = []
        if options_list:
            for x in options_list:
                s = str(x).strip()
                if s:
                    opts.append(s)
        if not opts:
            raise ValueError('Dropdown requires at least one option.')

        # Ensure existing saved values remain valid
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "SELECT DISTINCT COALESCE(value_text,'') FROM user_custom_field_values WHERE field_id=?",
            (int(field_id),),
        )
        existing_vals = [r[0] for r in cur.fetchall() if (r[0] or '').strip()]
        conn.close()
        bad = [v for v in existing_vals if v not in opts]
        if bad:
            raise ValueError(
                'Cannot update dropdown options: existing user values would become invalid. '
                'Include these values or clear them first: ' + ', '.join(bad[:20])
            )
        options_json = json.dumps(opts)

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE user_custom_fields SET label=?, required=?, options_json=?, active=? WHERE id=?",
        (label, 1 if required else 0, options_json, 1 if active else 0, int(field_id)),
    )
    conn.commit()
    conn.close()


def delete_custom_field(field_id: int):
    """Hard delete a custom field and associated values (FK cascade)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM user_custom_fields WHERE id=?", (int(field_id),))
    conn.commit()
    conn.close()



def get_custom_field_values(user_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT field_id, COALESCE(value_text,'')
        FROM user_custom_field_values
        WHERE user_id = ?
    """, (int(user_id),))
    rows = cur.fetchall()
    conn.close()
    return {int(fid): (val or '') for fid, val in rows}


def upsert_custom_field_value(user_id: int, field_id: int, value_text: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO user_custom_field_values (user_id, field_id, value_text)
        VALUES (?, ?, ?)
        ON CONFLICT(user_id, field_id) DO UPDATE SET value_text=excluded.value_text
    """, (int(user_id), int(field_id), (value_text or '')))
    conn.commit()
    conn.close()


def delete_custom_field_value(user_id: int, field_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM user_custom_field_values WHERE user_id=? AND field_id=?", (int(user_id), int(field_id)))
    conn.commit()
    conn.close()


def _parse_options_json(opt_json: str):
    try:
        arr = json.loads(opt_json or '[]')
        if isinstance(arr, list):
            return [str(x) for x in arr]
    except Exception:
        pass
    return []


def build_custom_fields_form_html(fields_rows, values_dict=None, name_prefix='cf_'):
    """Return HTML inputs for custom fields.
    fields_rows: list rows (id, key, label, type, required, options_json, active)
    values_dict: dict field_id->value_text
    """
    values_dict = values_dict or {}
    if not fields_rows:
        return "<p class='muted' style='margin:10px 0 0'>No custom fields defined.</p>"

    out = ["<hr style='margin:16px 0;border:none;border-top:1px solid #eef2f7'>",
           "<h3 style='margin:0 0 6px'>Custom Fields</h3>"]
    for fid, fkey, label, ftype, req, opt_json, active in fields_rows:
        fid = int(fid)
        ftype = (ftype or 'TEXT').upper()
        req = int(req or 0)
        v = (values_dict.get(fid) or '').strip()
        safe_label = html.escape(label or fkey or '', quote=False)
        name = f"{name_prefix}{fid}"
        required_attr = ' required' if req else ''

        if ftype == 'NUMBER':
            out.append(f"<label>{safe_label}</label><input type='number' step='any' name='{name}' value='{html.escape(v, quote=True)}'{required_attr}>")
        elif ftype == 'DATE':
            out.append(f"<label>{safe_label}</label><input type='date' name='{name}' value='{html.escape(v, quote=True)}'{required_attr}>")
        elif ftype == 'BOOLEAN':
            sel1 = "selected" if v in ('1','true','yes','on') else ''
            sel0 = "selected" if v in ('0','false','no','off','') else ''
            out.append(
                f"<label>{safe_label}</label>"
                f"<select name='{name}'{required_attr}>"
                f"<option value='0' {sel0}>No</option>"
                f"<option value='1' {sel1}>Yes</option>"
                f"</select>"
            )
        elif ftype == 'DROPDOWN':
            opts = _parse_options_json(opt_json)
            opt_tags = []
            if not req:
                opt_tags.append("<option value=''>-- Select --</option>")
            for o in opts:
                sel = ' selected' if v == o else ''
                opt_tags.append(f"<option value='{html.escape(o, quote=True)}'{sel}>{html.escape(o)}</option>")
            out.append(
                f"<label>{safe_label}</label>"
                f"<select name='{name}'{required_attr}>" + "".join(opt_tags) + "</select>"
            )
        else:  # TEXT default
            out.append(f"<label>{safe_label}</label><input type='text' name='{name}' value='{html.escape(v, quote=True)}'{required_attr}>")

    return "".join(out)


def validate_and_save_custom_fields_for_user(user_id: int, form: dict, active_only=True, name_prefix='cf_'):
    """Validate custom field inputs from form dict and persist them."""
    fields = list_custom_fields(active_only=active_only)
    values_to_save = []

    for fid, fkey, label, ftype, req, opt_json, active in fields:
        fid = int(fid)
        ftype = (ftype or 'TEXT').upper()
        req = int(req or 0)
        k = f"{name_prefix}{fid}"
        raw = (form.get(k) if isinstance(form, dict) else '')
        # handle parse_qs lists
        if isinstance(raw, list):
            raw = raw[0] if raw else ''
        raw = (raw or '').strip()

        if req and raw == '':
            raise ValueError(f"Custom field '{label or fkey}' is required.")

        if raw == '':
            # remove value if exists and empty
            delete_custom_field_value(user_id, fid)
            continue

        if ftype == 'NUMBER':
            try:
                float(raw)
            except Exception:
                raise ValueError(f"Custom field '{label or fkey}' must be a number.")
        elif ftype == 'DATE':
            try:
                datetime.date.fromisoformat(raw)
            except Exception:
                raise ValueError(f"Custom field '{label or fkey}' must be a valid date (YYYY-MM-DD).")
        elif ftype == 'BOOLEAN':
            raw = '1' if raw.lower() in ('1','true','yes','on') else '0'
        elif ftype == 'DROPDOWN':
            opts = _parse_options_json(opt_json)
            if raw not in opts:
                raise ValueError(f"Custom field '{label or fkey}' must be one of the configured options.")

        values_to_save.append((fid, raw))

    for fid, val in values_to_save:
        upsert_custom_field_value(user_id, fid, val)

    return True


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
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM users")
    count = cur.fetchone()[0]
    if count == 0:
        cur.execute(
            "INSERT INTO users (username, email, password, role, status, screen_access, can_email, can_export, can_help) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            ("admin", "admin@example.com", "admin", "Admin", "Active", "BOTH", 1, 1, 1),
        )
        conn.commit()
    conn.close()

# -------------------------
# Help Content (DB-backed, Admin editable)
# -------------------------

def ensure_help_table_and_seed():
    """Create a single-row help_content table and seed with default HTML if empty."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS help_content (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            html TEXT NOT NULL
        )
    """)
    cur.execute("SELECT html FROM help_content WHERE id=1")
    row = cur.fetchone()
    if not row:
        default_html = (
            "<div class='card' style='max-width:960px;margin:0 auto'>"
            "<h2><i class='fa fa-circle-question'></i> System Help</h2>"
            "<p class='muted'>Quick guide for using this portal.</p>"
            "<h3>1) Screen Print Module (PPM / NTT)</h3>"
            "<ul>"
            "<li><b>Capture</b> to take a screenshot, then <b>Save</b> to upload.</li>"
            "<li>Timesheet is per upload and grouped by Week (Sun–Sat).</li>"
            "<li><b>Submit</b> locks the week and enforces PPM and NTT hours to match for the same week.</li>"
            "</ul>"
            "<h3>2) Email Settings (Admin/Manager)</h3>"
            "<ul>"
            "<li>Recipients are restricted to <b>Active users</b> only.</li>"
            "<li>Schedule uses <b>Date + Time (IST)</b>.</li>"
            "<li>Status shows <b>Pending</b>, <b>Sent</b>, or an error message.</li>"
            "</ul>"
            "<h3>3) SMTP Configuration (for sending emails)</h3>"
            "<p class='muted'>Set these environment variables on the server before starting the app:</p>"
            "<pre style='white-space:pre-wrap;background:#0b1220;color:#e6eef8;padding:12px;border-radius:10px'>"
            "SMTP_HOST=...\nSMTP_PORT=587\nSMTP_FROM=...\nSMTP_USER=... (optional)\nSMTP_PASS=... (optional)\nSMTP_SSL=false (or true for port 465)"
            "</pre>"
            "<h3>4) Troubleshooting</h3>"
            "<ul>"
            "<li>If Email status shows 'SMTP not configured', set SMTP_HOST and SMTP_FROM and restart the server.</li>"
            "<li>If export fails, ensure <b>openpyxl</b> is installed.</li>"
            "</ul>"
            "</div>"
        )
        cur.execute("INSERT INTO help_content (id, html) VALUES (1, ?)", (default_html,))
        conn.commit()
    conn.close()


def get_help_html() -> str:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT html FROM help_content WHERE id=1")
    row = cur.fetchone()
    conn.close()
    return (row[0] if row and row[0] else "")


def set_help_html(new_html: str) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE help_content SET html=? WHERE id=1", ((new_html or ""),))
    if cur.rowcount == 0:
        cur.execute("INSERT INTO help_content (id, html) VALUES (1, ?)", ((new_html or ""),))
    conn.commit()
    conn.close()


def basic_sanitize_html(s: str) -> str:
    """Basic safety: remove script blocks and inline on* handlers (best-effort)."""
    s = s or ""
    s = re.sub(r"(?is)<script.*?>.*?</script>", "", s)
    s = re.sub(r"(?is)\son\w+\s*=\s*[^\s>]+", "", s)
    return s




def username_exists(username, exclude_user_id=None):
    conn = get_conn()
    cur = conn.cursor()
    if exclude_user_id:
        cur.execute("SELECT id FROM users WHERE username = ? AND id != ?", (username, exclude_user_id))
    else:
        cur.execute("SELECT id FROM users WHERE username = ?", (username,))
    row = cur.fetchone()
    conn.close()
    return bool(row)


def email_exists(email, exclude_user_id=None):
    conn = get_conn()
    cur = conn.cursor()
    if exclude_user_id:
        cur.execute("SELECT id FROM users WHERE email = ? AND id != ?", (email, exclude_user_id))
    else:
        cur.execute("SELECT id FROM users WHERE email = ?", (email,))
    row = cur.fetchone()
    conn.close()
    return bool(row)


def create_user_db(username, email, password, role, status, screen_access="BOTH", can_email=0, can_export=0, can_help=1):
    """Create a user with Screen Print access + module access flags.
    screen_access: PPM / NTT / BOTH (Employees and Managers configurable; Admin forced BOTH).
    can_email/can_export/can_help: 0/1 flags for feature access (Admin forced enabled).
    """
    username = (username or "").strip()
    email = (email or "").strip().lower()
    role = (role or "Employee").strip()
    status = (status or "Active").strip()
    if not username or not email:
        return False
    if username_exists(username) or email_exists(email):
        return False
    screen_access = normalize_screen_access(screen_access)
    def _to01(v, default=0):
        try:
            return 1 if str(v).strip().lower() in ("1", "true", "on", "yes") else 0
        except Exception:
            return default
    can_email = _to01(can_email, 0)
    can_export = _to01(can_export, 0)
    can_help = _to01(can_help, 1)
    # Employee: Help must be OFF and locked
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
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        try: conn.close()
        except Exception: pass

def list_users_db(status_filter=None):
    """ 
    status_filter: None, 'Active', or 'Inactive'
    Returns tuples:
      (id, username, email, role, status, screen_access, can_email, can_export, can_help)
    """
    conn = get_conn()
    cur = conn.cursor()
    base_sql = (
        "SELECT id, username, email, role, status, COALESCE(screen_access,'BOTH'), "
        "COALESCE(can_email,0), COALESCE(can_export,0), COALESCE(can_help,1) "
        "FROM users "
    )
    if status_filter in ("Active", "Inactive"):
        cur.execute(base_sql + "WHERE status = ? ORDER BY id", (status_filter,))
    else:
        cur.execute(base_sql + "ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    return rows
def list_active_users_for_reminders():
    """Return list of (username, email) for Active users only."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(username,''), LOWER(COALESCE(email,'')) FROM users WHERE status='Active' ORDER BY username")
    rows = cur.fetchall()
    conn.close()
    return [(r[0] or '', (r[1] or '').strip().lower()) for r in rows if (r[1] or '').strip()]


def is_active_user_email(email: str) -> bool:
    e = (email or '').strip().lower()
    if not e:
        return False
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM users WHERE LOWER(email)=LOWER(?) AND status='Active'", (e,))
    ok = cur.fetchone() is not None
    conn.close()
    return ok


def get_user_db_by_id(user_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, username, email, role, status, COALESCE(screen_access,'BOTH'), "
        "COALESCE(can_email,0), COALESCE(can_export,0), COALESCE(can_help,1) "
        "FROM users WHERE id = ?",
        (user_id,),
    )
    row = cur.fetchone()
    conn.close()
    return row


def get_user_db_by_email(email):
    if not email:
        return None
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, username, email, password, role, status, COALESCE(screen_access,'BOTH'), "
        "COALESCE(can_email,0), COALESCE(can_export,0), COALESCE(can_help,1) "
        "FROM users WHERE email = ?",
        (email.lower(),),
    )
    row = cur.fetchone()
    conn.close()
    return row
def get_screen_access_by_email(email):
    row = get_user_db_by_email(email)
    if not row:
        return "BOTH"
    role = row[4]
    if role == "Admin":
        return "BOTH"
    return normalize_screen_access(row[6])




def get_feature_flags_by_email(email: str):
    """Return (can_email, can_export, can_help) for a given email."""
    row = get_user_db_by_email(email)
    if not row:
        return (0, 0, 0)
    role = row[4]
    if role == "Admin":
        return (1, 1, 1)
    return (int(row[7] or 0), int(row[8] or 0), int(row[9] or 0))


def get_feature_flags_by_username(username: str):
    """Return (can_email, can_export, can_help) for a given username."""
    u = (username or '').strip()
    if not u:
        return (0, 0, 0)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT role, COALESCE(can_email,0), COALESCE(can_export,0), COALESCE(can_help,1) "
        "FROM users WHERE username = ?",
        (u,),
    )
    row = cur.fetchone()
    conn.close()
    if not row:
        return (0, 0, 0)
    role, ce, cx, ch = row
    if role == "Admin":
        return (1, 1, 1)
    return (int(ce or 0), int(cx or 0), int(ch or 0))

def update_user_db(user_id, updates: dict):
    if not updates:
        return False

    existing = get_user_db_by_id(user_id)
    existing_role = existing[3] if existing else "Employee"
    target_role = (updates.get("role") or existing_role or "Employee").strip()
    # Enforce Admin module access server-side (cannot be downgraded)
    force_admin_flags = (target_role == "Admin")
    if force_admin_flags:
        updates["can_email"] = 1
        updates["can_export"] = 1
        updates["can_help"] = 1


    # Employee: Help must be OFF and locked
    if target_role == "Employee":
        updates["can_help"] = 0

    if "username" in updates:
        new_username = (updates["username"] or "").strip()
        if not new_username:
            return False
        if username_exists(new_username, exclude_user_id=user_id):
            return False
        updates["username"] = new_username

    if "email" in updates:
        new_email = (updates["email"] or "").strip().lower()
        if not new_email:
            return False
        if email_exists(new_email, exclude_user_id=user_id):
            return False
        updates["email"] = new_email

    if "screen_access" in updates:
        sa = normalize_screen_access(updates.get("screen_access"))
        # Admin is always BOTH; Manager/Employee can be configured
        if target_role == "Admin":
            sa = "BOTH"
        updates["screen_access"] = sa
        # normalize feature flags if present
        for flag in ("can_email", "can_export", "can_help"):
            if flag in updates:
                try:
                    updates[flag] = 1 if str(updates[flag]).strip().lower() in ("1","true","on","yes") else 0
                except Exception:
                    updates[flag] = 0
    else:
        # If role changes to Admin, enforce BOTH
        if "role" in updates and target_role == "Admin":
            updates["screen_access"] = "BOTH"

    set_clause = ", ".join(f"{k}=?" for k in updates.keys())
    params = list(updates.values()) + [user_id]

    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(f"UPDATE users SET {set_clause} WHERE id = ?", params)
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def delete_user_db(user_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()


# -------------------------
# Week helpers (Sun–Sat)
# -------------------------
def _sunday_of_date(d: datetime.date) -> datetime.date:
    days_to_sub = (d.weekday() + 1) % 7
    return d - datetime.timedelta(days=days_to_sub)


def _week_start_from_iso_datetime(uploaded_at_str: str) -> str:
    try:
        dt = datetime.datetime.fromisoformat(uploaded_at_str)
    except Exception:
        dt = utc_now()
    ws = _sunday_of_date(dt.date())
    return ws.isoformat()

# -------------------------
# PPM OCR (auto-fill timesheet from Total row)
# -------------------------
MONTHS_MAP = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
}

def _preprocess_for_ocr(img: Image.Image) -> Image.Image:
    # ensure it's RGB (no alpha) before processing
    rgb = img.convert('RGB')
    gray = ImageOps.grayscale(rgb)
    gray = ImageEnhance.Contrast(gray).enhance(2.0)
    gray = gray.resize((gray.size[0] * 2, gray.size[1] * 2))
    # Original threshold (170) works better for general readability
    bw = gray.point(lambda x: 0 if x < 170 else 255, '1')
    return bw

def _parse_week_start_from_ppm_text(text: str):
    if not text:
        return None
    # Normalize OCR whitespace to single space
    t = ' '.join(text.split()).lower()
    
    # regex for "Month D to (Month) D, YYYY" or variations
    # month_pat: Match first 3 chars and allow any number of letters after.
    month_p = r"(jan[a-z]*|feb[a-z]*|mar[a-z]*|apr[a-z]*|may[a-z]*|jun[a-z]*|jul[a-z]*|aug[a-z]*|sep[a-z]*|oct[a-z]*|nov[a-z]*|dec[a-z]*)"
    
    # 1: month1
    # 2: day1
    # 3: optional month2
    # 4: day2
    # 5: year
    m = re.search(
        month_p + r"\s+(\d{1,2})\s+(?:to|t0|1o|f0|[a-z]{1,3})\s+(?:" + month_p + r"\s+)?(\d{1,2})[,\s]+(\d{4})",
        t,
        re.IGNORECASE,
    )
    if not m:
        # Fallback to Month Day, Year if "to" middle part failed
        m = re.search(month_p + r"\s+(\d{1,2}).*?(\d{4})", t, re.IGNORECASE)
        if not m: return None
        
        month1_str = m.group(1).lower()[:3]
        day1_int = int(m.group(2))
        year_int = int(m.group(3))
    else:
        month1_str = m.group(1).lower()[:3]
        day1_int = int(m.group(2))
        year_int = int(m.group(5))
        
    try:
        month1_val = MONTHS_MAP.get(month1_str)
        if not month1_val: return None
        start = datetime.date(year_int, month1_val, day1_int)
        ws = _sunday_of_date(start)
        return ws.isoformat()
    except Exception:
        return None

def _to_hour_value(token: str):
    if token is None:
        return None
    s = str(token).strip().lower()
    if s in ('oh', 'o h', '0h', '0 h', '0', 'o'):
        return 0.0
    
    # Handle common OCR misreads (g=9, q=9, s=5, l/i=1, o=0)
    # only for the first character if it's likely a number
    if s:
        c = s[0]
        if c == 'g' or c == 'q': s = '9' + s[1:]
        elif c == 's': s = '5' + s[1:]
        elif c == 'o' and (len(s) == 1 or s[1] == 'h' or s[1] == ' '): s = '0' + s[1:]
        elif (c == 'l' or c == 'i') and (len(s) == 1 or s[1] == 'h' or s[1] == ' '): s = '1' + s[1:]

    # Remove all spaces and pick out numeric parts
    s = s.replace(' ', '').replace(',', '.')
    # Extract just the first numeric part (digits and dots)
    import re
    nums = re.findall(r'[0-9.]+', s)
    if not nums:
        return None
    
    try:
        v = float(nums[0])
        if v < 0: v = 0.0
        if v > 24: return None
        # Round to nearest 0.25 (standard timesheet step)
        v = round(v * 4) / 4.0
        return v
    except Exception:
        return None

def extract_ppm_hours_from_screenshot_total_row(image_path: str):
    if pytesseract is None or Output is None: return None
    try: img = Image.open(image_path)
    except Exception: return None
    proc = _preprocess_for_ocr(img)

    # Helper function to extract a row of numbers using a given PSM config
    def _extract_row_with_psm(config_str: str, get_week_start: bool = False):
        try: data = pytesseract.image_to_data(proc, output_type=Output.DICT, config=config_str)
        except Exception: return [], None, None
        
        words = []
        for i in range(len(data.get('text', []))):
            txt = (data['text'][i] or '').strip()
            if not txt: continue
            try:
                x, y, w, h = float(data['left'][i]), float(data['top'][i]), float(data['width'][i]), float(data['height'][i])
                words.append({'t': txt, 'tl': txt.lower(), 'cx': x + w / 2.0, 'cy': y + h / 2.0, 'w': w})
            except (ValueError, TypeError): continue

        week_start = None
        if get_week_start:
            full_text = ' '.join([str(w['t']) for w in words])
            week_start = _parse_week_start_from_ppm_text(full_text)

        totals = sorted([w for w in words if str(w['tl']) in ('total', 'totel', 'tota1')], key=lambda x: x['cy'])
        if not totals:
            totals = sorted([w for w in words if w['cx'] < 150.0 and 'tot' in str(w['tl'])], key=lambda x: x['cy'])
        if not totals: return [], week_start, None
        
        total_y = totals[-1]['cy']
        total_x = totals[-1]['cx']

        y_band = 30.0
        row_words = sorted([w for w in words if abs(w['cy'] - total_y) <= y_band and w['cx'] > total_x + 15.0], key=lambda x: x['cx'])
        
        row_groups = []
        if row_words:
            curr = [row_words[0]]
            for i in range(1, len(row_words)):
                if (row_words[i]['cx'] - row_words[i-1]['cx']) < 50.0:
                    curr.append(row_words[i])
                else:
                    row_groups.append(' '.join([str(x['t']) for x in curr]))
                    curr = [row_words[i]]
            row_groups.append(' '.join([str(x['t']) for x in curr]))

        def _to_hour_value(s: str):
            s = str(s).lower().replace(' ', '').replace('hour', '').replace('s', '').replace(':', '.')
            s = s.replace('g', '9').replace('q', '9').replace('o', '0')
            c = re.sub(r'[^0-9.]+', '', s)
            if hasattr(c, 'count') and c.count('.') > 1:
                c = c.replace('.', '', c.count('.') - 1)
            if not c: return None
            return float(c)

        vals = []
        for txt in row_groups:
            v = _to_hour_value(txt)
            if v is not None:
                vals.append(float(v))

        remaining_planned = None
        if len(vals) >= 9:
            remaining_planned = vals[8]
        elif len(vals) == 8 and len(row_groups) >= 9:
             # Case where 9th group wasn't a clean float but RP exists
             remaining_planned = _to_hour_value(row_groups[8])

        remaining_planned = None
        if len(vals) >= 9:
            remaining_planned = vals[8]
        elif len(vals) == 8 and len(row_groups) >= 9:
             remaining_planned = _to_hour_value(row_groups[8])

        # ROBUST FALLBACK 2: Search for any numeric value near "Remaining Planned" keywords
        if remaining_planned is None:
            rem_words = [w for w in words if "rem" in w['tl']]
            plan_words = [w for w in words if "planned" in w['tl'] or "plan" in w['tl']]
            
            # Find the header column center
            anchor_x = None
            if plan_words:
                anchor_x = sum(w['cx'] for w in plan_words) / len(plan_words)
            elif rem_words:
                anchor_x = sum(w['cx'] for w in rem_words) / len(rem_words)
                
            if anchor_x is not None:
                # Look for numeric values in this column
                candidates = []
                for w in words:
                    if abs(w['cx'] - anchor_x) < 100:
                        v = _to_hour_value(w['t'])
                        if v is not None and w['cy'] > 150:
                            candidates.append(v)
                if candidates:
                    # If multiple, the smallest or the specific one (often 0.2 vs 45)
                    # We'll take the sum or the one that's clearly a planned value.
                    remaining_planned = sum(candidates)

        return vals, week_start, remaining_planned

    # Run PSM 11 and PSM 6 to mitigate Tesseract misreading '9' as '0' on certain modes
    vals_11, week_start, rem11 = _extract_row_with_psm('--psm 11', get_week_start=True)
    if not week_start:
        # Fallback to PSM 3 for a better shot at finding the header date
        _, week_start, _ = _extract_row_with_psm('--psm 3', get_week_start=True)
        if not week_start:
            # Final attempt for week_start with PSM 6
            _, week_start, _ = _extract_row_with_psm('--psm 6', get_week_start=True)

    vals_6, _, rem6 = _extract_row_with_psm('--psm 6', get_week_start=False)

    day_keys = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat']
    result = {
        'week_start': str(week_start) if week_start else "",
        '_found_days': 0,
        'remaining_planned': rem11 if rem11 is not None else rem6
    }
    for dk in day_keys: result[dk] = 0.0

    # Combine results by taking the max for each column.
    # A true '0' will be '0' in both, while a '9' misread as '0' in one mode will be naturally corrected by taking the max.
    v_len = max(len(vals_11), len(vals_6))
    for i in range(min(v_len, 7)):
        v11 = vals_11[i] if i < len(vals_11) else 0.0
        v6  = vals_6[i] if i < len(vals_6) else 0.0
        result[day_keys[i]] = max(v11, v6)
        result['_found_days'] += 1

    if DEBUG:
        try: logging.info('PPM OCR dual-pass row values: 11=%s, 6=%s => %s', vals_11, vals_6, result)
        except Exception: pass

    return result


# -------------------------
# NTT OCR (auto-fill timesheet by summing NTT "Entered" values per day)
# -------------------------
def _parse_date_from_ntt_row(text: str):
    """Parse strings like:
        'Sunday, February 23, 2026' OR 'February 23, 2026'
       Return datetime.date or None.
    """
    if not text:
        return None
    t = ' '.join(text.replace(',', ' ').replace('.', ' ').replace('/', ' ').replace('-', ' ').split())
    # Fuzzier month matching to handle OCR errors (e.g. 'Feo' or 'Feu')
    # Support both full month names and common 3-letter abbreviations
    month_regex = r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    
    # Try multiple common NTT formats
    # 1. Standard: [Day] Month Day Year (e.g. Monday February 24 2026)
    m1 = re.search(
        r"(?:(Sun|Mon|Tue|Wed|Thu|Fri|Sat)(?:day|day|sday|nesday|rsday|day|urday)?\s+)?"
        + month_regex + r"\s+"
        r"(\d{1,2})\s+(\d{4})",
        t, re.IGNORECASE
    )
    if m1:
        month_idx = MONTHS_MAP.get((m1.group(2) or '').lower())
        if month_idx:
            try: return datetime.date(int(m1.group(4)), int(month_idx), int(m1.group(3)))
            except: pass

    # 2. Alternative: Day Month name Year (e.g. 24 Feb 2026)
    m2 = re.search(r"(\d{1,2})\s+" + month_regex + r"\s+(\d{4})", t, re.IGNORECASE)
    if m2:
        month_idx = MONTHS_MAP.get((m2.group(2) or '').lower())
        if month_idx:
            try: return datetime.date(int(m2.group(3)), int(month_idx), int(m2.group(1)))
            except: pass

    # 3. Fallback: Just day name + any date nearby (e.g. "Tuesday 24") - used to split blocks
    m3 = re.search(r"(Sun|Mon|Tue|Wed|Thu|Fri|Sat)(?:day|day|sday|nesday|rsday|day|urday)?\s+(\d{1,2})", t, re.IGNORECASE)
    if m3:
        # We don't have year/month here, but the block logic can still use it to mark a new day
        # if we assume it's near the current month
        pass

    return None

def extract_ntt_hours_from_screenshot_entered_column(image_path: str):
    """Extract NTT daily hours from a screenshot of the NTT timesheet page.

    Strategy:
    1. Run OCR (PSM 6 preferred — detects all 7 day headers reliably).
    2. Find the "Entered" column header (cx ~ 1860).
    3. Find each day header row ("Sunday,", "Monday,", …).
    4. For each day, find "X.XX" tokens near the Entered column X and
       close to (but below) the day header Y.  Sum them for that day.
    """
    if pytesseract is None or Output is None:
        return None
    try:
        img = Image.open(image_path)
    except Exception:
        return None

    proc = _preprocess_for_ocr(img)

    def _run_ocr(config_str: str):
        try:
            data = pytesseract.image_to_data(proc, output_type=Output.DICT, config=config_str)
            out = []
            for i in range(len(data.get('text', []))):
                txt = (data['text'][i] or '').strip()
                if not txt:
                    continue
                x = float(data['left'][i])
                y = float(data['top'][i])
                w = float(data['width'][i])
                h = float(data['height'][i])
                out.append({'t': txt, 'tl': txt.lower(), 'cx': x + w / 2.0, 'cy': y + h / 2.0})
            return out
        except Exception as e:
            import traceback
            logging.error("OCR internal failure: %s\n%s", e, traceback.format_exc())
            return []

    # Run BOTH PSM modes and merge results to maximize detection
    # PSM 6 may miss some day headers (e.g. Friday); PSM 11 catches them
    words_6 = _run_ocr('--psm 6')
    words_11 = _run_ocr('--psm 11')
    # Use PSM 6 as the base, supplement with PSM 11 words not already present
    words = list(words_6) if words_6 else []
    if words_11:
        # Use a coordinate-based key to deduplicate identical words between passes
        # we round coordinates to 20px bins
        existing_keys = set()
        for w in words:
            existing_keys.add( (w['t'].lower(), int(float(w['cx'])/20), int(float(w['cy'])/20)) )
        for w in words_11:
            key = (w['t'].lower(), int(float(w['cx'])/20), int(float(w['cy'])/20))
            if key not in existing_keys:
                words.append(w)
                existing_keys.add(key)
    if not words:
        return None

    # ── 1. Locate column headers to find "Entered" column X ────────
    col_x = {}
    entered_y = 150.0 # fallback
    for w in words:
        tl = w['tl'].rstrip('.,:;')
        if tl in ('entered', 'entrd', 'enlered', 'enfered', 'entored', 'enterod'):
            if w['cy'] > 100: # avoid top-of-page buttons
                col_x['entered'] = float(w['cx'])
                entered_y = float(w['cy'])
        elif tl in ('recorded', 'target'):
            if w['cy'] > 100: col_x['recorded'] = float(w['cx'])
        elif tl in ('draft', 'drafl', 'draît'):
            if w['cy'] > 100: col_x['draft'] = float(w['cx'])
        elif tl in ('status', 'stafus', 'stats'):
            if w['cy'] > 100: col_x['status'] = float(w['cx'])
        elif tl in ('assignment', 'assign', 'assgn'):
            if w['cy'] > 100: col_x['assignment'] = float(w['cx'])

    if 'entered' in col_x:
        entered_x = col_x['entered']
    elif 'draft' in col_x:
        entered_x = col_x['draft'] - 180
    elif 'status' in col_x:
        entered_x = col_x['status'] - 400
    elif 'recorded' in col_x:
        entered_x = col_x['recorded'] + 850
    elif 'assignment' in col_x:
        entered_x = col_x['assignment'] + 500
    else:
        # Fallback to 75% of image width
        entered_x = img.size[0] * 0.75

    if DEBUG:
        logging.info("NTT OCR: col_x=%s, using entered_x=%.1f", col_x, entered_x)

    # ── 2. Locate day-header rows (e.g. "Sunday, February 22, 2026") ──
    # OCR produces tokens like  "Sunday,"  "Monday,"  "Wednesday,"  etc.
    # IMPORTANT: We must match FULL day names ("sunday", "monday", etc.)
    # NOT short 3-letter calendar headers ("Sun", "Mon") which all share the same Y,
    # and NOT usernames like "Sundaramoorthy" that start with "sun".
    DAY_NAMES = {
        'sunday': 'sun', 'monday': 'mon', 'tuesday': 'tue', 'wednesday': 'wed',
        'thursday': 'thu', 'friday': 'fri', 'saturday': 'sat',
        # Handle common OCR misspellings
        'wedinesday': 'wed', 'wednsday': 'wed', 'wednseday': 'wed',
        'thurday': 'thu', 'thuesday': 'tue', 'firday': 'fri',
    }
    day_header_ys = {}  # day_key -> cy
    for w in words:
        tl = w['tl'].rstrip('.,;:!')  # strip trailing punctuation from OCR
        for full_name, day_key in DAY_NAMES.items():
            if tl == full_name and day_key not in day_header_ys:
                day_header_ys[day_key] = float(w['cy'])
                break

    if DEBUG:
        logging.info("NTT OCR: day_header_ys = %s", day_header_ys)

    # ── 3. Extract week_start (Robust Month Detection) ─────────────
    found_date = None
    line_bins = {}
    for w in words:
        ybin = int(float(w['cy']) / 15) * 15
        if ybin not in line_bins:
            line_bins[ybin] = []
        line_bins[ybin].append(w)
    
    # Try preferred row-based parsing (most accurate)
    for ybin in sorted(line_bins.keys()):
        line_text = ' '.join(str(w['t']) for w in sorted(line_bins[ybin], key=lambda w: float(w['cx'])))
        dt = _parse_date_from_ntt_row(line_text)
        if dt:
            found_date = dt
            break

    # If row-based parsing fails, scan all tokens for month and year
    if not found_date:
        possible_months = []
        possible_years = []
        possible_days = []
        for w in words:
            tl = w['tl'].rstrip('.,;:!')
            if tl in MONTHS_MAP:
                possible_months.append((MONTHS_MAP[tl], float(w['cy'])))
            if re.match(r'^(202[4-9]|203[0-9])$', tl):
                possible_years.append((int(tl), float(w['cy'])))
            if re.match(r'^([1-9]|[12][0-9]|3[01])$', tl):
                possible_days.append((int(tl), float(w['cy'])))
        
        if possible_months and possible_years:
            # Pick the lowest (top-most) month and year if multiple
            month = sorted(possible_months, key=lambda x: x[1])[0][0]
            year = sorted(possible_years, key=lambda x: x[1])[0][0]
            # Try to find a day (like "5") near that month
            day = 1
            if possible_days:
                # Find day token closest to the month token vertically
                month_y = sorted(possible_months, key=lambda x: x[1])[0][1]
                day = min(possible_days, key=lambda x: abs(x[1] - month_y))[0]
            try:
                found_date = datetime.date(year, month, day)
                logging.info("NTT OCR: Inferred date %s from tokens", found_date)
            except:
                pass

    week_start_str = ""
    if found_date:
        week_start_str = _sunday_of_date(found_date).isoformat()
    else:
        logging.warning("NTT OCR: No date found in screenshot")

    # ── 4. Pick numeric tokens that are genuine hour values ────────
    # A genuine hour token looks like  "9.00"  "0.00"  "8.00!"
    # NOT "Hours", "Monday,", "Submitted", calendar numbers, etc.
    def _parse_entered_value(raw: str):
        # Keep digits and dots
        cleaned = re.sub(r'[^0-9.]+', '', raw.replace(',', '.'))
        if not cleaned: return None
        # Heuristic: NTT values are usually x.xx
        if '.' not in cleaned: return None
        try:
            v = float(cleaned)
            if 0.0 <= v <= 24.0: return v
        except: pass
        return None

    # ── 5. Identify "Hours" anchors for precise 'Entered' matching ─
    hours_anchors = [w for w in words if w['tl'] in ('hours', 'hour', 'hrs', 'hr', 'hars', 'haurs')]
    
    # ── 6. For each day, find the Entered value(s) ─────────────────
    days_order = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat']
    result = {
        'sun': 0.0, 'mon': 0.0, 'tue': 0.0, 'wed': 0.0,
        'thu': 0.0, 'fri': 0.0, 'sat': 0.0,
        '_found_days': 0,
        'has_draft': False,
    }
    if week_start_str:
        result['week_start'] = str(week_start_str)

    if day_header_ys:
        # Build ordered list of (day_key, header_y) sorted by Y
        ordered_days = sorted(day_header_ys.items(), key=lambda kv: kv[1])

        for idx, (day_key, hdr_y) in enumerate(ordered_days):
            # Define Y-band: from this header Y to next header Y (or +200 if last)
            if idx + 1 < len(ordered_days):
                next_y = ordered_days[idx + 1][1]
            else:
                next_y = hdr_y + 200.0

            # Collect genuine hour values in the Entered column within this Y-band
            day_total = 0.0
            found_any = False
            seen_cells = set() # (int(cx/30), int(cy/10)) -> avoid double counting same cell
            for w in words:
                wy = float(w['cy'])

                # Draft detection
                if wy >= hdr_y - 10 and wy < next_y:
                    if w['tl'] in ('draft', 'drafl', 'draît') and w['cx'] > entered_x - 100:
                        result['has_draft'] = True

                if wy < hdr_y - 10 or wy >= next_y: continue
                
                v = _parse_entered_value(w['t'])
                if v is not None:
                    is_entered_col = False
                    # Check 1: "Hours" anchor (primary indicator)
                    for anchor in hours_anchors:
                        if abs(float(anchor['cy']) - wy) < 35:
                            dist = float(anchor['cx']) - float(w['cx'])
                            if 10 < dist < 350:
                                is_entered_col = True
                                break
                    
                    # Check 2: Horizontal alignment (secondary indicator)
                    if not is_entered_col:
                        if abs(float(w['cx']) - entered_x) < 200:
                            # Safety: ensures it's not the Recorded column if we know where it is
                            if 'recorded' in col_x and float(w['cx']) < col_x['recorded'] + 200:
                                pass # skip, likely recorded column
                            else:
                                is_entered_col = True

                    if is_entered_col:
                        cell_key = (int(float(w['cx'])/50), int(float(w['cy'])/15))
                        if cell_key not in seen_cells:
                            day_total += v
                            found_any = True
                            seen_cells.add(cell_key)

            result[day_key] = float(round(day_total * 4) / 4.0)
            if found_any: result['_found_days'] += 1
    else:
        # Fallback: collect all Entered-column numeric values, cluster by Y
        vals = []
        for w in words:
            dx = abs(float(w['cx']) - entered_x)
            if dx > 80:
                continue
            v = _parse_entered_value(w['t'])
            if v is not None:
                vals.append((float(w['cy']), v))
        vals.sort(key=lambda x: x[0])

        clusters = []
        for cy, v in vals:
            if not clusters:
                clusters.append([(cy, v)])
            elif cy - clusters[-1][-1][0] < 60:
                clusters[-1].append((cy, v))
            else:
                clusters.append([(cy, v)])

        if len(clusters) == 7:
            for i, c in enumerate(clusters):
                s = sum(vv for _, vv in c)
                result[days_order[i]] = float(round(s * 4) / 4.0)
                result['_found_days'] = int(result['_found_days']) + 1

        # Fallback global check for Draft
        for w in words:
            if w['tl'] in ('draft', 'drafl', 'draît') and w['cy'] > entered_y + 40:
                result['has_draft'] = True
                break

    if DEBUG:
        logging.info("NTT OCR result: %s", result)

    return result
# -------------------------
# Upload + timesheet operations
# -------------------------
def ensure_timesheet_for_upload(upload_id: int, uploaded_at: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM timesheets WHERE upload_id = ?", (upload_id,))
    if cur.fetchone():
        conn.close()
        return

    week_start = _week_start_from_iso_datetime(uploaded_at)
    updated_at = utc_now_str(sep=" ", timespec="seconds")
    cur.execute("""
        INSERT INTO timesheets (
            upload_id, week_start, mon, tue, wed, thu, fri, sat, sun, total,
            submitted, submitted_at, updated_at
        )
        VALUES (?, ?, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, ?)
    """, (upload_id, week_start, updated_at))
    conn.commit()
    conn.close()


def add_upload_db(stored_filename, original_filename, module, uploaded_by):
    uploaded_at = utc_now_str(sep=" ", timespec="seconds")
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO uploads (stored_filename, original_filename, uploaded_at, module, uploaded_by) "
        "VALUES (?, ?, ?, ?, ?)",
        (stored_filename, original_filename, uploaded_at, module, (uploaded_by or "").strip().lower()),
    )
    conn.commit()
    upload_id = cur.lastrowid
    conn.close()

    ensure_timesheet_for_upload(upload_id, uploaded_at)
    return upload_id


def list_uploads_db(module, uploaded_by=None, month=None, name=None):
    """
    Returns rows as:
      (id, stored_filename, original_filename, uploaded_at, module, uploaded_by_email, uploaded_by_username)

    Employee view: pass uploaded_by (email) -> only their uploads.
    Admin/Manager view: uploaded_by=None -> all uploads for that module, optionally filtered:
      month: 'YYYY-MM' on uploaded_at
      name: search on users.username OR uploads.uploaded_by (email)
    """
    conn = get_conn()
    cur = conn.cursor()

    where = ["u.module = ?"]
    params = [module]

    if uploaded_by:
        where.append("LOWER(COALESCE(u.uploaded_by,'')) = LOWER(?)")
        params.append((uploaded_by or "").strip().lower())

    if month and re.match(r"^\d{4}-\d{2}$", month):
        where.append("substr(u.uploaded_at, 1, 7) = ?")
        params.append(month)

    if name:
        name_like = f"%{name.strip().lower()}%"
        where.append("(LOWER(COALESCE(usr.username,'')) LIKE ? OR LOWER(COALESCE(u.uploaded_by,'')) LIKE ?)")
        params.extend([name_like, name_like])

    sql = f"""
        SELECT u.id, u.stored_filename, u.original_filename, u.uploaded_at, u.module,
               LOWER(COALESCE(u.uploaded_by,'')) AS uploaded_by_email,
               COALESCE(usr.username,'') AS uploaded_by_username
          FROM uploads u
          LEFT JOIN users usr
                 ON LOWER(COALESCE(usr.email,'')) = LOWER(COALESCE(u.uploaded_by,''))
         WHERE {" AND ".join(where)}
         ORDER BY u.id DESC
    """
    cur.execute(sql, tuple(params))
    rows = cur.fetchall()
    conn.close()
    return rows


def get_upload_owner(upload_id: int) -> str:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT LOWER(COALESCE(uploaded_by,'')) FROM uploads WHERE id = ?", (upload_id,))
    row = cur.fetchone()
    conn.close()
    return (row[0] or "").strip().lower() if row else ""


def get_upload_owner_by_stored_filename(stored_filename: str) -> str:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT LOWER(COALESCE(uploaded_by,'')) FROM uploads WHERE stored_filename = ?", (stored_filename,))
    row = cur.fetchone()
    conn.close()
    return (row[0] or "").strip().lower() if row else ""

def get_upload_module(upload_id: int) -> str:
    """Return module string for an upload (PPM/NTT/Generic)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(module,'Generic') FROM uploads WHERE id = ?", (upload_id,))
    row = cur.fetchone()
    conn.close()
    return (row[0] or 'Generic').strip()


def get_timesheet_by_upload(upload_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT week_start, mon, tue, wed, thu, fri, sat, sun, total, updated_at,
               COALESCE(submitted,0), COALESCE(submitted_at,''), remaining_planned,
               COALESCE(has_draft,0)
        FROM timesheets WHERE upload_id = ?
    """, (upload_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return None
    return {
        "week_start": row[0],
        "mon": row[1], "tue": row[2], "wed": row[3], "thu": row[4],
        "fri": row[5], "sat": row[6], "sun": row[7],
        "total": row[8],
        "updated_at": row[9],
        "submitted": int(row[10] or 0),
        "submitted_at": row[11] or "",
        "remaining_planned": row[12],
        "has_draft": int(row[13] or 0)
    }


def get_timesheet_values(upload_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT week_start, sun, mon, tue, wed, thu, fri, sat, total, COALESCE(submitted,0), updated_at, remaining_planned
        FROM timesheets WHERE upload_id = ?
    """, (upload_id,))
    row = cur.fetchone()
    conn.close()
    return row


def is_timesheet_submitted(upload_id: int) -> bool:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(submitted,0) FROM timesheets WHERE upload_id = ?", (upload_id,))
    row = cur.fetchone()
    conn.close()
    return bool(row and int(row[0]) == 1)


def mark_timesheet_submitted(upload_id: int):
    conn = get_conn()
    cur = conn.cursor()
    submitted_at = utc_now_str(sep=" ", timespec="seconds")
    cur.execute("""
        UPDATE timesheets
           SET submitted = 1,
               submitted_at = ?
         WHERE upload_id = ?
    """, (submitted_at, upload_id))
    conn.commit()
    conn.close()


def upsert_timesheet(upload_id: int, week_start: str, mon, tue, wed, thu, fri, sat, sun, remaining_planned=None, has_draft=0):
    if is_timesheet_submitted(upload_id):
        raise ValueError("Timesheet is submitted and locked. Contact Admin if changes are required.")

    total = float(mon) + float(tue) + float(wed) + float(thu) + float(fri) + float(sat) + float(sun)
    updated_at = utc_now_str(sep=" ", timespec="seconds")

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM timesheets WHERE upload_id = ?", (upload_id,))
    existing = cur.fetchone()

    if existing:
        cur.execute("""
            UPDATE timesheets
               SET week_start = ?,
                   mon = ?, tue = ?, wed = ?, thu = ?, fri = ?, sat = ?, sun = ?,
                   total = ?,
                   remaining_planned = ?,
                   has_draft = ?,
                   updated_at = ?
             WHERE upload_id = ?
        """, (week_start, mon, tue, wed, thu, fri, sat, sun, total, remaining_planned, has_draft, updated_at, upload_id))
    else:
        cur.execute("""
            INSERT INTO timesheets (
                upload_id, week_start, mon, tue, wed, thu, fri, sat, sun, total,
                remaining_planned, has_draft,
                submitted, submitted_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL, ?)
        """, (upload_id, week_start, mon, tue, wed, thu, fri, sat, sun, total, remaining_planned, has_draft, updated_at))

    conn.commit()
    conn.close()
    return total


def delete_upload(upload_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT stored_filename FROM uploads WHERE id = ?", (upload_id,))
    row = cur.fetchone()
    stored_fn = row[0] if row else None

    # Also find VERIFY screenprint to delete
    cur.execute("SELECT id, stored_filename FROM uploads WHERE module='VERIFY' AND original_filename=?", (f"PPM_{upload_id}",))
    v_rows = cur.fetchall()

    cur.execute("DELETE FROM timesheets WHERE upload_id = ?", (upload_id,))
    cur.execute("DELETE FROM uploads WHERE id = ?", (upload_id,))
    
    for vid, _ in v_rows:
        cur.execute("DELETE FROM uploads WHERE id = ?", (vid,))
        
    conn.commit()
    conn.close()

    if stored_fn:
        path = os.path.join(UPLOAD_DIR, stored_fn)
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            logging.exception("Failed to remove uploaded file from disk: %s", path)




def get_ppm_hours_for_user_week(owner_email: str, week_start: str):
    """Return (sun, mon, tue, wed, thu, fri, sat, total, upload_id) for latest PPM timesheet of user+week; else None."""
    owner = (owner_email or '').strip().lower()
    ws = (week_start or '').strip()
    if not owner or not ws:
        return None
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT u.id,
                   COALESCE(t.sun,0), COALESCE(t.mon,0), COALESCE(t.tue,0), COALESCE(t.wed,0),
                   COALESCE(t.thu,0), COALESCE(t.fri,0), COALESCE(t.sat,0),
                   COALESCE(t.total,0)
            FROM uploads u
            JOIN timesheets t ON t.upload_id = u.id
            WHERE LOWER(COALESCE(u.uploaded_by,'')) = LOWER(?)
              AND COALESCE(u.module,'') = 'PPM'
              AND t.week_start = ?
            ORDER BY u.id DESC
            LIMIT 1
        """, (owner, ws))
        row = cur.fetchone()
    finally:
        conn.close()
    if not row:
        return None
    upload_id, sun, mon, tue, wed, thu, fri, sat, total = row
    return (
        float(sun or 0), float(mon or 0), float(tue or 0), float(wed or 0),
        float(thu or 0), float(fri or 0), float(sat or 0), float(total or 0), int(upload_id)
    )
# -------------------------
# Submit-only symmetric match enforcement (PPM <-> NTT)
# -------------------------
def timesheet_hours_equal(a, b, tol=0.0001) -> bool:
    return all(abs(float(x) - float(y)) <= tol for x, y in zip(a, b))


def list_upload_ids_for_user_week(uploaded_by: str, week_start: str) -> list:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT u.id
          FROM uploads u
          JOIN timesheets t ON t.upload_id = u.id
         WHERE LOWER(COALESCE(u.uploaded_by,'')) = LOWER(?)
           AND t.week_start = ?
         ORDER BY u.id
    """, ((uploaded_by or "").strip().lower(), week_start))
    rows = cur.fetchall()
    conn.close()
    return [r[0] for r in rows]


def submit_week_group(upload_id: int, owner_email: str, week_start: str,
                      mon, tue, wed, thu, fri, sat, sun):
    owner = (owner_email or "").strip().lower()
    group_ids = list_upload_ids_for_user_week(owner, week_start)
    if upload_id not in group_ids:
        group_ids.append(upload_id)

    candidate = (sun, mon, tue, wed, thu, fri, sat)
    current_module = get_upload_module(upload_id)

    # All users submit PPM on Wednesday and NTT on Friday.
    # Comparison should only happen for NTT.
    for uid in group_ids:
        row = get_timesheet_values(uid)
        if not row:
            continue
        _, sun2, mon2, tue2, wed2, thu2, fri2, sat2, total2, submitted2, _updated2, _rem2 = row
        existing = (sun2, mon2, tue2, wed2, thu2, fri2, sat2)

        # 1. If we are in NTT module, we MUST match any existing submitted (PPM or NTT) or any existing entered hours.
        if current_module == "NTT":
            if int(submitted2) == 1 and not timesheet_hours_equal(existing, candidate):
                raise ValueError(
                    "Cannot submit NTT: Hours must match the already submitted PPM/NTT timesheet for this week."
                )
            
            if float(total2 or 0) > 0 and not timesheet_hours_equal(existing, candidate):
                raise ValueError(
                    "Cannot submit: Entered hours in PPM and NTT must match for the same week. "
                    "Please make the hours identical in both modules before submitting NTT."
                )

    for uid in group_ids:
        if is_timesheet_submitted(uid):
            continue
        upsert_timesheet(uid, week_start, mon, tue, wed, thu, fri, sat, sun)
        mark_timesheet_submitted(uid)


# -------------------------
# Email reminders

import smtplib
from email.message import EmailMessage

SMTP_HOST = os.environ.get('SMTP_HOST', '').strip()
SMTP_PORT = int(os.environ.get('SMTP_PORT', '587') or '587')
SMTP_USER = os.environ.get('SMTP_USER', '').strip()
SMTP_PASS = os.environ.get('SMTP_PASS', '').strip()
SMTP_FROM = os.environ.get('SMTP_FROM', SMTP_USER).strip()
SMTP_SSL  = os.environ.get('SMTP_SSL', '').strip().lower() in ('1','true','yes')


def send_email_smtp(to_addr: str, subject: str, body: str) -> None:
    """Send email via SMTP. Success means SMTP server accepted the message."""
    if not SMTP_HOST or not SMTP_FROM:
        raise RuntimeError('SMTP not configured. Set SMTP_HOST and SMTP_FROM (and optionally SMTP_USER/SMTP_PASS).')

    msg = EmailMessage()
    msg['From'] = SMTP_FROM
    msg['To'] = (to_addr or '').strip()
    msg['Subject'] = subject or ''
    msg.set_content(body or '')

    if SMTP_SSL or SMTP_PORT == 465:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            if SMTP_USER and SMTP_PASS:
                s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            try:
                s.ehlo(); s.starttls(); s.ehlo()
            except Exception:
                pass
            if SMTP_USER and SMTP_PASS:
                s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)


def mark_email_reminder_sent(reminder_id: int, ok: bool, error_msg: str = ''):
    conn = get_conn()
    cur = conn.cursor()
    sent_at = utc_now_str(sep=' ', timespec='seconds') if ok else None
    cur.execute(
        "UPDATE email_reminders SET sent=?, sent_at=?, last_error=? WHERE id=?",
        (1 if ok else 0, sent_at, (None if ok else (error_msg or 'Unknown error')), int(reminder_id))
    )
    conn.commit()
    conn.close()


def _due_email_reminders(now_ist: datetime.datetime):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, recipient, subject, message, scheduled_date, scheduled_time FROM email_reminders WHERE COALESCE(sent,0)=0")
    rows = cur.fetchall()
    conn.close()
    due = []
    for rid, recip, subj, msg, sdate, stime in rows:
        try:
            dt_local = datetime.datetime.fromisoformat(f"{sdate} {stime}").replace(tzinfo=IST)
        except Exception:
            continue
        if dt_local <= now_ist:
            due.append((rid, recip, subj, msg))
    return due


def email_scheduler_loop(stop_flag):
    import time
    while not getattr(stop_flag, 'stop', False):
        try:
            now_ist = datetime.datetime.now(tz=IST)
            for rid, recip, subj, msg in _due_email_reminders(now_ist):
                try:
                    send_email_smtp(recip, subj, msg)
                    mark_email_reminder_sent(rid, True, '')
                except Exception as ex:
                    mark_email_reminder_sent(rid, False, str(ex))
        except Exception:
            pass
        time.sleep(30)

# -------------------------


def add_email_reminder(recipient, subject, message, scheduled_date, scheduled_time, created_by=None):
    conn = get_conn()
    cur = conn.cursor()
    created_at = utc_now_str(sep=" ", timespec="seconds")
    cur.execute("""
    INSERT INTO email_reminders (
      recipient, subject, message, scheduled_date, scheduled_time,
      sent, sent_at, last_error, created_by, created_at
    )
    VALUES (?, ?, ?, ?, ?, 0, NULL, NULL, ?, ?)
    """, (
        recipient.strip(),
        subject.strip(),
        message.strip(),
        scheduled_date.strip(),
        (scheduled_time or '09:00').strip(),
        created_by,
        created_at
    ))
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid


def list_email_reminders():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
    SELECT id, recipient, subject, message, scheduled_date, scheduled_time,
           COALESCE(sent,0), COALESCE(sent_at,''), COALESCE(last_error,''),
           created_by, created_at
    FROM email_reminders
    ORDER BY id DESC
    """)
    rows = cur.fetchall()
    conn.close()
    return rows


def get_email_reminder(reminder_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, recipient, subject, message, scheduled_date, scheduled_time, COALESCE(sent,0), COALESCE(sent_at,''), COALESCE(last_error,''), created_by, created_at FROM email_reminders WHERE id = ?",
        (int(reminder_id),)
    )
    row = cur.fetchone()
    conn.close()
    return row


def update_email_reminder(reminder_id: int, recipient: str, subject: str, message: str, scheduled_date: str, scheduled_time: str):
    conn = get_conn()
    cur = conn.cursor()
    # Reset sent flags on update
    cur.execute(
        "UPDATE email_reminders SET recipient=?, subject=?, message=?, scheduled_date=?, scheduled_time=?, sent=0, sent_at=NULL, last_error=NULL WHERE id=?",
        (recipient.strip(), subject.strip(), message.strip(), scheduled_date.strip(), (scheduled_time or '09:00').strip(), int(reminder_id))
    )
    conn.commit()
    conn.close()


def delete_email_reminder(reminder_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM email_reminders WHERE id = ?", (int(reminder_id),))
    conn.commit()
    conn.close()
# -------------------------
# Export helpers (Admin/Manager)
# -------------------------

def _month_start_end(month_str: str):
    if not month_str or not re.match(r'^\d{4}-\d{2}$', month_str):
        return None, None
    y, m = month_str.split('-', 1)
    year = int(y)
    month = int(m)
    first = datetime.date(year, month, 1)
    if month == 12:
        nxt = datetime.date(year + 1, 1, 1)
    else:
        nxt = datetime.date(year, month + 1, 1)
    last = nxt - datetime.timedelta(days=1)
    return first, last

def _week_overlaps_month(week_start_iso: str, first_day: datetime.date, last_day: datetime.date) -> bool:
    try:
        ws = datetime.date.fromisoformat(week_start_iso)
    except Exception:
        return False
    we = ws + datetime.timedelta(days=6)
    return (ws <= last_day) and (we >= first_day)






def _submission_cutoff_color(week_start_iso: str, submitted_at: str, module: str):
    """Return a color indicating timeliness of submission relative to cutoff.

    - 'RED'    : submitted after cutoff day
    - 'YELLOW' : on cutoff day after 3:00 PM IST
    - None     : on/before cutoff (timely) or invalid inputs

    module: 'PPM' => cutoff = Wednesday (Sun+3)
            'NTT' => cutoff = Friday    (Sun+5)
    """
    try:
        ws = datetime.date.fromisoformat(week_start_iso)
    except Exception:
        return None

    try:
        sub_dt = datetime.datetime.fromisoformat(str(submitted_at))
    except Exception:
        return None
    if sub_dt.tzinfo is None:
        sub_dt = sub_dt.replace(tzinfo=UTC)
    sub_dt_ist = sub_dt.astimezone(IST)

    mod = (module or '').strip().upper()
    if mod == 'PPM':
        cutoff_date = ws + datetime.timedelta(days=3)  # Wednesday (Sun+3)
    elif mod == 'NTT':
        cutoff_date = ws + datetime.timedelta(days=5)  # Friday (Sun+5)
    else:
        return None

    if sub_dt_ist.date() > cutoff_date:
        return 'RED'

    if sub_dt_ist.date() == cutoff_date:
        cutoff_dt_ist = datetime.datetime.combine(cutoff_date, datetime.time(15, 0)).replace(tzinfo=IST)
        if sub_dt_ist > cutoff_dt_ist:
            return 'YELLOW'

    return None





def generate_monthly_export_xlsx(month_str: str) -> bytes:
    if Workbook is None:
        raise RuntimeError('openpyxl is not available. Install openpyxl to enable Excel export.')

    first_day, last_day = _month_start_end(month_str)
    if not first_day:
        raise ValueError('Invalid month. Use YYYY-MM.')

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
    SELECT
      LOWER(COALESCE(u.uploaded_by,'')) AS email,
      COALESCE(usr.username, LOWER(COALESCE(u.uploaded_by,''))) AS resource_name,
      COALESCE(u.module,'') AS module,
      COALESCE(t.week_start,'') AS week_start,
      MAX(COALESCE(t.total,0)) AS week_total,
      '' AS submitted_at
    FROM uploads u
    JOIN timesheets t ON t.upload_id = u.id
    LEFT JOIN users usr ON LOWER(COALESCE(usr.email,'')) = LOWER(COALESCE(u.uploaded_by,''))
    WHERE COALESCE(u.module,'') IN ('PPM','NTT')
    GROUP BY LOWER(COALESCE(u.uploaded_by,'')),
             COALESCE(usr.username, LOWER(COALESCE(u.uploaded_by,''))),
             COALESCE(u.module,''),
             COALESCE(t.week_start,'')
    """)
    rows = cur.fetchall()
    conn.close()

    week_starts = sorted({
        r[3] for r in rows
        if r and r[3] and _week_overlaps_month(r[3], first_day, last_day)
    })

    # Bucket: take first 4 week-starts (Sunday) that overlap the month; any extra goes into Week4
    buckets = []
    for ws in week_starts:
        if len(buckets) < 4:
            buckets.append(ws)

    def week_bucket(ws: str):
        if ws in buckets:
            return buckets.index(ws) + 1
        if len(buckets) >= 4 and ws in week_starts and ws not in buckets:
            return 4
        return None

    data = {}    # resource -> wk -> {'PPM': hours, 'NTT': hours}

    for email, rname, module, ws, total, sub_at in rows:
        if not ws or not _week_overlaps_month(ws, first_day, last_day):
            continue
        wk = week_bucket(ws)
        if wk is None:
            continue
        mod = (module or '').strip().upper()
        if mod not in ('PPM', 'NTT'):
            continue
        rkey = (email or '').strip()
        if not rkey:
            continue

        data.setdefault(rkey, {})
        data[rkey].setdefault(wk, {'PPM': 0.0, 'NTT': 0.0})

        data[rkey][wk][mod] += float(total or 0.0)

    wb = Workbook()
    sh = wb.active
    sh.title = month_str

    # Header layout: Email Address | Week1 (PPM, NTT) | ... | Week4 (PPM, NTT) | Total (NTT only)
    sh.cell(row=1, column=1, value='Email Address')
    sh.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    week_cols = {1: 2, 2: 4, 3: 6, 4: 8}
    for wk, c0 in week_cols.items():
        header_value = f'Week{wk}'
        if wk - 1 < len(buckets):
            dt_str = buckets[wk - 1]
        elif len(buckets) > 0:
            try:
                dt = datetime.date.fromisoformat(buckets[-1]) + datetime.timedelta(days=7 * (wk - len(buckets)))
                dt_str = dt.isoformat()
            except:
                dt_str = None
        else:
            dt_str = None
        
        if dt_str:
            try:
                dt = datetime.date.fromisoformat(dt_str)
                dt_fmt = "%b %d '%y"
                header_value = f"{dt.strftime(dt_fmt)} - {(dt + datetime.timedelta(days=6)).strftime(dt_fmt)}"
            except:
                pass
                
        sh.cell(row=1, column=c0, value=header_value)
        sh.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0+1)
        sh.cell(row=2, column=c0, value='PPM')
        sh.cell(row=2, column=c0+1, value='NTT')

    # Total column (NTT only)
    sh.cell(row=1, column=10, value='Total')
    sh.merge_cells(start_row=1, start_column=10, end_row=2, end_column=10)

    # Styles
    thin = Side(style='thin', color='000000') if Side is not None else None
    border = Border(left=thin, right=thin, top=thin, bottom=thin) if Border is not None and thin is not None else None

    # Heading fill: Orange Accent 6, 60% lighter (approx)
    header_fill = PatternFill('solid', fgColor='FCD5B5') if PatternFill is not None else None

    # Header styles (rows 1-2)
    for r in (1, 2):
        for c in range(1, 11):
            cell = sh.cell(row=r, column=c)
            if Font is not None:
                cell.font = Font(bold=True)
            if Alignment is not None:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if border is not None:
                if c in (1, 10):
                    if r == 1:
                        cell.border = Border(left=thin, right=thin, top=thin, bottom=Side(style=None))
                    else:
                        cell.border = Border(left=thin, right=thin, top=Side(style=None), bottom=thin)
                else:
                    cell.border = border
            if header_fill is not None:
                cell.fill = header_fill

    sh.row_dimensions[1].height = 22
    sh.row_dimensions[2].height = 20

    r_out = 3
    for rname in sorted(data.keys(), key=lambda x: x.lower()):
        sh.cell(row=r_out, column=1, value=rname)

        col = 2
        for wk in range(1, 5):
            ppm = data.get(rname, {}).get(wk, {}).get('PPM', 0.0)
            ntt = data.get(rname, {}).get(wk, {}).get('NTT', 0.0)

            ppm_cell = sh.cell(row=r_out, column=col, value=round(ppm, 2))
            ntt_cell = sh.cell(row=r_out, column=col+1, value=round(ntt, 2))
            col += 2

        # Total (NTT only) across Week1-Week4
        ntt_total = sum(data.get(rname, {}).get(wk2, {}).get('NTT', 0.0) for wk2 in range(1, 5))
        sh.cell(row=r_out, column=10, value=round(ntt_total, 2))

        # Row styles
        for c in range(1, 11):
            cell = sh.cell(row=r_out, column=c)
            if Alignment is not None:
                cell.alignment = Alignment(horizontal='left', vertical='center') if c == 1 else Alignment(horizontal='center', vertical='center')
            if border is not None:
                cell.border = border

        r_out += 1

    sh.freeze_panes = 'B3'
    sh.column_dimensions['A'].width = 22
    for letter in ['B','C','D','E','F','G','H','I','J']:
        sh.column_dimensions[letter].width = 12

    from io import BytesIO
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()




# -------------------------
# Export Preview (Detailed) helpers
# -------------------------

def build_monthly_export_detail_rows(month_str: str):
    """Return row-level (per-upload) records for detailed export preview for the selected month.

    Rows are included if their week_start overlaps the selected month (Sun–Sat week range).
    """
    first_day, last_day = _month_start_end(month_str)
    if not first_day:
        raise ValueError('Invalid month. Use YYYY-MM.')

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT
          u.id AS upload_id,
          LOWER(COALESCE(u.uploaded_by,'')) AS email,
          COALESCE(usr.username, LOWER(COALESCE(u.uploaded_by,''))) AS resource_name,
          COALESCE(u.module,'') AS module,
          COALESCE(u.uploaded_at,'') AS uploaded_at,
          COALESCE(t.week_start,'') AS week_start,
          COALESCE(t.sun,0), COALESCE(t.mon,0), COALESCE(t.tue,0), COALESCE(t.wed,0),
          COALESCE(t.thu,0), COALESCE(t.fri,0), COALESCE(t.sat,0),
          COALESCE(t.total,0) AS total,
          COALESCE(t.submitted,0) AS submitted,
          COALESCE(t.submitted_at,'') AS submitted_at
        FROM uploads u
        JOIN timesheets t ON t.upload_id = u.id
        LEFT JOIN users usr ON LOWER(COALESCE(usr.email,'')) = LOWER(COALESCE(u.uploaded_by,''))
        WHERE COALESCE(u.module,'') IN ('PPM','NTT')
        ORDER BY u.id DESC
    """)
    raw = cur.fetchall()
    conn.close()

    out = []
    for r in raw:
        week_start = (r[5] or '').strip()
        if not week_start:
            continue
        if not _week_overlaps_month(week_start, first_day, last_day):
            continue
        out.append({
            'upload_id': int(r[0]),
            'email': (r[1] or '').strip(),
            'resource_name': (r[2] or '').strip(),
            'module': (r[3] or '').strip().upper(),
            'uploaded_at': r[4] or '',
            'week_start': r[5] or '',
            'sun': float(r[6] or 0), 'mon': float(r[7] or 0), 'tue': float(r[8] or 0), 'wed': float(r[9] or 0),
            'thu': float(r[10] or 0), 'fri': float(r[11] or 0), 'sat': float(r[12] or 0),
            'total': float(r[13] or 0),
            'submitted': int(r[14] or 0),
            'submitted_at': r[15] or '',
        })
    return out




def build_export_details_table_html(from_date: str, to_date: str, rows, page: int = 1, page_size: int = 200, show_days: bool = True,
                                  base_path: str = '/export', extra_qs: dict = None):
    """Return ONLY the detailed preview table section (date range mode)."""
    try:
        page = int(page or 1)
    except Exception:
        page = 1
    page = max(1, page)
    try:
        page_size = int(page_size or 200)
    except Exception:
        page_size = 200
    page_size = max(50, min(1000, page_size))

    total_rows = len(rows or [])
    start = (page - 1) * page_size
    end = start + page_size
    page_rows = (rows or [])[start:end]

    def q(url_base, **params):
        params = dict(params or {})
        if extra_qs:
            for k, v in (extra_qs or {}).items():
                if v is None or str(v).strip() == '':
                    continue
                params.setdefault(k, v)
        params.setdefault('from_date', from_date)
        params.setdefault('to_date', to_date)
        parts = [f"{k}={quote(str(v))}" for k, v in params.items() if v is not None]
        return url_base + ('?' + '&'.join(parts) if parts else '')

    prev_link = q(base_path, page=page-1, ps=page_size, days=('1' if show_days else '0')) if page > 1 else ''
    next_link = q(base_path, page=page+1, ps=page_size, days=('1' if show_days else '0')) if end < total_rows else ''
    toggle_days = q(base_path, page=1, ps=page_size, days=('0' if show_days else '1'))

    day_cols_head = ''
    if show_days:
        day_cols_head = '<th>Sun</th><th>Mon</th><th>Tue</th><th>Wed</th><th>Thu</th><th>Fri</th><th>Sat</th>'

    body = []
    for r in page_rows:
        locked_badge = "<span class='badge lock-badge'>Submitted</span>" if int(r.get('submitted') or 0) == 1 else "<span class='badge'>Draft</span>"
        days_html = ''
        if show_days:
            days_html = (
                f"<td style='text-align:left'>{round(float(r.get('sun') or 0),2)}</td>"
                f"<td style='text-align:left'>{round(float(r.get('mon') or 0),2)}</td>"
                f"<td style='text-align:left'>{round(float(r.get('tue') or 0),2)}</td>"
                f"<td style='text-align:left'>{round(float(r.get('wed') or 0),2)}</td>"
                f"<td style='text-align:left'>{round(float(r.get('thu') or 0),2)}</td>"
                f"<td style='text-align:left'>{round(float(r.get('fri') or 0),2)}</td>"
                f"<td style='text-align:left'>{round(float(r.get('sat') or 0),2)}</td>"
            )
        body.append(f"""
        <tr>
          <td style='text-align:left'>{int(r.get('upload_id') or 0)}</td>
          <td style='text-align:left'>
            <div style='font-weight:600;line-height:1.2'>{html.escape((r.get('resource_name') or r.get('email') or '').strip())}</div>
            <div class='muted' style='font-size:11px'>{html.escape((r.get('email') or '').strip())}</div>
          </td>
          <td style='text-align:left'>{html.escape((r.get('module') or '').strip())}</td>
          <td style='text-align:left;font-size:12px;color:#64748b'>{html.escape(format_dt_ist(r.get('uploaded_at') or ''))}</td>
          <td style='text-align:left'>{html.escape((r.get('week_start') or '').strip())}</td>
          {days_html}
          <td style='text-align:left;font-weight:800;color:var(--accent)'>{round(float(r.get('total') or 0),2)}</td>
          <td style='text-align:left'>{locked_badge}</td>
          <td style='text-align:left;font-size:12px;color:#64748b'>{html.escape(format_dt_ist(r.get('submitted_at') or '')) if (r.get('submitted_at') or '').strip() else ''}</td>
        </tr>
        """)

    if not body:
        colspan = 9 + (7 if show_days else 0)
        body_html = f"<tr><td colspan='{colspan}' class='muted' style='padding:12px'>No rows found for this date range.</td></tr>"
        tfoot_html = ""
    else:
        body_html = ''.join(body)
        grand_total = sum(float((row_.get('total') or 0)) for row_ in (rows or []))
        cfmt = 5 + (7 if show_days else 0)
        tfoot_html = f"<tfoot style='background:#f1f5f9;border-top:2px solid #cbd5e1;font-weight:800;color:#0f1724'><tr><td colspan='{cfmt}' style='text-align:left;padding:12px 14px'>Grand Total:</td><td style='text-align:left;padding:12px 14px'>{round(grand_total, 2)}</td><td colspan='2'></td></tr></tfoot>"

    page_from = (start + 1) if total_rows else 0
    page_to = min(end, total_rows)

    return f"""
      <div class='muted' style='margin-top:10px'>Rows: <b>{total_rows}</b> • Showing: <b>{page_from}-{page_to}</b></div>
      <div class='controls' style='margin-top:12px;justify-content:space-between'>
        <div style='display:flex;gap:10px;flex-wrap:wrap'>
          <a class='btn outline small' href='{toggle_days}'><i class='fa fa-calendar'></i> {'Hide Daily Hours' if show_days else 'Show Daily Hours'}</a>
        </div>
        <div style='display:flex;gap:10px;align-items:center;flex-wrap:wrap'>
          {f"<a class='btn outline small' href='{prev_link}'><i class='fa fa-chevron-left'></i> Prev</a>" if prev_link else ''}
          {f"<a class='btn outline small' href='{next_link}'>Next <i class='fa fa-chevron-right'></i></a>" if next_link else ''}
        </div>
      </div>
      <div style='overflow-x:auto; overflow-y:hidden; margin-top:12px; border:1px solid #e2e8f0; border-radius:12px; background:#fff;'>
        <table class='um-table' style='width:100%;border-collapse:collapse'>
          <thead style='background:#f8fafc'>
            <tr style='border-bottom:1px solid #e2e8f0'>
              <th>Upload ID</th>
              <th>Resource</th>
              <th>Module</th>
              <th>Uploaded At (IST)</th>
              <th>Week Start</th>
              {day_cols_head}
              <th>Total</th>
              <th>Status</th>
              <th>Submitted At (IST)</th>
            </tr>
          </thead>
          <tbody>
            {body_html}
          </tbody>
          {tfoot_html}
        </table>
      </div>
    """


def _qs_first(qs: dict, key: str, default: str = "") -> str:
    """parse_qs helper: get first value"""
    try:
        v = (qs.get(key, [default]) or [default])[0]
    except Exception:
        v = default
    return (v or default).strip()


def _parse_op_value(raw: str):
    """Parse simple operator expressions for NUMBER/DATE custom field filters.

    Supported:
      >=10, <=5, >3, <7
      10..20  (between)
      plain   (treated as '=' or 'contains' depending on field type)

    Returns: (op, v1, v2)
    """
    s = (raw or "").strip()
    if not s:
        return (None, None, None)

    if ".." in s:
        a, b = s.split("..", 1)
        return ("between", a.strip(), b.strip())

    for op in (">=", "<=", ">", "<"):
        if s.startswith(op):
            return (op, s[len(op):].strip(), None)

    return ("=", s, None)




def parse_export_filters_from_qs(qs: dict):
    """Parse Export filters from a parse_qs() dict.

    Supported:
    - from_date (YYYY-MM-DD)
    - to_date (YYYY-MM-DD)
    - resource (name/email)
    - u_role (role)
    - u_status (status)
    - custom fields (cf_<id>)

    Returns:
      filters: dict
      persist_qs: dict for pagination/toggles
    """
    from_date = _qs_first(qs, 'from_date', '')
    to_date = _qs_first(qs, 'to_date', '')
    resource = _qs_first(qs, 'resource', '')

    filters = {
        'from_date': from_date,
        'to_date': to_date,
        'resource': resource,
        'u_role': _qs_first(qs, 'u_role', ''),
        'u_status': _qs_first(qs, 'u_status', ''),
        'cf': {},
    }

    for k in (qs or {}).keys():
        if not k.startswith('cf_'):
            continue
        suf = k[3:]
        if not suf.isdigit():
            continue
        fid = int(suf)
        raw = _qs_first(qs, k, '')
        if not raw:
            continue
        op, v1, v2 = _parse_op_value(raw)
        filters['cf'][fid] = {'raw': raw, 'op': op, 'v1': v1, 'v2': v2}

    persist_qs = {}
    for key in ('from_date','to_date','resource','u_role','u_status'):
        v = (filters.get(key) or '').strip()
        if v:
            persist_qs[key] = v

    for fid, meta in (filters.get('cf') or {}).items():
        raw = (meta.get('raw') or '').strip()
        if raw:
            persist_qs[f'cf_{int(fid)}'] = raw

    return filters, persist_qs


def build_monthly_export_detail_rows_filtered(month_str: str, filters: dict):
    """Return per-upload export rows filtered by uploader profile + custom fields.

    Month logic remains: include rows whose week_start overlaps selected month.
    """
    fd = (filters.get('from_date') or '').strip()
    td = (filters.get('to_date') or '').strip()
    try:
        from_d = datetime.date.fromisoformat(fd)
        to_d = datetime.date.fromisoformat(td)
    except Exception:
        raise ValueError('Invalid date range. Use YYYY-MM-DD.')
    if to_d < from_d:
        raise ValueError('Invalid date range: To Date cannot be earlier than From Date.')

    # Map field_id -> definition row
    try:
        cf_defs = list_custom_fields(active_only=False)
    except Exception:
        cf_defs = []
    field_by_id = {int(r[0]): r for r in (cf_defs or [])}

    where = ["COALESCE(u.module,'') IN ('PPM','NTT')"]
    params = []
    where.append("substr(COALESCE(u.uploaded_at,''), 1, 10) BETWEEN ? AND ?")
    params.extend([from_d.isoformat(), to_d.isoformat()])

    # Resource filter (username/email)
    resource = (filters.get('resource') or '').strip()
    if resource:
        like = f"%{resource.lower()}%"
        where.append("(LOWER(COALESCE(usr.username,'')) LIKE ? OR LOWER(COALESCE(u.uploaded_by,'')) LIKE ?)")
        params.extend([like, like])

    # Uploader built-in fields
    u_role = (filters.get('u_role') or '').strip()
    if u_role:
        where.append("COALESCE(usr.role,'') = ?")
        params.append(u_role)

    u_status = (filters.get('u_status') or '').strip()
    if u_status:
        where.append("COALESCE(usr.status,'') = ?")
        params.append(u_status)

    u_sa = (filters.get('u_screen_access') or '').strip().upper()
    if u_sa:
        where.append("UPPER(COALESCE(usr.screen_access,'BOTH')) = ?")
        params.append(u_sa)

    for flag in ('u_can_export', 'u_can_email', 'u_can_help'):
        v = (filters.get(flag) or '').strip()
        if v in ('0', '1'):
            col = flag.replace('u_', '')
            where.append(f"COALESCE(usr.{col},0) = ?")
            params.append(int(v))

    # Custom field filters (EXISTS)
    cf = filters.get('cf') or {}
    for fid, meta in cf.items():
        try:
            fid = int(fid)
        except Exception:
            continue
        def_row = field_by_id.get(fid)
        if not def_row:
            continue
        _id, _key, _label, ftype, _req, opt_json, _active = def_row
        ftype = (ftype or 'TEXT').strip().upper()

        op = meta.get('op')
        v1 = (meta.get('v1') or '').strip()
        v2 = (meta.get('v2') or '').strip()
        if not v1 and op != 'between':
            continue

        if ftype in ('DROPDOWN', 'BOOLEAN'):
            where.append("""
                EXISTS (
                    SELECT 1 FROM user_custom_field_values v
                    WHERE v.user_id = usr.id
                      AND v.field_id = ?
                      AND COALESCE(v.value_text,'') = ?
                )
            """)
            params.extend([fid, v1])

        elif ftype == 'NUMBER':
            if op == 'between':
                where.append("""
                    EXISTS (
                        SELECT 1 FROM user_custom_field_values v
                        WHERE v.user_id = usr.id
                          AND v.field_id = ?
                          AND CAST(COALESCE(v.value_text,'0') AS REAL) BETWEEN CAST(? AS REAL) AND CAST(? AS REAL)
                    )
                """)
                params.extend([fid, v1, v2])
            elif op in ('>=', '<=', '>', '<'):
                where.append(f"""
                    EXISTS (
                        SELECT 1 FROM user_custom_field_values v
                        WHERE v.user_id = usr.id
                          AND v.field_id = ?
                          AND CAST(COALESCE(v.value_text,'0') AS REAL) {op} CAST(? AS REAL)
                    )
                """)
                params.extend([fid, v1])
            else:
                where.append("""
                    EXISTS (
                        SELECT 1 FROM user_custom_field_values v
                        WHERE v.user_id = usr.id
                          AND v.field_id = ?
                          AND CAST(COALESCE(v.value_text,'0') AS REAL) = CAST(? AS REAL)
                    )
                """)
                params.extend([fid, v1])

        elif ftype == 'DATE':
            # assume stored as YYYY-MM-DD
            if op == 'between':
                where.append("""
                    EXISTS (
                        SELECT 1 FROM user_custom_field_values v
                        WHERE v.user_id = usr.id
                          AND v.field_id = ?
                          AND COALESCE(v.value_text,'') BETWEEN ? AND ?
                    )
                """)
                params.extend([fid, v1, v2])
            elif op in ('>=', '<=', '>', '<'):
                where.append(f"""
                    EXISTS (
                        SELECT 1 FROM user_custom_field_values v
                        WHERE v.user_id = usr.id
                          AND v.field_id = ?
                          AND COALESCE(v.value_text,'') {op} ?
                    )
                """)
                params.extend([fid, v1])
            else:
                where.append("""
                    EXISTS (
                        SELECT 1 FROM user_custom_field_values v
                        WHERE v.user_id = usr.id
                          AND v.field_id = ?
                          AND COALESCE(v.value_text,'') = ?
                    )
                """)
                params.extend([fid, v1])

        else:
            # TEXT default = contains
            like = f"%{v1.lower()}%"
            where.append("""
                EXISTS (
                    SELECT 1 FROM user_custom_field_values v
                    WHERE v.user_id = usr.id
                      AND v.field_id = ?
                      AND LOWER(COALESCE(v.value_text,'')) LIKE ?
                )
            """)
            params.extend([fid, like])

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(f"""
        SELECT
            u.id AS upload_id,
            LOWER(COALESCE(u.uploaded_by,'')) AS email,
            COALESCE(usr.username, LOWER(COALESCE(u.uploaded_by,''))) AS resource_name,
            COALESCE(u.module,'') AS module,
            COALESCE(u.uploaded_at,'') AS uploaded_at,
            COALESCE(t.week_start,'') AS week_start,
            COALESCE(t.sun,0), COALESCE(t.mon,0), COALESCE(t.tue,0), COALESCE(t.wed,0),
            COALESCE(t.thu,0), COALESCE(t.fri,0), COALESCE(t.sat,0),
            COALESCE(t.total,0) AS total,
            COALESCE(t.submitted,0) AS submitted,
            COALESCE(t.submitted_at,'') AS submitted_at
        FROM uploads u
        JOIN timesheets t ON t.upload_id = u.id
        LEFT JOIN users usr ON LOWER(COALESCE(usr.email,'')) = LOWER(COALESCE(u.uploaded_by,''))
        WHERE {" AND ".join(where)}
        ORDER BY u.id DESC
    """, tuple(params))
    raw = cur.fetchall()
    conn.close()
    out = []
    for r in raw:
        em = (r[1] or '').strip()
        ws = r[5] or ''
        mod = (r[3] or '').strip().upper()
        
        out.append({
            'upload_id': int(r[0]),
            'email': em,
            'resource_name': (r[2] or '').strip(),
            'module': mod,
            'uploaded_at': r[4] or '',
            'week_start': ws,
            'sun': float(r[6] or 0), 'mon': float(r[7] or 0), 'tue': float(r[8] or 0), 'wed': float(r[9] or 0),
            'thu': float(r[10] or 0), 'fri': float(r[11] or 0), 'sat': float(r[12] or 0),
            'total': float(r[13] or 0),
            'submitted': int(r[14] or 0),
            'submitted_at': r[15] or '',
        })
    return out


def generate_monthly_export_xlsx_filtered(month_str: str, filters: dict) -> bytes:
    """Generate monthly export Excel respecting uploader-profile filters.

    Output format matches generate_monthly_export_xlsx().
    """
    if Workbook is None:
        raise RuntimeError('openpyxl is not available. Install openpyxl to enable Excel export.')

    fd = (filters.get('from_date') or '').strip()
    td = (filters.get('to_date') or '').strip()
    if not fd or not td:
        raise ValueError('Invalid date range. Provide from_date and to_date (YYYY-MM-DD).')
    try:
        datetime.date.fromisoformat(fd); datetime.date.fromisoformat(td)
    except Exception:
        raise ValueError('Invalid date range. Use YYYY-MM-DD.')

    rows = build_monthly_export_detail_rows_filtered(month_str, filters)

    week_starts = sorted({r.get('week_start','') for r in rows if (r.get('week_start','') or '').strip()})

    buckets = []
    for ws in week_starts:
        if len(buckets) < 4:
            buckets.append(ws)

    def week_bucket(ws: str):
        if ws in buckets:
            return buckets.index(ws) + 1
        if len(buckets) >= 4 and ws in week_starts and ws not in buckets:
            return 4
        return None

    data = {}
    for r in rows:
        ws = (r.get('week_start') or '').strip()
        if not ws:
            continue
        wk = week_bucket(ws)
        if wk is None:
            continue
        mod = (r.get('module') or '').strip().upper()
        if mod not in ('PPM','NTT'):
            continue
        rkey = (r.get('email') or '').strip()
        if not rkey:
            continue
        data.setdefault(rkey, {})
        data[rkey].setdefault(wk, {'PPM': 0.0, 'NTT': 0.0})
        data[rkey][wk][mod] += float(r.get('total') or 0.0)

    wb = Workbook()
    sh = wb.active
    sh.title = (f'{fd}_to_{td}'[:31])

    sh.cell(row=1, column=1, value='Email Address')
    sh.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    week_cols = {1: 2, 2: 4, 3: 6, 4: 8}
    for wk, c0 in week_cols.items():
        header_value = f'Week{wk}'
        if wk - 1 < len(buckets):
            dt_str = buckets[wk - 1]
        elif len(buckets) > 0:
            try:
                dt = datetime.date.fromisoformat(buckets[-1]) + datetime.timedelta(days=7 * (wk - len(buckets)))
                dt_str = dt.isoformat()
            except:
                dt_str = None
        else:
            dt_str = None
        
        if dt_str:
            try:
                dt = datetime.date.fromisoformat(dt_str)
                dt_fmt = "%b %d '%y"
                header_value = f"{dt.strftime(dt_fmt)} - {(dt + datetime.timedelta(days=6)).strftime(dt_fmt)}"
            except:
                pass
                
        sh.cell(row=1, column=c0, value=header_value)
        sh.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0+1)
        sh.cell(row=2, column=c0, value='PPM')
        sh.cell(row=2, column=c0+1, value='NTT')

    sh.cell(row=1, column=10, value='Total')
    sh.merge_cells(start_row=1, start_column=10, end_row=2, end_column=10)

    thin = Side(style='thin', color='000000') if Side is not None else None
    border = Border(left=thin, right=thin, top=thin, bottom=thin) if Border is not None and thin is not None else None
    header_fill = PatternFill('solid', fgColor='FCD5B5') if PatternFill is not None else None

    for r in (1, 2):
        for c in range(1, 11):
            cell = sh.cell(row=r, column=c)
            if Font is not None:
                cell.font = Font(bold=True)
            if Alignment is not None:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if border is not None:
                if c in (1, 10):
                    if r == 1:
                        cell.border = Border(left=thin, right=thin, top=thin, bottom=Side(style=None))
                    else:
                        cell.border = Border(left=thin, right=thin, top=Side(style=None), bottom=thin)
                else:
                    cell.border = border
            if header_fill is not None:
                cell.fill = header_fill

    sh.row_dimensions[1].height = 22
    sh.row_dimensions[2].height = 20

    r_out = 3
    for rname in sorted(data.keys(), key=lambda x: x.lower()):
        sh.cell(row=r_out, column=1, value=rname)
        col = 2
        for wk in range(1, 5):
            ppm = data.get(rname, {}).get(wk, {}).get('PPM', 0.0)
            ntt = data.get(rname, {}).get(wk, {}).get('NTT', 0.0)
            sh.cell(row=r_out, column=col, value=round(ppm, 2))
            sh.cell(row=r_out, column=col+1, value=round(ntt, 2))
            col += 2
        ntt_total = sum(data.get(rname, {}).get(wk2, {}).get('NTT', 0.0) for wk2 in range(1, 5))
        sh.cell(row=r_out, column=10, value=round(ntt_total, 2))

        for c in range(1, 11):
            cell = sh.cell(row=r_out, column=c)
            if Alignment is not None:
                cell.alignment = Alignment(horizontal='left', vertical='center') if c == 1 else Alignment(horizontal='center', vertical='center')
            if border is not None:
                cell.border = border
        r_out += 1

    sh.freeze_panes = 'B3'
    sh.column_dimensions['A'].width = 22
    for letter in ['B','C','D','E','F','G','H','I','J']:
        sh.column_dimensions[letter].width = 12

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()




def build_export_filter_ui_html(filters: dict, ps: int = 200, show_days: bool = True):
    """Build export filter UI.

    DATE RANGE MODE (requested):
    - Filter by Uploaded Date (from_date .. to_date), inclusive.
    - Built-in fields Resource/Role/Status are optional and can be added/removed (show/hide) via +/–.
    - Removed fields from Export UI: Screen Access, Can Export, Can Email, Can Help.
    - Custom field filters keep the +/– behavior.

    Notes:
    - GET /export applies filters and refreshes preview.
    - POST /export downloads Excel with current filter values.
    """

    from_date = html.escape((filters.get('from_date') or ''), quote=True)
    to_date = html.escape((filters.get('to_date') or ''), quote=True)

    resource = html.escape((filters.get('resource') or ''), quote=True)
    u_role = (filters.get('u_role') or '').strip()
    u_status = (filters.get('u_status') or '').strip()

    # -----------------------------
    # Custom fields (Add/Remove with + / –)
    # -----------------------------
    cf_rows = list_custom_fields(active_only=True)
    cf_map = filters.get('cf') or {}
    cf_html = []
    picker_options = []

    for fid, fkey, label, ftype, req, opt_json, active in cf_rows:
        fid = int(fid)
        ftype = (ftype or 'TEXT').upper()
        val = (cf_map.get(fid, {}).get('raw') or '').strip()
        safe_label = html.escape((label or fkey or ''), quote=False)
        safe_val = html.escape(val, quote=True)
        hidden_attr = '' if val else " data-hidden='1' style='display:none'"

        if ftype == 'DROPDOWN':
            opts = _parse_options_json(opt_json)
            opt_tags = ["<option value=''>-- Any --</option>"]
            for o in opts:
                sel = ' selected' if o == val else ''
                opt_tags.append(f"<option value='{html.escape(o, quote=True)}'{sel}>{html.escape(o)}</option>")
            control = "<select name='cf_%d' class='cf-input'>%s</select>" % (fid, ''.join(opt_tags))
        elif ftype == 'BOOLEAN':
            opt_tags = [
                "<option value=''>-- Any --</option>",
                f"<option value='0' {'selected' if val=='0' else ''}>No</option>",
                f"<option value='1' {'selected' if val=='1' else ''}>Yes</option>",
            ]
            control = "<select name='cf_%d' class='cf-input'>%s</select>" % (fid, ''.join(opt_tags))
        elif ftype in ('NUMBER', 'DATE'):
            hint = "Use >=x, <=x, x..y" if ftype == 'NUMBER' else "Use >=YYYY-MM-DD, <=YYYY-MM-DD, YYYY-MM-DD..YYYY-MM-DD"
            control = f"<input name='cf_{fid}' class='cf-input' value='{safe_val}' placeholder='{html.escape(hint, quote=True)}'>"
        else:
            control = f"<input name='cf_{fid}' class='cf-input' value='{safe_val}' placeholder='Contains...'>"

        cf_html.append(f"""
        <div class='filter-field grow cf-row' data-fid='{fid}'{hidden_attr}>
          <label>{safe_label} (Custom)</label>
          <div style='display:flex;gap:8px;align-items:center;'>
            <div style='flex:1'>{control}</div>
            <button type='button' class='btn outline small cf-remove' title='Remove this filter' aria-label='Remove {safe_label}'>–</button>
          </div>
        </div>
        """)
        picker_options.append({'fid': fid, 'label': safe_label})

    # Built-in add/remove picker (Resource/Role/Status) - Hidden by request
    builtin_add_html = """
      <div class='filter-field' style='display:none;min-width:280px;'>
        <label style='visibility:hidden'>Add Filter</label>
        <div style='display:flex;gap:4px;align-items:center;'>
          <button type='button' id='biAddBtn' class='btn secondary' style='display:none;margin-top:0;height:40px;padding:0 12px;' title='Add this filter'>+ Add Filter</button>
          <input list='biOpts' id='biAddPicker' placeholder='Type or select filter...' style='margin-top:0;height:40px;padding:8px 12px;border-radius:8px;flex:1;background:#fff;border:1px solid #94a3b8;color:#000;'>
          <datalist id='biOpts'>
            <option value='Resource'>
            <option value='Role'>
            <option value='Status'>
          </datalist>
        </div>
      </div>
    """

    # Custom-field add UI
    cf_entries = [ (f"{label} (Custom)" if ftype != 'TEXT' else label) for fid, fkey, label, ftype, req, opt_json, active in cf_rows ]
    add_filter_html = f"""
      <div class='filter-field' style='min-width:280px;'>
        <label style='visibility:hidden'>Add Custom</label>
        <div style='display:flex;gap:4px;align-items:center;'>
          <button type='button' id='cfAddBtn' class='btn secondary' style='margin-top:0;height:40px;padding:0 12px;' title='Add custom field'>+ Add Custom</button>
          <input list='cfOpts' id='cfAddPicker' placeholder='Type custom field name...' style='margin-top:0;height:40px;padding:8px 12px;border-radius:8px;flex:1;background:#fff;border:1px solid #94a3b8;color:#000;'>
          <datalist id='cfOpts'>
            {"".join([f"<option value='{html.escape(opt, quote=True)}'>" for opt in cf_entries])}
          </datalist>
        </div>
      </div>
    """

    # Hidden inputs for download form
    dl_hidden = "".join([
        f"<input type='hidden' name='{html.escape(k, quote=True)}' value=''>"
        for k in ('from_date','to_date','resource','u_role','u_status')
    ])

    picker_json = json.dumps(picker_options, ensure_ascii=False)

    script_tpl = r"""
    <script>
    (function(){
      const pickerData = __PICKER_JSON__;
      const get = (id) => document.getElementById(id);
      const isHid = (r) => (!r || r.style.display === 'none' || r.getAttribute('data-hidden') === '1');

      function setup(){
        const flt = get('fltForm');
        const dl = get('dlForm');
        const cfPicker = get('cfAddPicker');
        const cfAddBtn = get('cfAddBtn');
        const biPicker = get('biAddPicker');
        const biAddBtn = get('biAddBtn');

        if(!flt) return;

        const refreshBI = () => {
          if(!biPicker) return;
          const shown = new Set(Array.from(flt.querySelectorAll('.builtin-row')).filter(r => !isHid(r)).map(r => r.getAttribute('data-key')));
          Array.from(biPicker.options).forEach(opt => { if(opt.value) opt.disabled = shown.has(opt.value); });
        };

        const refreshCF = () => {
          if(!cfPicker) return;
          const prev = cfPicker.value;
          cfPicker.innerHTML = "<option value=''>-- Choose a custom field --</option>";
          const hiddenFids = new Set(Array.from(flt.querySelectorAll('.cf-row')).filter(isHid).map(r => r.getAttribute('data-fid')));
          pickerData.forEach(opt => {
            if(hiddenFids.has(String(opt.fid))){
              const o = document.createElement('option');
              o.value = String(opt.fid); o.textContent = opt.label;
              cfPicker.appendChild(o);
            }
          });
          if(Array.from(cfPicker.options).some(o => o.value === prev)) cfPicker.value = prev;
        };

        const showRow = (row, focus = true) => {
          if(!row) return;
          row.style.display = '';
          row.removeAttribute('data-hidden');
          if(focus){ const i = row.querySelector('input,select,textarea'); if(i) i.focus(); }
        };

        if(biAddBtn && biPicker){
          const addBI = () => {
            const v = (biPicker.value || '').trim().toLowerCase();
            if(!v) return;
            let key = (v.indexOf('resource')>=0 ? 'resource' : (v.indexOf('role')>=0 ? 'u_role' : (v.indexOf('status')>=0 ? 'u_status' : '')));
            if(!key) return;
            const row = flt.querySelector(".builtin-row[data-key='" + key + "']");
            if(row){
              row.style.display = '';
              row.removeAttribute('data-hidden');
              const i = row.querySelector('input,select,textarea'); if(i) i.focus();
              biPicker.value = '';
            }
          };
          biAddBtn.onclick = addBI;
          biPicker.oninput = addBI;
        }

        if(cfAddBtn && cfPicker){
          const addCF = () => {
             const v = (cfPicker.value || '').trim();
             if(!v) return;
             const rows = Array.from(flt.querySelectorAll('.cf-row'));
             const row = rows.find(r => {
                const label = (r.querySelector('label')||{}).textContent || '';
                return label.includes(v) || v.includes(label.replace('(Custom)','').trim());
             });
             if(row){
                row.style.display = '';
                row.removeAttribute('data-hidden');
                const i = row.querySelector('input,select,textarea'); if(i) i.focus();
                cfPicker.value = '';
             }
          };
          cfAddBtn.onclick = addCF;
          cfPicker.oninput = addCF;
        }

        flt.onclick = (e) => {
          const btn = e.target.closest('.cf-remove, .builtin-remove');
          if(!btn) return;
          const row = btn.closest('.cf-row, .builtin-row');
          if(!row) return;
          const i = row.querySelector('input,select,textarea'); if(i) i.value = '';
          row.setAttribute('data-hidden', '1'); row.style.display = 'none';
          refreshBI(); refreshCF();
        };

        if(dl){
          dl.onsubmit = () => {
            const els = flt.querySelectorAll('input[name], select[name], textarea[name]');
            els.forEach(el => {
              const n = el.getAttribute('name');
              const r = el.closest('.cf-row, .builtin-row');
              if(!n || (r && isHid(r))) return;
              let h = dl.querySelector('input[name="' + n + '"]');
              if(!h){ h = document.createElement('input'); h.type = 'hidden'; h.name = n; dl.appendChild(h); }
              h.value = el.value || '';
            });
          };
        }

        const syncAll = () => {
          Array.from(flt.querySelectorAll('.cf-row, .builtin-row')).forEach(r => {
            const i = r.querySelector('input,select,textarea');
            if(!i || (i.value || '').trim() === ''){ r.style.display = 'none'; r.setAttribute('data-hidden', '1'); }
            else { showRow(r, false); }
          });
          refreshBI(); refreshCF();
        };
        syncAll();
      }

      if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', setup);
      else setup();
    })();
    </script>
    """

    script = script_tpl.replace('__PICKER_JSON__', picker_json)

    builtin_rows_html = f"""
      <div class='filter-field grow filter-search builtin-row' data-key='resource' data-hidden='1' style='display:none'>
        <label>Resource Name / Email</label>
        <i class='fa fa-search'></i>
        <div style='display:flex;gap:8px;align-items:center;'>
          <input type='text' name='resource' value='{resource}' placeholder='Search uploader name/email...'>
          <button type='button' class='btn outline small builtin-remove' title='Remove this filter' aria-label='Remove Resource filter'>–</button>
        </div>
      </div>

      <div class='filter-field builtin-row' data-key='u_role' data-hidden='1' style='display:none'>
        <label>Role</label>
        <div style='display:flex;gap:8px;align-items:center;'>
          <select name='u_role'>
            <option value=''>-- Any --</option>
            <option value='Admin' {'selected' if u_role=='Admin' else ''}>Admin</option>
            <option value='Manager' {'selected' if u_role=='Manager' else ''}>Manager</option>
            <option value='Employee' {'selected' if u_role=='Employee' else ''}>Employee</option>
          </select>
          <button type='button' class='btn outline small builtin-remove' title='Remove this filter' aria-label='Remove Role filter'>–</button>
        </div>
      </div>

      <div class='filter-field builtin-row' data-key='u_status' data-hidden='1' style='display:none'>
        <label>Status</label>
        <div style='display:flex;gap:8px;align-items:center;'>
          <select name='u_status'>
            <option value=''>-- Any --</option>
            <option value='Active' {'selected' if u_status=='Active' else ''}>Active</option>
            <option value='Inactive' {'selected' if u_status=='Inactive' else ''}>Inactive</option>
          </select>
          <button type='button' class='btn outline small builtin-remove' title='Remove this filter' aria-label='Remove Status filter'>–</button>
        </div>
      </div>
    """

    return f"""
      <div class='filter-card'>
        <div class='filter-title'>
          <p class='muted'><b>Export Filters</b> (Uploaded Date Range + Custom Fields)</p>
        </div>

        <form method='get' action='/export' id='fltForm'>
          <div class='filter-form'>
            <div class='filter-field'>
              <label>From Date</label>
              <input type='date' name='from_date' value='{from_date}' required>
            </div>
            <div class='filter-field'>
              <label>To Date</label>
              <input type='date' name='to_date' value='{to_date}' required>
            </div>

            {builtin_rows_html}

            {builtin_add_html}

            {''.join(cf_html)}

            {add_filter_html}

            <div class='filter-actions' style='margin-top:8px;'>
              <button class='btn small' type='submit'><i class='fa fa-rotate'></i> Refresh Preview</button>
              <a class='btn secondary small' href='/export' style='text-decoration:none;'><i class='fa fa-rotate-left'></i> Reset</a>
            </div>
          </div>

          <input type='hidden' name='page' value='1'>
          <input type='hidden' name='ps' value='{int(ps)}'>
          <input type='hidden' name='days' value='{1 if show_days else 0}'>
        </form>

        <div style='margin-top:10px;display:flex;justify-content:flex-end'>
          <form method='post' action='/export' id='dlForm' style='margin:0'>
            {dl_hidden}
            <button class='btn small' type='submit'><i class='fa fa-file-excel'></i> Download Excel</button>
          </form>
        </div>

        {script}
      </div>
    """


# Start email scheduler background thread
class _StopFlag:
    stop = False

# Global stop flag (used by email_scheduler_loop)
STOP_EMAIL_SCHEDULER = _StopFlag()

def start_email_scheduler():
    import threading
    t = threading.Thread(target=email_scheduler_loop, args=(STOP_EMAIL_SCHEDULER,), daemon=True)
    t.start()
    return t

EMAIL_SCHEDULER_STARTED = start_email_scheduler()

# -------------------------
# UI templates
# -------------------------
PAGE_STYLE = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
:root{--bg:#f4f6f8;--nav:#0f1724;--accent:#2563eb;--card:#ffffff;--muted:#64748b}
*{box-sizing:border-box}
body{margin:0;font-family:'Inter',"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:#1e293b;font-size:14px;line-height:1.5}
header.app-header{position:fixed;top:0;left:0;right:0;height:64px;background:linear-gradient(90deg,var(--nav),#0b1220);color:#fff;display:flex;align-items:center;justify-content:space-between;padding:0 20px;z-index:1000;box-shadow:0 2px 8px rgba(2,6,23,0.15)}
.brand{display:flex;align-items:center;gap:12px;font-weight:700}
.brand .logo{width:36px;height:36px;border-radius:8px;background:linear-gradient(135deg,var(--accent),#7c3aed);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;box-shadow:0 2px 6px rgba(37,99,235,0.2)}
.header-right{display:flex;align-items:center;gap:12px;font-weight:600}
.layout{display:flex;margin-top:64px;min-height:calc(100vh - 64px)}
aside.sidebar{width:200px;background:#0b1220;color:#e6eef8;padding:12px 8px;border-right:1px solid rgba(255,255,255,0.03);position:sticky;top:64px;height:calc(100vh - 64px)}
.nav-item{display:flex;align-items:center;gap:12px;padding:10px 12px;border-radius:8px;color:inherit;text-decoration:none;margin-bottom:6px;font-weight:600;position:relative}
.nav-item i{width:20px;text-align:center;color:#9fb3d8}
.nav-item:hover{background:rgba(255,255,255,0.04);cursor:pointer}
.dropdown-content{display:none;position:absolute;top:100%;left:0;background:#fff;min-width:220px;box-shadow:0 8px 20px rgba(2,6,23,0.12);border-radius:8px;overflow:hidden;z-index:999}
.nav-item:hover .dropdown-content{display:block}
.dropdown-content a{display:block;padding:10px 12px;color:#111827;text-decoration:none}
.dropdown-content a:hover{background:#f4f6f8}
@media(max-width:880px){aside.sidebar{display:none}main.content{padding:12px}}
main.content{flex:1;padding:16px}
.card{background:var(--card);border-radius:12px;padding:20px;box-shadow:0 6px 18px rgba(15,23,42,0.06)}
.card h2{margin:0 0 12px;font-size:18px;color:#0f1724}
.muted{color:var(--muted);font-size:14px}
label{display:block;font-weight:600;margin-top:12px;color:#111827}
input,select,textarea{width:100%;padding:10px 12px;margin-top:8px;border:1px solid #e2e8f0;border-radius:10px;background:#fff;font-size:14px;color:#0f1724 !important;outline:none !important;-webkit-appearance:none;appearance:none;font-weight:500;transition:border-color 0.2s, box-shadow 0.2s;text-align:left;}
input:focus,select:focus,textarea:focus,*:focus{outline:none !important;border-color:var(--accent) !important;box-shadow:0 0 0 3px rgba(37,99,235,0.1) !important;}
input:disabled,select:disabled,textarea:disabled,input:read-only,select:read-only,textarea:read-only{background:#f8fafc !important;color:#64748b !important;border-color:#e2e8f0;cursor:default;opacity:1 !important;-webkit-text-fill-color:#64748b !important;}
select option { color: #0f1724 !important; background: #fff !important; font-weight:500; }
textarea{min-height:110px;resize:vertical}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;background:var(--accent);color:#fff;padding:10px 18px;border-radius:10px;border:none;font-weight:600;font-size:14px;cursor:pointer;text-decoration:none;transition:all 0.2s;line-height:1;margin-top:16px;font-family:inherit;}
.btn:hover{filter:brightness(1.1);transform:translateY(-1px)}
.btn:active{transform:translateY(0)}
.btn.secondary{background:#1e293b;color:#fff}
.btn.danger{background:#ef4444}
.btn:disabled{opacity:.5;cursor:not-allowed;transform:none}
.small{padding:8px 12px;font-size:13px;border-radius:8px}
.preview{margin-top:12px;border:1px solid #e6e9ee;padding:10px;border-radius:8px;background:#fafafa}
.controls{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}


/* ===== User List Filter (compact) ===== */
.um-filter select{height:32px;padding:6px 10px;font-size:12px;border-radius:8px;}
.um-filter label{font-size:12px;}
.um-filter .btn.small,.um-filter a.btn.small{height:32px;padding:0 10px;font-size:12px;border-radius:8px;display:inline-flex;align-items:center;gap:6px;margin-top:0;}
.upload-item{display:flex;gap:12px;align-items:flex-start;padding:10px 0;border-bottom:1px dashed #f1f5f9}
.upload-item img{max-width:140px;border-radius:8px;border:1px solid #e6e9ee}
.upload-meta{font-size:13px;color:#374151;flex:1}
.badge{display:inline-block;font-size:11px;padding:2px 8px;border-radius:999px;margin-left:8px;background:#eef2ff;color:#4338ca;border:1px solid #c7d2fe;line-height:1.2}
.lock-badge{background:#ecfdf5;color:#065f46;border-color:#a7f3d0}

/* Header row within each upload item */
.upload-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}
.upload-actions{display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap}
.upload-actions form{margin:0}

/* ===== Same-page preview modal ===== */
.modal-overlay{
  position:fixed; inset:0;
  background:rgba(2,6,23,.75);
  display:none;
  align-items:center;
  justify-content:center;
  z-index:2000;
  padding:20px;
}
.modal{
  width:min(1100px, 96vw);
  max-height:90vh;
  background:#fff;
  border-radius:14px;
  overflow:hidden;
  box-shadow:0 20px 60px rgba(0,0,0,.35);
  display:flex;
  flex-direction:column;
}
.modal-header{
  display:flex;
  align-items:center;
  justify-content:space-between;
  padding:12px 14px;
  border-bottom:1px solid #e6e9ee;
  background:#f8fafc;
}
.modal-title{
  font-weight:800;
  font-size:14px;
  color:#0f1724;
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
  max-width:70%;
}
.modal-body{
  padding:14px;
  overflow:auto;
  background:#0b122012;
}
.modal-body img{
  width:100%;
  height:auto;
  border-radius:12px;
  border:1px solid #e6e9ee;
  background:#fff;
}
.modal-actions{display:flex;gap:10px;align-items:center;}

/* ===== Timesheet ===== */
.ts-wrap{margin-top:10px;border:1px solid #e6e9ee;border-radius:10px;background:#fff}
.ts-head{display:flex;gap:10px;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid #eef2f7;background:#f8fafc}
.ts-grid{width:100%;border-collapse:collapse}
.ts-grid th,.ts-grid td{padding:8px;border-bottom:1px solid #eef2f7;text-align:left;font-size:13px}
.ts-grid th{color:#0f1724;background:#fff;vertical-align:top}
.ts-grid td input{margin-top:0;padding:8px;border-radius:8px}
.ts-total{font-weight:900}
details.ts-details summary{cursor:pointer;user-select:none;font-weight:800;color:#111827;margin-top:8px}
details.ts-details{margin-top:8px}
.ts-grid th .ts-dd{
  display:block;
  font-size:12px;
  color:var(--muted);
  font-weight:800;
  margin-top:2px;
}

.filter-card{
  margin-top:14px;
  padding:14px 14px 12px;
  border:1px solid #e6e9ee;
  border-radius:14px;
  background:#fff;
  box-shadow:0 6px 18px rgba(15,23,42,0.05);
}
.filter-title{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:12px;
  margin-bottom:10px;
}
.filter-title .muted{margin:0;font-weight:700}
.filter-form{
  display:flex;
  gap:12px;
  align-items:flex-end;
  flex-wrap:wrap;
}
.filter-field{min-width:220px}
.filter-field.grow{flex:1;min-width:260px}
.filter-field label{
  margin:0;
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.05em;
  color:#64748b;
  font-weight:700;
}
.filter-field input,.filter-field select{
  margin-top:6px;
  height:40px;
  padding:8px 12px;
  border-radius:10px;
  background:#fff !important;
  color:#0f1724 !important;
  border:1px solid #cbd5e1;
}
.filter-search{position:relative}
.filter-search i{
  position:absolute;
  left:12px;
  top:50%;
  transform:translateY(-50%);
  color:#94a3b8;
  pointer-events:none;
}
.filter-search input{padding-left:38px}
.filter-actions{
  display:flex;
  gap:10px;
  align-items:center;
}

/* Buttons inside toolbars should not inherit the big top margin */
/* Filter action buttons: match input height and align icons */
.filter-actions .btn,
.filter-actions a.btn{
  height:40px;
  padding:0 16px;
  display:inline-flex;
  align-items:center;
  justify-content:center;
  gap:8px;
  border-radius:10px;
  white-space:nowrap;
  margin-top:0;
}
.filter-actions a.btn{ text-decoration:none; }
.filter-actions .btn.outline{ border-color:#cbd5e1; color:#475569; }
.filter-actions .btn.outline:hover{ background:#f1f5f9; color:#0f1724; border-color:#94a3b8; }

.controls .btn,
.filter-form .btn{margin-top:0}

/* Outline button variant */
.btn.outline{
  background:#fff;
  color:var(--accent);
  border:1px solid #cbd5e1;
}
.btn.outline:hover{background:#f8fafc}

/* Optional: make small buttons more pill-like */
.btn.small{border-radius:10px}

@media(max-width:880px){aside.sidebar{display:none}main.content{padding:12px}}

/* ===== User List Table (compact) ===== */
.um-table{font-size:12px;color:#334155;table-layout:auto;width:100%;border-spacing:0;}
.um-table th,.um-table td{padding:6px 4px !important;border-bottom:1px solid #f1f5f9;word-break:break-word;text-align:left;vertical-align:middle;}
.um-table th{font-size:11px;text-transform:uppercase;letter-spacing:0.02em;color:#64748b;font-weight:700;background:#f8fafc;}
.um-table select{height:28px;padding:2px 6px;font-size:12px;border-radius:6px;margin:0;}
.um-table .btn.small,.um-table a.btn.small{height:28px;padding:0 8px;font-size:11px;border-radius:6px;display:inline-flex;align-items:center;gap:4px;margin-top:0;}
.um-rowform{display:flex;gap:4px;align-items:center;flex-wrap:wrap;}

</style>
"""


def NAV(display_name="Visitor"):
    brand = (
        '<div class="brand"><div class="logo">VS</div>'
        '<div><div style="font-size:14px">Vaisesika</div>'
        '<div style="font-size:12px;color:#9fb3d8;margin-top:2px">Timesheet Portal</div></div></div>'
    )
    if display_name == "Visitor":
        return (
            '<header class="app-header">'
            f'{brand}'
            '<div class="header-right"><a class="btn outline small" style="padding:8px 12px;font-size:13px;margin-right:8px" href="/help"><i class="fa fa-circle-question"></i> Help</a>Welcome</div>'
            '</header>'
        )
    return (
        '<header class="app-header">'
        f'{brand}'
        f'<div class="header-right">Signed in as <strong>{html.escape(display_name)}</strong>'
        f'<a class="btn outline" style="padding:8px 12px;font-size:13px;margin-left:12px" href="/help"><i class="fa fa-circle-question"></i> Help</a><a class="btn secondary" style="padding:8px 12px;font-size:13px;margin-left:12px" href="/logout">Logout</a>'
        '</div></header>'
    )






def build_sidebar_for_role(role, screen_access="BOTH", can_email=0, can_export=0, can_help=1):
    # Admin always has full access
    if role == "Admin":
        screen_access = "BOTH"
        can_email, can_export, can_help = 1, 1, 1
    # Admin always has full access
    if role == "Admin":
        screen_access = "BOTH"
        can_email, can_export, can_help = 1, 1, 1

    screen_access = normalize_screen_access(screen_access)

    module_links = []
    if role == "Admin":
        module_links = [
            '<a href="/upload/PPM" title="PPM screen prints">PPM</a>',
            '<a href="/upload/NTT" title="NTT screen prints">NTT</a>',
            '<a href="/upload/EMAIL" title="Email screen prints">EMAIL</a>',
        ]
    elif screen_access == "BOTH":
        module_links = [
            '<a href="/upload/PPM" title="PPM screen prints">PPM</a>',
            '<a href="/upload/NTT" title="NTT screen prints">NTT</a>',
        ]
    elif screen_access == "PPM":
        module_links = ['<a href="/upload/PPM" title="PPM screen prints">PPM</a>']
    elif screen_access == "NTT":
        module_links = ['<a href="/upload/NTT" title="NTT screen prints">NTT</a>']
    elif screen_access == "EMAIL":
        module_links = ['<a href="/upload/EMAIL" title="Email screen prints">EMAIL</a>']

    screen_module_dropdown = (
        '<div class="nav-item"><i class="fas fa-image"></i> Screen Print Module'
        '<div class="dropdown-content">'
        + "".join(module_links) +
        '</div></div>'
    )

    parts = ['<aside class="sidebar">']

    if role == "Admin":
        parts.append(
            '<div class="nav-item"><i class="fas fa-users"></i> User Management'
            '<div class="dropdown-content">'
            '<a href="/user-management/create">Create User</a>'
            '<a href="/user-management/list">User List</a><a href="/user-management/create-fields">Create Fields</a>'
            '</div></div>'
        )

    parts.append(screen_module_dropdown)

    if can_email:
        parts.append('<a class="nav-item" href="/email-settings"><i class="fas fa-envelope"></i> Email Settings</a>')
    if can_export:
        parts.append('<a class="nav-item" href="/export"><i class="fas fa-file-export"></i> Export Data</a>')
    parts.append('</aside>')
    return "".join(parts)


def render_page(content, display_name="Visitor", role=None, screen_access="BOTH"):
    if display_name == "Visitor":
        return (
            f"<html><head>{PAGE_STYLE}</head>"
            f"<body>{NAV(display_name)}"
            f"<main class='content' style='margin-top:64px;padding:28px'>{content}</main>"
            f"</body></html>"
        )

    if role == "Admin":
        can_email, can_export, can_help = (1, 1, 1)
    else:
        if isinstance(display_name, str) and "@" in display_name:
            can_email, can_export, can_help = get_feature_flags_by_email(display_name)
        else:
            can_email, can_export, can_help = get_feature_flags_by_username(display_name)

    sidebar_html = build_sidebar_for_role(role or "", screen_access or "BOTH",
                                         can_email=can_email, can_export=can_export, can_help=can_help)
    return (
        f"<html><head>{PAGE_STYLE}</head>"
        f"<body>{NAV(display_name)}"
        f"<div class='layout'>{sidebar_html}<main class='content'>{content}</main></div>"
        f"</body></html>"
    )

def parse_multipart(body: bytes, boundary: bytes):
    out = {}
    if not boundary:
        return out

    parts = body.split(b"--" + boundary)
    for part in parts:
        if not part or part in (b"--", b"--\r\n"):
            continue
        if part.startswith(b"\r\n"):
            part = part[2:]
        try:
            hdr, data = part.split(b"\r\n\r\n", 1)
        except ValueError:
            continue

        data = data.rstrip(b"\r\n")
        hdrs = hdr.decode(errors="ignore")

        name = None
        filename = None
        for line in hdrs.split("\r\n"):
            line = line.strip()
            if line.lower().startswith("content-disposition:"):
                for seg in line.split(";"):
                    seg = seg.strip()
                    if seg.startswith("name="):
                        name = seg.split("=", 1)[1].strip().strip('"')
                    if seg.startswith("filename="):
                        filename = seg.split("=", 1)[1].strip().strip('"')

        if not name:
            continue

        if filename:
            out[name] = {"filename": filename, "content": data}
        else:
            out[name] = data.decode("utf-8", errors="ignore")

    return out


# -------------------------
# Session handling
# -------------------------
SESSIONS = {}  # session_id -> email


def parse_cookies(header_value):
    if not header_value:
        return {}
    cookie = http.cookies.SimpleCookie()
    cookie.load(header_value)
    return {k: morsel.value for k, morsel in cookie.items()}


def make_session(email):
    session_id = str(uuid.uuid4())
    SESSIONS[session_id] = email.lower()
    return session_id


def clear_session_from_header(header_value):
    cookies = parse_cookies(header_value)
    sid = cookies.get("session_id")
    if sid and sid in SESSIONS:
        del SESSIONS[sid]


def get_email_from_request(handler):
    cookie_header = handler.headers.get("Cookie")
    cookies = parse_cookies(cookie_header)
    sid = cookies.get("session_id")
    if not sid:
        return None
    return SESSIONS.get(sid)


def get_display_name_by_email(email):
    row = get_user_db_by_email(email)
    return row[1] if row else (email or "Visitor")


def get_role_by_email(email):
    row = get_user_db_by_email(email)
    return row[4] if row else None


# -------------------------
# Modal for saved screenprints (same page)
# -------------------------
def build_saved_preview_modal():
    return """
<!-- Same-page preview modal -->
<div id="imgModalOverlay" class="modal-overlay" aria-hidden="true">
  <div class="modal" role="dialog" aria-modal="true" aria-labelledby="imgModalTitle">
    <div class="modal-header">
      <div class="modal-title" id="imgModalTitle">Preview</div>
      <div class="modal-actions">
        <a class="btn secondary small" id="imgModalDownload" href="#" download>
          <i class="fa fa-download"></i> Download
        </a>
        <button type="button" class="btn danger small" id="imgModalClose">
          <i class="fa fa-times"></i> Close
        </button>
      </div>
    </div>
    <div class="modal-body">
      <img id="imgModalImg" src="" alt="screenprint preview">
    </div>
  </div>
</div>

<script>
(function(){
  const overlay = document.getElementById('imgModalOverlay');
  const img     = document.getElementById('imgModalImg');
  const title   = document.getElementById('imgModalTitle');
  const btnDl   = document.getElementById('imgModalDownload');
  const btnClose= document.getElementById('imgModalClose');

  function openModal(url, filename){
    img.src = url;
    title.textContent = filename || 'Preview';
    btnDl.href = url;
    btnDl.setAttribute('download', filename || 'screenprint.png');
    overlay.style.display = 'flex';
    overlay.setAttribute('aria-hidden','false');
  }

  function closeModal(){
    overlay.style.display = 'none';
    overlay.setAttribute('aria-hidden','true');
    img.src = '';
    btnDl.href = '#';
  }

  btnClose.addEventListener('click', closeModal);

  overlay.addEventListener('click', function(e){
    if(e.target === overlay) closeModal();
  });

  document.addEventListener('keydown', function(e){
    if(e.key === 'Escape' && overlay.style.display === 'flex'){
      closeModal();
    }
  });

  document.addEventListener('click', function(e){
    const a = e.target.closest('a.thumb-link');
    if(!a) return;
    e.preventDefault();
    openModal(a.getAttribute('data-url'), a.getAttribute('data-filename'));
  });
})();
</script>
"""


# -------------------------
# Timesheet UI — now accepts admin_delete_html for swapping positions
# -------------------------
def build_timesheet_ui(upload_id: int, return_to: str, admin_delete_html: str = ""):
    ts = get_timesheet_by_upload(upload_id)
    module_for_upload = get_upload_module(upload_id)
    if not ts:
        ensure_timesheet_for_upload(upload_id, utc_now_str(sep=" ", timespec="seconds"))
        ts = get_timesheet_by_upload(upload_id) or {
            "week_start": _sunday_of_date(datetime.date.today()).isoformat(),
            "mon": 0, "tue": 0, "wed": 0, "thu": 0, "fri": 0, "sat": 0, "sun": 0,
            "total": 0, "updated_at": "", "submitted": 0, "submitted_at": ""
        }

    def fnum(v):
        try:
            vv = float(v)
            return str(int(vv)) if vv.is_integer() else str(vv)
        except Exception:
            return "0"

    is_submitted = int(ts.get("submitted", 0)) == 1
    # Separate disable flags: hours may be locked while week_start stays editable until submission
    hours_disabled_attr = "disabled" if is_submitted else ""
    week_disabled_attr = "disabled" if is_submitted else ""
    # PPM & NTT: auto-extracted hours; prevent manual edits of hours, but allow week_start changes until submitted
    auto_locked_hours = (module_for_upload in ("PPM", "NTT"))
    if auto_locked_hours:
        hours_disabled_attr = "readonly"
    try:
        ws_date = datetime.date.fromisoformat(ts["week_start"])
    except Exception:
        ws_date = datetime.date.today()
    ws_date = _sunday_of_date(ws_date)
    we_date = ws_date + datetime.timedelta(days=6)

    def fmt_range(d: datetime.date) -> str:
        return d.strftime("%d/%b/%Y")

    week_range_text = f"{fmt_range(ws_date)} – {fmt_range(we_date)}"
    day_dates = [ws_date + datetime.timedelta(days=i) for i in range(7)]
    dd = [d.strftime("%d") for d in day_dates]

    sun = html.escape(fnum(ts["sun"]))
    mon = html.escape(fnum(ts["mon"]))
    tue = html.escape(fnum(ts["tue"]))
    wed = html.escape(fnum(ts["wed"]))
    thu = html.escape(fnum(ts["thu"]))
    fri = html.escape(fnum(ts["fri"]))
    sat = html.escape(fnum(ts["sat"]))
    total = html.escape(fnum(ts["total"]))
    week_start_iso = html.escape(ws_date.isoformat())

    ppm_is_verify_attached = False
    verify_html = ""
    submit_disabled_attr = ""
    
    if module_for_upload == "PPM":
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT id, stored_filename FROM uploads WHERE module='VERIFY' AND original_filename=?", (f"PPM_{upload_id}",))
        v_row = cur.fetchone()
        conn.close()
        if v_row:
            ppm_is_verify_attached = True
            v_img = f"/uploads/{html.escape(v_row[1], quote=True)}"
            verify_html = f"""
            <div style="margin-top:12px; padding:10px 14px; border-radius:8px; background:linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%); border:1px solid #6ee7b7; display:flex; align-items:center; gap:10px;">
                <div style="width:32px;height:32px;border-radius:50%;background:#10b981;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                  <i class="fa fa-check" style="color:#fff;font-size:14px;"></i>
                </div>
                <span style="color:#065f46;font-weight:600;font-size:13px;">Confirmation Screen Print Attached</span>
            </div>
            """
        else:
            verify_html = f"""
            <div style="margin-top:12px; padding:10px 14px; border-radius:8px; background:linear-gradient(135deg, #fef2f2 0%, #fde8e8 100%); border:1px solid #fca5a5;">
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                <div style="width:32px;height:32px;border-radius:50%;background:#ef4444;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                  <i class="fa fa-exclamation" style="color:#fff;font-size:14px;"></i>
                </div>
                <span style="color:#991b1b; font-weight:600; font-size:13px;">Attach confirmation screen print and click on submit.</span>
              </div>
              <button type="button" class="btn small outline" onclick="captureVerify({upload_id}, this)" {'disabled' if is_submitted else ''}
                style="border-color:#ef4444; color:#dc2626; font-size:12px;">
                <i class="fa fa-camera"></i> Capture Confirmation Screen Print
              </button>
            </div>
            """
            
    if module_for_upload == "PPM" and not ppm_is_verify_attached:
        submit_disabled_attr = "disabled"

    has_draft_error = ""
    if module_for_upload == "NTT" and int(ts.get("has_draft", 0)) == 1:
        submit_disabled_attr = "disabled"
        has_draft_error = f"""
        <div style="margin-top:12px; padding:12px 16px; border-radius:10px; background:#fff1f2; border:1px solid #fecaca; color:#991b1b; display:flex; align-items:center; gap:12px; font-weight:600; font-size:14px; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
            <i class="fa fa-triangle-exclamation" style="font-size:18px; color:#ef4444;"></i>
            <span>NTT status is in Draft. Please submit the timesheet in NTT portal and re-capture the screenshot.</span>
        </div>
        """

    rem_planned = ts.get("remaining_planned")
    rem_warning = ""
    if module_for_upload == "PPM" and rem_planned is not None:
        try:
            if float(rem_planned) < 45.0:
                rem_warning = f"""
                <div style="margin-top:12px; padding:12px 16px; border-radius:10px; background:#fff1f2; border:1px solid #fecaca; color:#991b1b; display:flex; align-items:center; gap:12px; font-weight:600; font-size:14px; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
                    <i class="fa fa-triangle-exclamation" style="font-size:18px; color:#ef4444;"></i>
                    <span>Warning: Your Remaining Planned hours are less than 45 hrs. Please contact your manager to top up your hours for next week.</span>
                </div>
                """
        except (ValueError, TypeError):
            pass

    if not is_submitted:
        save_btn_html = ""
        if module_for_upload not in ("PPM", "NTT"):
            save_btn_html = """<button class="btn small" type="submit">
            <i class="fa fa-save"></i> Save Hours
          </button>"""
        controls_html = f"""
        {rem_warning}
        {has_draft_error}
        {verify_html}
        <div class="controls" style="margin-top:12px;">
          {save_btn_html}

          <button class="btn secondary small ts-submit-btn" type="submit" formaction="/timesheet/submit" {submit_disabled_attr}>
            <i class="fa fa-check"></i> Submit
          </button>
        </div>
        """
    else:
        controls_html = f"""
        {rem_warning}
        <div class="muted" style="margin-top:8px">
          This timesheet is locked. Contact Admin if changes are required.
        </div>
        """

    # these are defined above
    ts_head_right_extra = admin_delete_html or ""

    return f"""
<details class="ts-details">
  <summary>
    Timesheet (Week-wise) — <span id="ts_range_{upload_id}">{html.escape(week_range_text)}</span>
  </summary>

  <div class="ts-wrap ts-form" data-upload-id="{upload_id}">
    <div class="ts-head">
      <div class="muted" style="font-weight:700;">
        Week Start (Sun):
        <input class="ts-week" type="date" value="{week_start_iso}"
               style="width:auto;display:inline-block;margin-left:8px;" {week_disabled_attr}>
      </div>
      <div class="muted" style="font-weight:700; display:flex; align-items:center; gap:10px;">
        <span>Total: <span class="ts-total" id="ts_total_lbl_{upload_id}">{total}</span></span>
        {ts_head_right_extra}
      </div>
    </div>

    <form method="post" action="/timesheet/save" style="padding:10px 12px;">
      <input type="hidden" name="upload_id" value="{upload_id}">
      <input type="hidden" name="return_to" value="{html.escape(return_to, quote=True)}">
      <input type="hidden" name="week_start" value="{week_start_iso}" class="ts-week-hidden">

      <table class="ts-grid">
        <thead>
          <tr>
            <th style="width:12%">Sun <span class="ts-dd ts-dd-0">{dd[0]}</span></th>
            <th style="width:12%">Mon <span class="ts-dd ts-dd-1">{dd[1]}</span></th>
            <th style="width:12%">Tue <span class="ts-dd ts-dd-2">{dd[2]}</span></th>
            <th style="width:12%">Wed <span class="ts-dd ts-dd-3">{dd[3]}</span></th>
            <th style="width:12%">Thu <span class="ts-dd ts-dd-4">{dd[4]}</span></th>
            <th style="width:12%">Fri <span class="ts-dd ts-dd-5">{dd[5]}</span></th>
            <th style="width:12%">Sat <span class="ts-dd ts-dd-6">{dd[6]}</span></th>
            <th style="width:16%">Total</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="sun" value="{sun}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="mon" value="{mon}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="tue" value="{tue}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="wed" value="{wed}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="thu" value="{thu}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="fri" value="{fri}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="sat" value="{sat}" {hours_disabled_attr}></td>
            <td>
              <input type="text" readonly value="{total}" class="ts-total-input" id="ts_total_in_{upload_id}">
            </td>
          </tr>
        </tbody>
      </table>

      {controls_html}
    </form>
  </div>
</details>
"""


def build_timesheet_js():
    return """
<script>
(function(){
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  window.captureVerify = async function(uploadId, btn) {
    try {
      btn.disabled = true;
      const stream = await navigator.mediaDevices.getDisplayMedia({
        video: { cursor: 'always' },
        audio: false
      });
      const track = stream.getVideoTracks()[0];
      const v = document.createElement('video');
      v.srcObject = stream;
      v.autoplay = true;
      v.muted = true;
      v.playsInline = true;
      await v.play();
      
      // small delay to let user settle window
      await new Promise(r => setTimeout(r, 2000));
      
      const w = v.videoWidth || 1280;
      const h = v.videoHeight || 720;
      const c = document.createElement('canvas');
      c.width = w; c.height = h;
      const ctx = c.getContext('2d');
      ctx.drawImage(v, 0, 0, w, h);
      
      c.toBlob(async (b) => {
        stream.getTracks().forEach(t=>t.stop());
        if(!b) { alert('Could not create image.'); btn.disabled = false; return; }
        
        btn.innerHTML = "<i class='fa fa-spinner fa-spin'></i> Uploading...";
        const fd = new FormData();
        fd.append('file', b, 'Verify_capture_' + Date.now() + '.png');
        fd.append('module', 'VERIFY');
        fd.append('parent_upload_id', uploadId);
        
        const res = await fetch('/upload/VERIFY', { method:'POST', body:fd });
        if(res.ok) {
           const form = btn.closest('.ts-form');
           if(form) {
               const submitBtn = form.querySelector('.ts-submit-btn');
               if (submitBtn) submitBtn.disabled = false;
           }
           const container = btn.parentElement;
           container.style.background = 'linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%)';
           container.style.borderColor = '#6ee7b7';
           container.style.display = 'flex';
           container.style.alignItems = 'center';
           container.style.gap = '10px';
           container.innerHTML = `
                <div style="width:32px;height:32px;border-radius:50%;background:#10b981;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                  <i class="fa fa-check" style="color:#fff;font-size:14px;"></i>
                </div>
                <span style="color:#065f46;font-weight:600;font-size:13px;">Confirmation Screen Print Attached</span>
           `;
        } else {
           alert('Upload failed: ' + res.statusText);
           btn.disabled = false;
           btn.innerHTML = "<i class='fa fa-camera'></i> Capture Confirmation Screen Print";
        }
      }, 'image/png');
    } catch(e) {
      console.error(e);
      alert('Capture cancelled or error.');
      btn.disabled = false;
    }
  };

  function toNum(v){
    const x = parseFloat(v);
    return isNaN(x) ? 0 : x;
  }

  function parseISO(iso){
    if(!/^\\d{4}-\\d{2}-\\d{2}$/.test(iso)) return null;
    const d = new Date(iso + 'T00:00:00');
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function toISO(d){
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const da = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${da}`;
  }

  function fmtRange(d){
    const dd = String(d.getDate()).padStart(2,'0');
    const mmm = MONTHS[d.getMonth()];
    const yyyy = d.getFullYear();
    return `${dd}/${mmm}/${yyyy}`;
  }

  function sundayOf(d){
    const dow = d.getDay(); // Sun=0
    const s = new Date(d.getTime());
    s.setDate(s.getDate() - dow);
    s.setHours(0,0,0,0);
    return s;
  }

  function updateWeekUI(container, sunday){
    const upId = container.getAttribute('data-upload-id');

    const rangeSpan = document.getElementById('ts_range_' + upId);
    if(rangeSpan){
      const sat = new Date(sunday.getTime());
      sat.setDate(sat.getDate() + 6);
      rangeSpan.textContent = fmtRange(sunday) + ' – ' + fmtRange(sat);
    }

    const dds = container.querySelectorAll('.ts-dd');
    dds.forEach((el, idx) => {
      const x = new Date(sunday.getTime());
      x.setDate(x.getDate() + idx);
      el.textContent = String(x.getDate()).padStart(2,'0');
    });
  }

  function recalc(container){
    const days = container.querySelectorAll('input.ts-day');
    let total = 0;
    days.forEach(inp => total += toNum(inp.value));
    total = Math.round(total * 100) / 100;

    const upId = container.getAttribute('data-upload-id');
    const totalInput = document.getElementById('ts_total_in_' + upId);
    const totalLbl   = document.getElementById('ts_total_lbl_' + upId);
    if(totalInput) totalInput.value = total.toString();
    if(totalLbl) totalLbl.textContent = total.toString();

    const submitBtn = container.querySelector('.ts-submit-btn');
    if(submitBtn){
      submitBtn.disabled = (total <= 0);
      submitBtn.title = submitBtn.disabled
        ? 'Enter hours (Total must be > 0) before submitting.'
        : 'Submit and lock this week (PPM and NTT must match).';
    }
  }

  document.addEventListener('input', function(e){
    if(e.target.classList && e.target.classList.contains('ts-day')){
      const wrap = e.target.closest('.ts-form');
      if(wrap) recalc(wrap);
    }
  });

  document.addEventListener('change', function(e){
    if(e.target.classList && e.target.classList.contains('ts-week')){
      const wrap = e.target.closest('.ts-form');
      if(!wrap) return;

      const chosen = parseISO(e.target.value);
      if(!chosen) return;

      const sunday = sundayOf(chosen);
      const sundayISO = toISO(sunday);

      e.target.value = sundayISO;
      const hidden = wrap.querySelector('.ts-week-hidden');
      if(hidden) hidden.value = sundayISO;

      updateWeekUI(wrap, sunday);
    }
  });

  window.addEventListener('load', function(){
    document.querySelectorAll('.ts-form').forEach(function(wrap){
      const weekInput = wrap.querySelector('input.ts-week');
      const d = weekInput ? parseISO(weekInput.value) : null;
      const sunday = d ? sundayOf(d) : sundayOf(new Date());

      if(weekInput){
        weekInput.value = toISO(sunday);
        const hidden = wrap.querySelector('.ts-week-hidden');
        if(hidden) hidden.value = weekInput.value;
      }

      updateWeekUI(wrap, sunday);
      recalc(wrap);
    });
  });
})();
</script>
"""


# -------------------------
# Capture UI builder — PPM/NTT only (NO download for unsaved)
# -------------------------
def build_capture_ui(module_ctx: str):
    title = f"Screen Print Capture - {"Email" if module_ctx=="EMAIL" else module_ctx}"
    upload_url = f"/upload/{module_ctx}"
    hint = f"This page captures and lists only {module_ctx} screen prints."
    capture_label = f"Capture {module_ctx} Screen Print"

    parts = []
    parts.append(f"<h2>{html.escape(title)}</h2>")
    parts.append(
        "<p class='muted' style='margin-bottom:8px'>"
        "Click <b>Capture</b> to select a window/tab/screen. A snapshot will be automatically saved without clicking any save button. "
        + html.escape(hint) +
        "</p>"
    )

    parts.append(
        "<div style='display:flex;gap:12px;flex-wrap:wrap;align-items:center'>"
        f"<button type='button' class='btn' id='btnCapture'>{html.escape(capture_label)}</button>"
        "</div>"
    )

    parts.append("<video id='capPreview' style='display:none' autoplay muted playsinline></video>")

    # For PPM & NTT, hide Save and Close buttons entirely (auto-upload handles it)
    save_btn_style = "display:none" if module_ctx in ("PPM", "NTT", "EMAIL") else ""
    close_btn_style = "display:none" if module_ctx in ("PPM", "NTT", "EMAIL") else ""
    parts.append(
        "<div id='shotContainer' class='preview' style='display:none;margin-top:12px'>"
        "  <div id='shotBox'></div>"
        "  <div class='controls'>"
        f"    <button type='button' class='btn' id='btnSave' style='{save_btn_style}'><i class='fa fa-cloud-upload-alt'></i> Save</button>"
        f"    <button type='button' class='btn danger' id='btnCloseShot' style='{close_btn_style}'><i class='fa fa-times'></i> Close</button>"
        "  </div>"
        "</div>"
    )

    js = f"""
<script>
(function(){{
  let stream=null, track=null, lastBlob=null, lastUrl=null;
  let isUploading=false;

  const v=document.getElementById('capPreview');
  const btnCapture=document.getElementById('btnCapture');
  const shotContainer=document.getElementById('shotContainer');
  const shotBox=document.getElementById('shotBox');
  const btnSave=document.getElementById('btnSave');
  const btnCloseShot=document.getElementById('btnCloseShot');

  const moduleName = {module_ctx!r};
  const uploadUrl  = {upload_url!r};

  const AUTO_UPLOAD = (moduleName === "PPM" || moduleName === "NTT" || moduleName === "EMAIL");

  function stopStream(){{
    try {{
      if(track) track.stop();
      if(stream) stream.getTracks().forEach(t=>t.stop());
    }} catch(e) {{}}
    stream=null; track=null;
    try {{
      v.pause();
      v.srcObject=null;
      v.style.display='none';
    }} catch(e) {{}}
  }}

  function clearShot(){{
    try {{
      if(lastUrl) URL.revokeObjectURL(lastUrl);
    }} catch(e) {{}}
    lastUrl=null;
    lastBlob=null;
    shotBox.innerHTML='';
    shotContainer.style.display='none';
  }}

  function waitForVideoReady(video, timeoutMs=2500){{
    return new Promise((resolve) => {{
      if(!video) return resolve();
      if (video.readyState >= 2 && video.videoWidth > 0 && video.videoHeight > 0) {{
        return resolve();
      }}
      let done=false;
      const finish = () => {{
        if(done) return;
        done=true;
        video.onloadedmetadata=null;
        resolve();
      }};
      const t = setTimeout(() => {{
        clearTimeout(t);
        finish();
      }}, timeoutMs);

      video.onloadedmetadata = () => {{
        clearTimeout(t);
        finish();
      }};
    }});
  }}

  async function captureOnce(){{
    clearShot();
    try {{
      btnCapture.disabled = true;

      stream = await navigator.mediaDevices.getDisplayMedia({{
        video: {{ cursor: 'always' }},
        audio: false
      }});

      const tracks = stream.getVideoTracks();
      track = (tracks && tracks[0]) ? tracks[0] : null;

      if(track) {{
        track.onended = () => stopStream();
      }}

      v.srcObject = stream;
      v.style.display = 'block';

      try {{ await v.play(); }} catch(e) {{}}

      await waitForVideoReady(v, 2500);
      await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));

      const w = v.videoWidth || 1280;
      const h = v.videoHeight || 720;

      const c = document.createElement('canvas');
      c.width = w;
      c.height = h;

      const ctx = c.getContext('2d');
      ctx.drawImage(v, 0, 0, w, h);

      c.toBlob((b) => {{
        if(!b) {{
          alert('Could not create image.');
          stopStream();
          btnCapture.disabled = false;
          return;
        }}

        lastBlob = b;
        if(lastUrl) URL.revokeObjectURL(lastUrl);
        lastUrl = URL.createObjectURL(b);

        shotBox.innerHTML = '<img style="max-width:100%;border-radius:10px;box-shadow:0 4px 18px #0001" src="'+lastUrl+'">';
        shotContainer.style.display='block';
             // Auto-upload for PPM and NTT to avoid manual Save click
             if(AUTO_UPLOAD){{
               setTimeout(() => {{ try{{ saveShot(); }}catch(e){{}} }}, 50);
             }}
             stopStream();
        btnCapture.disabled = false;
      }}, 'image/png');
    }} catch(e) {{
      console.error(e);
      alert('Capture was not started (permission denied or cancelled).');
      stopStream();
      btnCapture.disabled = false;
    }}
  }}

  async function saveShot(){{
    if(isUploading){{ return; }}
    if(!lastBlob){{ alert('No captured image to save.'); return; }}
    isUploading = true;
    try {{
      const fd = new FormData();
      const tmpName = (moduleName||'Generic') + '_capture_' + Date.now() + '.png';
      fd.append('file', lastBlob, tmpName);
      fd.append('module', moduleName);

      const res = await fetch(uploadUrl, {{ method:'POST', body:fd }});
      if(res.ok) {{
        try{{ btnSave.disabled = true; }}catch(e){{}}
        alert('Captured image uploaded successfully.');
        location.reload();
      }} else {{
        alert('Upload failed ('+res.status+').');
        isUploading = false;
      }}
    }} catch(e) {{
      console.error(e);
      alert('Upload error.');
      isUploading = false;
    }}
  }}

  btnCapture.onclick = captureOnce;
  btnSave.onclick = saveShot;

  // Auto-upload: enabled for PPM and NTT; Save remains for re-upload if needed
  if(AUTO_UPLOAD){{
    try{{ btnSave.style.display = 'none'; }}catch(e){{}}
  }}
  btnCloseShot.onclick = clearShot;
}})();
</script>
"""
    parts.append(js)
    return "".join(parts)


def normalize_return_to(rt: str, role: str, email: str) -> str:
    rt = rt or "/upload/PPM"
    if rt not in ("/upload/PPM", "/upload/NTT", "/upload/EMAIL"):
        rt = "/upload/PPM"

    if role != "Employee":
        return rt

    access = get_screen_access_by_email(email)
    if access == "PPM":
        return "/upload/PPM"
    if access == "NTT":
        return "/upload/NTT"
    if access == "EMAIL":
        return "/upload/EMAIL"
    return rt


# -------------------------
# HTTP Handler
# -------------------------
class Handler(SimpleHTTPRequestHandler):
    def _normalize_path(self, raw_path):
        parsed = urlparse(raw_path)
        path = parsed.path or "/"
        if path != "/" and path.endswith("/"):
            path = path[:-1]
        return path

    def _render_traceback_page(self, exc: Exception):
        tb = traceback.format_exc()
        html_tb = (
            "<pre style='white-space:pre-wrap;background:#111;color:#fff;padding:12px;border-radius:6px'>"
            + html.escape(tb)
            + "</pre>"
        )
        content = (
            "<div class='card' style='max-width:900px;margin:24px auto;background:#fff7f7;border:1px solid #ffdede'>"
            "<h2 style='color:#7f1d1d'>Server error (debug)</h2>"
            "<p class='muted'>An exception occurred while processing your request. The traceback is shown below for debugging.</p>"
            f"{html_tb}"
            "</div>"
        )
        return content

    def _forbidden(self, message, display_name, role, screen_access):
        self.send_response(403)
        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        content = f"<div class='card' style='text-align:center'><h2>{html.escape(message)}</h2></div>"
        self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))

    def _msg(self, text, display_name="Visitor"):
        session_email = get_email_from_request(self)
        role = get_role_by_email(session_email) if session_email else None
        screen_access = get_screen_access_by_email(session_email) if session_email else "BOTH"

        content = f"<div class='card' style='text-align:center'><h2>{html.escape(text)}</h2></div>"
        self.send_response(200)
        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))

    def _redirect_login_popup(self, msg: str):
        """Redirect back to login page with an error query param so login page can show alert()."""
        self.send_response(303)
        self.send_header("Location", "/?err=" + quote(msg))
        self.end_headers()

    # -------------------------
    # GET
    # -------------------------
    def do_GET(self):
        try:
            path = self._normalize_path(self.path)
            session_email = get_email_from_request(self)
            display_name = get_display_name_by_email(session_email) if session_email else "Visitor"
            role = get_role_by_email(session_email) if session_email else None
            screen_access = get_screen_access_by_email(session_email) if session_email else "BOTH"

            # Home / Login
            if path == "/":
                if not session_email:
                    parsed = urlparse(self.path)
                    qs = parse_qs(parsed.query)
                    err = (qs.get('err', [''])[0] or '').strip()

                    alert_js = ''
                    if err:
                        safe = html.escape(err, quote=True)
                        safe = safe.replace('\\', '\\\\').replace("'", "\\'")
                        alert_js = "<script>window.addEventListener('load', function(){alert('\"Invalid User ID or Password\"');});</script>"

                    content = (
                        "<div class='card' style='max-width:420px;margin:40px auto;text-align:left'>"
                        "<h2 style='text-align:center'>Sign in</h2>"
                        f"{alert_js}"
                        "<form method='post' action='/login'>"
                        "<label>Email</label>"
                        "<input type='email' name='email' required>"
                        "<label>Password</label>"
                        "<input type='password' name='password' required>"
                        "<button class='btn' type='submit' style='width:100%'>Sign In</button>"
                        "</form>"
                        "</div>"
                    )
                    self.send_response(200)
                    self.send_header("Content-type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(render_page(content, "Visitor").encode("utf-8"))
                    return
                content = (
                    "<div class='card'>"
                    "<h2 style='text-align:center'>Welcome to Timesheet Portal</h2>"
                    "<div style='text-align:center;margin-top:14px'>"
                    "<img src='data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AAB"
                    "AAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
                    "AAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJD"
                    "AAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZ"
                    "WiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApb"
                    "AAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAEBAQEBAQEBAQEBAQEBAQEB"
                    "AQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB"
                    "AQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAJYAlgDAREAAhEBAxEB/8QAHwABAAICAwEBAQEAAAAAAAAAAAkKBwgFBgsEAwIB/8QAQhAAAAYCAQMD"
                    "AwMDAwIFAgQHAAECAwQFBhEHCBIhCRMxFCJBChVRFjJhFyNCGHEkJ1KBkRk2GiUzcjQ3Q2KCocH/xAAeAQEAAgIDAQEBAAAAAAAAAAAACAkGBwIEBQMBCv/E"
                    "AFcRAAEEAgEDAwMCAwUEBgUBGQEAAgMEBREGEiExBxNBCCJRFGEVMnEWI0KBkQkkUqEXJTNiscEYJidD4fA0NUVW0TZERlNVcoI3VGNmdYOFkpSVpdTx/9oA"
                    "DAMBAAIRAxEAPwC/wCICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICICICLiay+pLp+5i1FtW2crHbU6O/jQZsaVIpLkoEC2/ardhhxbtdYqqrWrtEQ5iGZC62zrp6G1"
                    "RJ0V537S154GwPmhliZah/UVnyRvY2xB7ssHvQOcAJYhPDNCZGFzRLDLGT1xvaPhDarWHWGQTwzPqT/prTI5GPfWse1FY9idrXF0MpgngnEcga8wzQyhpjlY"
                    "53LD4r7oCICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICL4HbWtYs4VK9PiN29lDsbGBVrfbKfMr6h2uj2"
                    "s+PE7vfdhVsi4qI86WlBx4si1rWH3EOz4qHfoIZXRSTtjeYYnxRSShp9tkkwldDG5+ukSSthmdGwnqe2GVzQRG8j4usQMniqumjbZmjmmhgLx70kNd0LJ5mR"
                    "763RQPs12TSAFkb7EDXua6aMO+8fNfZARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARRV+oGvlTp1u8a6zOEX22Spm"
                    "IeD83Y4qvhP0uQUUufFLDsqyOKVnTy7RuJYF/QU8on7nkb55HhSqqbi9XjVpdxpL/T/T4p6gz3vSLmzCIs02xk+HZZkr4ruGztetI67BVnLJYooLtKN1ySKy"
                    "z9GZce4vjmmnbFLDv6pZ+d+mEOJ9fPTSwwW+MSVMP6h8dst68RyriNy0IaMmRrMb7pt4bKWYo6+RqPiyFKnfsSNkmq1n0bG0fSn1fcb9VWJM2mOqLHc0hQI8"
                    "zJcCnzW5c6vbeJok21HPJiEWS4rLU9HdgXbMGDLaYmQGsgpsetZSatGvvVT0h5P6U5h9LLRm9iJpnx4vkFaF7KV3oLtwTNLpP0WRi6HtsUnyyt64pX07F2q1"
                    "tp+0vRD1+4P654N17j1g47P0I2f2g4jkJov4xhpXOMfugM6W5DFyyAirlKzfZkBZHYjqW/cqx7YjVK3kgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgI"
                    "gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgItYOo/qnwHp4x2zlWcqPcZVFrv3FGNxnlrXXQnVG3Gs8iXFbkvVkWa8S2KSCbKrbJ5jMmPSRXotdeWVNtH009KOR+"
                    "pOTq1qEL6mLmtfpZMrO0NiklY33JoKQkdG2zLBGWvuShwq46KSGS9NC6xUitR99d/qL4V6GYWebJyszXKpKosYziVGzC2/KyR5iguZJ7iW4vHSSh8deaw338"
                    "jLFPBjILUle0a+K+hyDnPIOP5L1TcsOokZjzfIQrB4JnN9vDOGal95WMY9WRpRuRK+FdWzttk6HKJ+VXZBTSsXu5djbWrkqe7kHrg7juAzdT024ewDCcIiFb"
                    "KXWvjfJnuWTxsOYytqWMNfLLXaIcWyOdrHU31bUEMNaEtgZjv0t1uZ8m4re9Z/Uqw+fmHqjYdkcXQfBLVrcS4DWlkZxnjeIqSkmrjrG7PIepzpbd4ZavPkrV"
                    "u5G6Zb5jRylIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIuj8mceYzyzx9mPGmYwin4zm+P2WO2zJEkn2mLGOtlE"
                    "+A+ZGuFbVcg2bOnsmDRLq7WJDsYbrMuKy6j1sDmshxzNYrP4qd1fI4e/VyNOZpI6Z6srZmB4GuuJ5b7c0Z2yWJz4ngse4HxuRYHGcpwOZ43mazLeJzuMu4nI"
                    "1pGhzZal+vJWnaAQQHhkhdG8fdHIGyNIc0EUxodjyX0VdRF/xZd5DOw7JMFzZxnDswjL/ZoVVdtS27CAUtFjImMOcfchRJdZkNQ3kE2bAr4eQss37S8ey7Lu"
                    "26CuzjHrZ6aYvkgx1XL4fkeIaM7hnxvtyxvZF+muOpuhjinizOFsR2qkklKOKxY/TMmpuN3H4su/n85JiedehfqrmMbh8rewPOuBZRwxeYoubX/jmG3HdxFu"
                    "eCYuqXaOSxX6d9nHXI7NO5EZaFuOUVjFJZR6N/UAxrnxxjjjkmPDwbmaMaYUaMbnsUGevMRH35a8fS+s5FRkEYoUx6fitgpS3YrZT6CdaIZu6/HK2/Wz6esx"
                    "6avfn8C+fO8IsalZeDQ+5hmSyMZFFk/aHtzVXmWKODKQBsL5Htisx1pJKxtWq/TP9W3HfWurW4zyVtTjPqdBC4TYkPdHjOSNgjfJNe45JO5z/dbFG6e7hZpJ"
                    "LdJofJBLdqxyWI5IBG1THQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEWk/UN1Ww8JkzePONVot8/dbfhz7xCI0"
                    "mnw11cd4j9v6knYdvk0Z32VfQvMyKWjI3JuRnJeisYveb09NPSGfkUcXI+TB9HjcZjlr1CZI7mbAlaC0GMtlp41zRIHWGuZcuu6K+NaxssuUx8MfqO+qvH+m"
                    "os8J4AK2e9R7DXV55j0WMTxAywvMc9/ZMV/MgmN9fEb9iqwvu5d7WRV8ZlYO6lV11n9SGFcPUtlKssMsMrO5zfKGrGWxOyeDXR2384ziTZkmTOX7WNVruOcf"
                    "z5zLq35r1CqxkwEX7EWqnbmJaPoR6V5jlk9SKnyGbHDE8bxhhifFh5bPufwrFxQkMhY9lyV2Vz0UD+nbLEUImdTksWawPTvjOa+pj12wXGMlkLeZx9jKO5J6"
                    "gZuzLKZsnRxfsOy9mSwGEtbNXFbAYJj2MbELEMrYYIia9a0hXV0Gor4NTVxGIFbWQ4tdXQYraWYsKDCYRGiRIzKCJDTEaO02yy2kiShtCUpIiIhVHPPNZmms"
                    "2JXzWLEsk880ji+SWaV5kllkedlz5Huc57iSXOJJ7lf0D168FSvBUqxRwVqsMVevBE0MihghY2OKKNg0GxxxtaxjQNNa0AdgvsHyX2QEQEQEQEQEQEQEQEQE"
                    "QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEVen1w+lxq4oMb6laGGwRM/t/HnJaG2ZyjM5Lzv9B5RIS0Z1EdCpLkrCbeynIRZWsuw44o4"
                    "0hxmAxHbn59EXqc3H5nIemmTlf7OTM+b465z2Bkd2vXd/FcewOIlLrVWOO/XjjJjY6lfkLfcsdSrZ+v70qknxeC9aMJX3b46YOPcrELHCSXC3bJdiMpM9jT9"
                    "uKyMslKSR7vtiysRP91Xd0wW8Rcl2dv9NithPsInIuGRTl45cMS5JWGUYxRsFMM0SW1HKVlWFRYz0p+Wh/3rXE4iLcm27DFbuxup+ZrFUqjrImqwT8ezb3QX"
                    "K08UJr07l4mB0MjHn2jj8tJK2OOExgQZKY1uqSHJVYadXeRqv6anKMJNNSu0LVW4ZqD3VbWOyNadklLJ0pK3tyVCyeOJ8UsLozVthgY6My12Kwp0ReqEzJVA"
                    "4q6m76O0bZJj4xzJaOMQmijIYiojU3JD5rKM4+04mYuNn6ThocrfpGcyiKsK63zvIK8/XX6U7WN/Wcu9MaclvHjrnyXE4Oqa3UJe50lnBxhvXPWALevFjrmh"
                    "f1nHmSB8dCnZz9Mn1w1c5/DeA+tN2Ghm3OZSw3PZvbrYzLOPSyrT5IfthxuTfoxsyxMePvvEYtfpLkrXW51BBNWYoCICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICL/AAzIiMzMiIiMzMz0REXkzMz8ERF8mHnwvwkAEk6ABJJ8ADuSf6KLPqN60pNu7JwLg+xdTAkE5Bsc/rFupn2zpuus"
                    "SYOFuoSh2JXqaRtvMIq0y5zD52GLya2I3WZRNlp6X+hTYGx8j59VZuNrLFTj1nXt1+zXx2M7G4dD5GkjWIkJjZKBBk45nizi21h/Uv8AWk4ut8B9E8mXSOdN"
                    "RznPqWnh3aSGfHcMnY7vMdOEnJY9thYDJhHe4YcxWiB55zlrF6p/iqgeSV5NgxYeeT4ym0MUlA7CbTFwSJ7bPtIesa5cT98RCdbYrMfSzjKm3DubaJVzW9Pc"
                    "AMpZj5ZfiDcbUsPk45Uezvevsnd7mela5/UYalgSigZonSWcg52Va5opUprVZWfvHG13YuvP7uVvxl+WsmUySVKzmn/cDLpx/WWQ9gtujlJhrh1KQe7Pbhhl"
                    "e9J3gIsK4wv+bchrJUXKeSZsuixwp7MVtcHAceney7KiNqht2sN3KcojTnJ/1M16vtqbGcMta2JHbccmWULvq59RhyXmdbhuNnY/D8RjDrftE9NjkFyGM2ut"
                    "zJXQStx1QQVYgI2S17c2Vie9wc1sduv+z/8ASAcM9NLPqPlqxZyL1HfFNQfKB1VuG0HvGIbE10bZIjlrUlzKyuDzHapuw7+hpgBdLaIjKfqAiAiAiAiAiAiA"
                    "iAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAix5yzxljXM3GmccV5emX/TmeY3Z45ZP1zrTFrXpnsKRGuKaS+xJZh3lJNKNb0k1y"
                    "O+mFawocv2XPZ7D9zjXIclxTkGH5Lh5Ww5PCZCtkab3gujMtaRr/AGp2AtMledodDYi6gJYJJIyQHFeByrjOI5nxrO8Uz1YW8NyHF3MTka56QX1bsL4ZCxzm"
                    "uDJY+oSQyFrvblYx4BLVQS6iuJcz4N5YyXE7ZSabPOMMsciOWFG6S2otzj1l7sC9pHJDSZJV8z2oeRY0uyiMS5NLMrpFhAbKScYr4eBcqxHqLwrE8iptFnGc"
                    "hxjZJK1jTjGLERgvY+yIz/2taUz07JY7p92OT23FunL+enkPGcp6Wc95P6f50e7ZwGUtYuSSRj2QZGi9gfQvMBLHGtlcbLXnOiHME/S1zZIuoZLxTNIXIePq"
                    "yevZj1lzDfYiZrj8M2226S6fN1cW7p4ye1xnFMkWy7IqyQj26GzTNxh914okGdZ8hBNi7ZxduSSdhjfJi7koc512mzTX1rMjgQ/I0wWsshzi+5XLMgA0yTRw"
                    "a45BhP4bOZYty4+51uryuEYcw9e5a0nR0gSVw8M7sa2WMsmEbGytY2Vroa9SXJ+AbCm4v5fsbDK+DzTDqKl1TKp1/wAXNqnmgpdK8glWFniNdGkuJmYeopzl"
                    "fUwYDOAtQFVh4xk8TPXz6XMbzeG7y3gNWDGcx3Lbv45rm18dyR/tmSQuD+mCnmJ5BsXA6Gvame85L+8mN6CbH0wfWfl/Th9Hg3qncuZvgDfYpYrPPbJdzHDo"
                    "WhscUcvtNkt5bj0MYHVB02cjjWsDaPv1QyjFZoxDMMWz7GqfMcKv6rKMWyCImdTX1JMZn1thGUtbSlsSWFKQa2X23Y0lhfa/ElMvxJLbUll1pFYWTxeRwuQt"
                    "4rL0rWNyVCZ1e5RuQvr2a0zNbZLFIGvadEOaSNPY5r2FzHNJuTxGXxefxlHNYTIU8ticnWjuY/JY+xFbpXKszeqOevYhc+KWNw8Oa46ILTpwIHYx0F6KAiAi"
                    "AiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAi4fIMho8UpbHIsktYVJR1Mc5VjZ2L6I8SKySktpNbizLbjrq22I7KCU9JkOtR47bj7rba"
                    "u7jsdfy96tjcZUnvX7kghrVK0bpZppCC4hrGjw1rXPkedMjja+SRzWMc4eVnM5h+NYjIZ7kGSpYfDYqs+3kcnkLEdWnUrx66pJppXNY0Fxaxjdl8kjmRRtdI"
                    "9rTC/wBS/Vbk3M1jIwbBE2FVx6mVJhfQx/8Aass5NElo41ldGtDciNVKajLegY+pUeMxGkvTsnKZNKFFxicvpV6N4rhdaLkPJf01vkXtR2Gvl3JVwPVG4Pr1"
                    "Q0uinuNdK1s18CV5njjhxZiiEs+Spv8AqX+rjN+q1q7wf0/nu4b08E1ilZsRB9fK83a14aLF3qEdjH4GRrC+tiXGCazBK+XOsc50ePx+oWeZjF4QxIrEvpbD"
                    "knJylRMWZeS1JYgH2pRY5G/HWtBqpseUpj6cnWXU3+TPwq9ceRTMXxVG7MDhpeeZkVCZa3GMWYp8s9hkiksNDuqvi45Gtdq7kwHiQskacfi2TWWviuSY43Il"
                    "GVnF8WcpMyKxlZw6viK8jWPY6Yg+9bdEHtP6HHHoeXPaWXb7oa4isVor8VfUjgvijIeoDmjCOMa16xXJzPIzdyPIvak2ciuqiVJvcuym0lud6XJ7dXGtrFt+"
                    "5lMIvL1USrdmqsreKT+0fVDm+P8ATfgWY5I9teN2Potp4THNa2CCa++I1sTjq8MbQI6sUnsh0VdgbWpxSOja2OEhva9EPTLIes/qlxvhLHWJosvdOQ5PkS9z"
                    "56mApSx2M7kZJZCXGxJC41q7nu+/IXKsfV1OVxLG8co8Px6ixPGKuHR43jNPW4/QU1cyiNAqaanhs19ZWwo7ZEhmLChR2YzDSSJKGm0pL4FLV69byd25kshY"
                    "lt3r9qe7dtTvMk9m1alfPYnmkcS58ksr3yPcSS5ziT5X9ImPx9LE0KWLxtaGljsbUr0KFOuxsVerTqQsr1q0EbQGxxQwxsjjY0ANY0ADQXNDqruICICICICI"
                    "CICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICKu1623SzHdTjfVBi9a8a530uDcotwmYCGDlxYizwrL5nZGjylPya+NJw6"
                    "5sZ0+wcfKs43o6iBFP61yXYN9EfqjJVu5L0tyUzPYtmbO8bMhkLxZaIxlsbGS9zOh8LGZGvCyOMMdHlJ5HvdK1orK/2gvpIyajg/WfEQFljGvq8Y5gY2jUmO"
                    "szPHHspIGM6jJUyMxxM8r3HqgyNEF0cdLTqxkC6u+MsuiZRRpS/7KZMOwqpDjzVZkdFOJBWmO3CWlF3wZ6G2XWHux1yrtolXfQWysqqG41ZLfoQ5Wm6tK4xO"
                    "Dmy17TGtdNTtRk+1bgLv/eREuY9gIbYryTVZgYJ5GOrqxlitlqE1C6A6KdjBI4BrpIZGE+zah6h2kiJI0HN92J8tYyNjme5bWuyqu4rK/JcckPTsZyBp2TUS"
                    "pSWkzY7rCyRY0Vy0wtxmLkNA+soNrGaedYUfs2EF6TVzoEp7wKM8pdNVssZFkaLmxXIGHcZ2C6C1WLvufTtta6WtI5ofsSQShs8E0bNdZTGz4u1JUtM7sd1M"
                    "lb1e3PG4B0csL3MaXRSRlr2l7A7w17WvDmjbPo+63+TukPLPcqHl5DxlfWLEnOOO5kl36CzNphcRNzSrdW41j2UNRzjR1XcJjvtq+uq6u/jXEClomavTHrZ9"
                    "P3F/V/Hy2i2PD8wqwObi+QQQsHuOaA5tLLMjY2S9ReRphkeZqbnvmqvaZbEViQ308fU7zL0FyjKUZm5DwC9b93M8SsWCBXMrgLGU49PKXsxmSYwmaSuAyjlX"
                    "NbFc9mVzL1a2RwB1FcT9TGBxOQOJ8lYuq8/p417TSPai5Nh127FalvY7llOl59yrtY7byHGnG3ZdTcQ1MXOO2lzQza+1l1F844Hyj06z1njvK8bLj70DnGGX"
                    "vJSyFYOIju460AGWqsw05rh0ywuJgtRV7McsEd53p16k8O9VeM0+WcJzEOVxdpobMwf3V7GXGjU+NytJx96jfqyB0U0MgLXFhlryT13xzPzgMOWdoCICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICLHXKHKuE8P4w9lecWqa+EbqoVXAZJL9xkNucWTMZpKCv721z7J6NDlSlka2YdfXxJtvby66mr7Cx"
                    "i5JxXiec5nlo8PgabrVgtE1mZxMdTH1BJHFJev2CC2vVjklijDiHSzzyw1KsVi5Yr15cF9RPUfiHpZxm1yvmmWixeLruEFeP/tb+VyEkcklbFYik0iW/kbQi"
                    "kMUEWmxxRzWrUlelXs2YYQ+cOd866hMoYKSTkLHayXNdxPEoT7n7dUNSG0RV2to59jNnfrhpcYevpbKTq4c6zqKBuuhWt0m8nt6f+m3HvTfFmbbLWXsQwjK5"
                    "eaNpmnkbuQ06LNOfWoNkcCypE8m5LDBcvOmkr02UKRPX36i+ZevWdNSRs2G4Tj7ssvH+JQzFzAY3dMWWz0sZbHfyz44w9rnNdXxDZp6eNGprlzI9OUjGuMcY"
                    "tsvyRxa49XGQU2RHWlM60lyFKOvxjHikpcbbn20lv6eKbxIbMm3Le3VHqax92BkXXlOV5WlhcW1omtyu/TxSAmCpCwAWcrkjG5jnV6UR9yYMJd9zaVMSW7UT"
                    "bGocfUo42tPk8g8srVGNfdnjLhJIXEmDHUWmOUG1ckb7Nfcbx1B1u0G1Kkz4I0s5zO5zi/n5dk0hJzJZJajw2lvKr6Wpim4dfRU7b6lKj1dcl10kpI0uzJr8"
                    "23nqft7OzmyZO4LCUeP46vhsXHqCEl8s7mx/qb1yUNFnIXnxhrZLdksaCdFkFdkFOsGU6lWGPXWRylrNXJL1t3hja9Su0uMFOnG6V0FOu1xOoo3SOc5x+6ee"
                    "Sa1K6SzNNI+bb0hOBYMTFco6mbhtmTbZc5b8eYC8zMkPsQsWo7hpnOZxNtrYhfXW+a0TOMzIsqJNl0T3Hsv9ts2Gslu4S69frD9RJcryal6eUZtY3jYiyWWa"
                    "zp1Yzd2t1VWPILyW4/G2OuMtcwPfkpmyx9cERbcH/s9fSJnHODZT1XytV7M1zuR2Pwn6mIslp8VxdqVhlg6jsMzuTidac/QE1ShjZGfaSXzUiGCsVQEQEQEQ"
                    "EQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEWNOZOKsU5w4uzjifNYbMzHM4oJdNKU5DhTn6uafZKpcjqmrGPKiM5"
                    "Bit5GrclxuwWwt2qyCprbOMaJMRlxHu8X5DkeJ8iwnJsTJ7eRweTp5OqS57GPkqzMlMExjcx7q1ljXV7MYcBLXlljdtryFjvLuM4rmfF+QcTzcLbGJ5FiL+H"
                    "vRlrXEQXq0ld0sXWCGzwF4nrya6op445GEOYCKEPPXEWRcYZ7m/GWX1z1ZkmHZBaUM1p2DJgtKlVsx6K1PgsyzU69TWTbf11HZkp6PcUUyBcQXn4c+O67fbw"
                    "Xl2N5xxbCcrw72y0M1j612MCRkr4Xyt1YqTOYGhtmnYbLWst00xWYJY3NBaQv5z+UcazPpjzrkXDc2x7cjxrMWcXac5hZ+srNcJaeQhB8wZPHzU8lWI+0xWY"
                    "j1O31LDHGueIwC0nVGSLklguRyWEZIhuO7Ok43ZR0rjwM1qoTDUiXJm0qFGzd1UIicybHvqa/wBt+0h487A9jL4+Sdkd+hG12UpMcYWFwjjvQOc10+Mne7pY"
                    "1s/T1Vp3j/c7Yjn2IZLcc/qXcdByKjHWBjFqJhfjrBMUYe57eptOWeR0bIq9lx3HJK8R1Z3e450UUtmQ7MW9c5XSXob647qkJacakxHm5MGdElNokwrKultd"
                    "zE2ssYbzE6umsOLYlwn2JLSloWlR9OjbiuwRWYmvDJN7jkBZJDJG9zJoJ43gPiswzMfFNE5ofFMxzHNDgtQSxS15XxTNdG+N5Y9jgQ9rmEhzXAnbXAghwIDt"
                    "j4JGshcE9QHKPTVnsTkXim/VT3cVl+LKhy23J1DfVkg0KkUmR06JMRu4p5TrMZ56Gp6NIZfYj2FTOqLuJX3MLEPUf0v4l6rYN+B5VQM0bCZaN+uWxZLF2NaF"
                    "mhZcxxik6SWyMkbNXmbtk8ErPsWyvSb1h5v6K8nj5NwrJCvJKI4MvibbXzYbkFJji/8AR5Wm1zPc6NudUuV5IbtGZ7pIJvafYrz22+j7rt4j6vqd2LQKVhvJ"
                    "9RXlY5NxddTkSrSJBKQUV24xu1+lr4+X48xKWwxMnwIcWwpXZlWnJKajK8oTtKgPWD0N5j6O5JsWYi/iWAtymPFcnpQvbj7jukvFazGXSuxuQ6Gueac8rxK1"
                    "kr6di3HDLIy9b0I+ongvrzhH28DN/CeSUI2HPcPv2IpMpiy49Is15GtiblMTK/7a+TrwxtJ1Fbr07W67d2xpdb9QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQ"
                    "EQEQEQEQEQEQEQEWCOeOoDD+B8dbsbojuMktCdaxnEIcpuPPuJCELNUqZIU3JOnx+GtKSs7x2JKKObjUSBCtLiXX1M3P/T/07zfqDlDTxw/S4+sWPymYmic+"
                    "tRhc4AMYzqiFu/KCf0tFk0Tpel8s01WnDZtwaW9bPXThvodxz+McilN/MXhLHx7i1KaNuWztqNu3FgcH/o8XWJacjl5on16bXxxMjtXrFOjahOz7kbkfnnKV"
                    "3uW2iZjkdDzECNHYXExzFa+YuG5JgUsFT8lUduW9Civv+7KlW9u7DhKs58pusiOQp58Y4pxj05wzaGLre2+X2327EjhJksvZhbM2Oxbl6WgiETzMhEccVSnH"
                    "LK2CFstmb9RSH6q+rvOfXTk7s5ym8RBWdOzCYSu6RuC41QsPje6rQru0ZLU7YK7r+Qn672SmgidK+KrVqVaXYaKjgVcdbaHo8aOww9MsrOwfZjsMxYEZ2TMn"
                    "2M15SI8WDBjIfkvuuuNxYMRDzzhpT7rh9HIZCxZkDnMklkkfHDVqVo3yvfNZlZDDWqwMDpJrFmYxxRMY181iZ0cbAXFjVjOMxreuOCEsaXBz5p5nMjYyONj5"
                    "JrFmd7mxw168TZJpZZC2KCFkkz3Bgc4R8c5csr5OyNEanOQxguNrlR8VjSGFQ5Fm68aWrDLraK+yzMZsLpLbaIVbMNKqCkbj15sR7WTkcmzkXwLh54njHy3m"
                    "xychyjYpMvLFI2eGpGzb6+GqSxvfBJXoF7jZtQdsledLZbI+ozGQ1cI5Vm4svaix+NkP8CxUr203Oi9uXIWndDbOVsh+pAbJjYKtVxP6KmIoT1Wn3rNvGvHX"
                    "HOTcycn4Rw/hna5kOb3kCqafNcMmqqKtL820u5TU1+M3Kg4xRRLLLLaAytVhPp6adGq48uxejw3vU5zy7H8B4dneZZMbhxVKZ9aIiQ/q7shZWpVWFgc5puXZ"
                    "YKrJRpkJk92RzWMLhkHpT6dZH1X9ROJ+n+N6hJyDINiv2GNI/QYevFJazGRe4EBhpY6KxJF1bElo1oGn3Jmh1zDB8OpOPMMxLAsZZdj47hWN0mKUbMh45Ehu"
                    "px+tjVVemTIUSVSZP0sVo5EhZEuQ8bjyy7lmKTMtk7maymRzGQk96/lb1vI3ZQ0NElq7PJZne1g7Ma6WRxaxv2sbprdABf0nYTDY7juGxPH8RXbUxWDxtHE4"
                    "2q3u2vQx1aKpUhBPd3twQsb1Hu4guOySu0Dz16iAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAirh+t5"
                    "0xlGssQ6n8Yr0pjXpQ+P+S0Qq59WryDFkPYdlVnJhsLbaTZUsN/ErC2t3o7DDtBg1FCU9Lt2misT+h71O6JMz6X5W0elzZM/xgTTEgFp1mcZWY47DiXsysMM"
                    "I76yk7/Herz/AGhXpJ1Q4H1lw1NokhdBxfmT4IvufXlLjx/L2S1uiK85lw80rz1EWsXFsshAZWdyynIyVLbSRmaiS+kk72eySh4zPWz3ptaj3v7N6Lu3ZLG7"
                    "fc70QD/kdeQenv8APxrYPhVw4HIFpFZ7ukgdUZ+Py+Pe+wHd7ANa28d+2smcLZom7YY4lu3GW7WG3J/0us5C2o6JaNu2EvjWVIdNtHfJeXJsMBVIV7yrN6wx"
                    "Fl9abHE6pnGctA7E2H5mPrNCzJH/ABmFrXPbA93TAzNRsa0kMaPbiyxH2iCOK+5rRDdnf6PKcJ/EqZzlQNdcgj3lIt/dNC3bW5Abdp8sYDIbgaASz2rha8/r"
                    "5295kF9yyWlTZpNSVIUlRGlSVGlSVIMu4lJUXYru1rt8mRloesw9gQ7qBAPUCCCNdiCOxaQARrq7Hej3Wq377jQ6h2Ogex8+B5+d+NnuVy+J5hlPHuT0uaYb"
                    "fWeN5NjVpEuqK6q3nGZtdZQHifjSG1ESm1pPSmZUd9DsSwgvSa6fHkwZUuK95XIOOYTlmJu4DkONq5bE5GF0Fylab1RSse0tJD2ESQTMJ64J6747EErRNBJH"
                    "Kxkjfe4ryvkXCc/jeT8Vy1vC57Eziehkab2iSJ4I3HLFI18FurKB02KduOapZj3HNE9uwrR3QP6oOI9RSKbirmRyswnnBf01dTz0GUTFeUXzae7DqiWRNY5m"
                    "KjjqTLxiQ6qvuXHYk7Eprz8+diWL1Oev/wBMWc9LZbfJeNttZrgTpS90zg2TJceEjmBsOUbHr36RfJ7dfJxsaBr2rscMjWTWrtfpm+rzj3rLFV4nysU+N+pc"
                    "UPam1xjxPKWRMc6S3gXyve6G2yNvXcw08rp4z1TU5LVbr/Ty6CJ6migIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgItQupPqvx3hWM7juOlX5PyS"
                    "8bCDplvLXX4uxLjOSWbPJfpnWXVueyTL8XHo8qLaTWZUOU+/V1ktiyXuX0v9IMrz2dmQve/iuMRF7n5DoDZ8k+GRsb6mLEjJGEl5cyW/LFJUruinY1tm1Eah"
                    "ij9Rv1T8X9EKUmFxgq8l9RrUURqcdZMTVwsNmN74sryaWF7H1awYGyV8ZHLHksl7sBiFWlK/IwQ0uvZZyVeTcly+/tbyznONOXeT230y7GzkspJvTDUOLCq2"
                    "HG0bJiFVV9fQUkY0wq2ugV7EKsKcsFfCcQxsGJwmPq0q8LHCpjqvWI4mvOzJYkkklsyh5A65bM81248e5NPNKZbApZ5LyblnqPyS9yvl+cu5rMZB7TcyVpwb"
                    "0hhcIqVGtEyKtSqV2uLa9KnBBSqNdqOJvUGHJ1dXMx2WocGN7bKPKG0fKlKL7nHFGXc68syIzdWZqWfanwlKEljNm0+V757Epe93Yud4A+GsHhjANgNH2jud"
                    "nqcTzrwNAbBBH0taW6a3ySdEucT3e4juXAdZIIb2AaNXepfltlLcriLFJDftxn2WuSreI42s5NlAcQ4WBw5Ed5xs4tVNQ2/mX9khy+hM4w4qO1TZFEttq+ln"
                    "DnPdHzTLscHSMkfxanK1zfarWA5h5DNHLE14mtQOfFhBsxtx08mUAlkvYyan1eaZv+EQP4pRDBflEX9pbjXffCYgJW8egMUhYBFMI582Xj3/ANfBFjR7LaV7"
                    "9fpg64iIy9OkF3MRmzdWjuIlKSjZ+02a9F3vKMm2+4zJTim0kZ7Mj3dozPZC0kOkdrenHR+XODfho252h2AJ7a0tZxN6WlwaABs/huyN9z8fGyfg6B7EqZD0"
                    "cOCZ0p3kHqdyqsdYKaqTx1xv9bWEyiSlL8edn2T1UySTxT69uUzT4TU2NacJyvsqjkeinKnE4lMSA31oeocdnI4T0zxVoOr4ZrczyBsMuwchYi6MXSnDSCJK"
                    "1R816SKQODv11OYacwaty/2dfpH/AA7Acg9ZcxUIu8mfJx3iT5gSYcDjbLm5fIV2uaGt/iuVhbUbMxzia+KLWOY2eZj55hBFWaoCICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICICICICICICLFfOHEeNc8cS57xFlyVlR5zQSalyU0nvkVNi241Pob+G2bjbbs/Hb6HW3sBp9Rxn"
                    "Ztcw3KbdjqdaXknD+U5ThPJ8HyvDSCPJYLIQ3q+/5JQwls9WUaJMFys+arYA04wzSBpDiCMV5xw/D8/4jyLhefhM+I5Ji7WLuNaemSNs7P7qzC7/AAWKdhsV"
                    "us/v0TwxvIOtGhVytx1f8cZrmXHmXwocPI8RyK8xfI4VfLmWNWxcUllKqbOPWz50Cpm2NV9XEklV2cisrnbOvOLYlDjlJShN9fDeT4/mHF8DyfFyB9PN4qlk"
                    "YATGZIm2YGSGvP7b3MbPWe50FhjXn254nxl226X84/LuLZb095nyHh2ZDm5Xi+YuYmewIzEy1+isFtbIwMJ7V8hWMF6r5/uZ4we5IGq+Q06okh1BKeadZNMi"
                    "JJYddYkNe2snIsll5lTbzEplxtDjUhhaHmHmScaUhaCUWWDTh0vDXNcHMe1wBY4PBDmuDttc1wd0lrgQQdEdxvJ8LlHPbDKzWieiRhax7Cd6ljcxwc18bwT1"
                    "RyBzJI39L2ua4g7T4jmhcqY7IuHybTyDi8Zs+QIbCOz98rVOsQIHJcSKS3DS1Pkux63N2mEoj12USI1i1Gg1uSVsRnFYWnCW48XLKTjbbiMJNI4AxPDHSzYV"
                    "7+3U6FjXTY0u++WkHwOfLLTllkw3l3HG46RuRoRPOMuyODCHGQUrJBkfSlcQCD0h0lIkvM1RhY6WSxXs9H6uGala2k/BmR/b+D2WtbPWy8GZ7Iy/Bnoe4NDw"
                    "d689gf8AUaAHwSCd63vwCsFI6e5afwdjudDudnZJ38aJG99ivkaekQpLEyE+uNJjrQ+xIaPtW243paFJUXnaFpI9mRl3EW/BeONivBcrzU7UUc9azG+GxBKx"
                    "skUsMgLZIpGPBa+KVpLJGO2x7C5rgWlwPZqWrFOxBbqzzVbdWaKxVs1pZIbNWxC8SQWas8TmzV7EEoZJBPC+OaGVjJYntkaCLBPQB6tyWU0vC/VjeF7DCY9V"
                    "iXOFlIaT9JBi16W4sDliS+tDklLS4qo6eRUqkzVpkxZGeNL+kyDPn62vqG+kebHvucz9KKE1mi90lrL8PrsdLPUc+Quks8ejb1vmqDZfJitiStpwx4dB7dKv"
                    "bX9MH1sw5ZuO4D60ZKCrldRUsFzuwYq1TKEajr0eTv2yGnk3johr5cNjq5KUBt0Vr0jJbtiRtxt5tDrS0OtOoS4242pK23G1pJSFoWkzStC0mSkqSZpUkyMj"
                    "MjFfRBBIIIIJBBGiCOxBB7gg9iD4VmQIIBBBBGwQdgj8gjsQv7H4v1ARARARARARARARARARARARARARARARARARARARRz9TfWpExVVngXDs6PYZTFflV19m"
                    "aGWJtTjs6FNXAsaekRIQ7FuL6DIjzIlnNUzKpqOewqpUmwumbaLQSW9KPQufOfpeQ8zglq4WWOOxRw5fJXvZOKeBs9W3aLeiSnjrEckU9dgfHcvV3tsx+xTl"
                    "qT3a/fqc+sujwQ5DgfpVaqZfnEMslLM8gEcd3D8TlifJFaqVw7qgynIYHsdDLC5r8fibG2Xf1dyvYxkcXFLQS7l521t5EpxiTIfnSJMyS9JsrqbNlOzZ8uTK"
                    "fcelyJNjNkyp1pcS3n7CynyJMp59+ZJkzUS9tXa2MhjoY2GCIwRsgihgjZHVpQwxtihjZExrYmthjZHFXqxtbBDC1jAxkTGQuqTkdkM3ft5fM3bd+5krM17I"
                    "ZC9PJYvZO3Yk96xZnszufPLJPKXPlsSPc9+9DYPVHk9ommkNsNJbaZaR2NstpShDSEkaSShCddqUnvXgz7tmZdyu4Ys8ve50jy573nqc97iXOcdElzjskkf6"
                    "jQ8ABeqwNa1rQA1gHYa0GtG9AAHeta/HfyuhczcnlxLisdqqkEjkTL4z5Ywj4dxyjQ7IhTs7fbNSDbfblsSqrDi7ybkZDEnWxlJiYpZVs33OE8V/tnmJDbaD"
                    "xnCSxnMvGunJ33Mjnr8dY7TtsfBJFbzZLeuPGzV6YMc2Xq2q/rZO/wD2Rw0OWIeM9lOocbgeyRrq8MTpYbPJ3NcwNdDTtRuo4jpc1s2XisWXPdHh5qluN6Og"
                    "3VGalqdMzPanFKecWotn3LcWa3HFq+VLcWpZ7NSzUozMShld0DTWhjQAAA1rGNG9aDWgNDRoaDQGgAADQWlGNc93W9znSP09z3OcS97u7nEuJcS4kkknqJJO"
                    "z1LjHccyTknOcF4cweLDmZdnGRUtTXR7RVizTuWdxPZhUTN1KpoVpawKKPIdO5ySxg1s16no45XaYzrUJ4i87McgocN4tyPm2We39DhMXcmY1jme9NLHA8mO"
                    "ASlsbrFiYRVKrS9ofM+Vju5WccD4TlvUbmfGOBYIdGT5Pl6mMbP3DKNWV7X3clISNiLHUWWLzx0uMjYBFGHSPjDrn3E3GmOcOca4TxfibCWaHCcfgUcRw20t"
                    "v2D0dvvsbqwNJqN+3v7NyZd3Mxxbj862sJs2Q66++44ukPkmfyHKc9l+RZWUzZDM37OQtPJ2A+xI54ij8dMMDC2CBgAbHDGyNoDWgD+lXifGcTwvjGA4lgq7"
                    "KuH45iaOHx8DGhobWo12QMc4DzJJ0GWVx2Xyve9xLnErIY8VZCgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgI"
                    "gIgIgIgIgIq1nrc9M6aXKcT6nMZrnU1+dGxhHJrqH4f0zOY0NOk8Lu1wzYKyU/f4bUWVHaWKpaqmA1geMQm4saxu3XrGx76HfU50jMt6W5Obq/TNnz/Gerr2"
                    "2vLI0Zug092FrLEsGQgiDRITYyMnU9jdR1Y/7Qr0jZE/Bes+IrAGR9bi3MPbaS5xLXnjuUcGjYLeibE2JXAhwdi2F7TG0ProZRVHIiLfSlJvRkuOaL5cZSX+"
                    "6gz8mZoIieQnez7VISWlkYsYa8B2v6D9gfJ12PjWu2x89j2NbmEuiGZkTyBHMQ3502TqPQQPjqLulxOuzuonpasU1N/d4VkdVl2NPMsXFO864ymS0qRAnxJL"
                    "Tkayp7eN3t/WU9xAekV1lF7m1uRZC1xn4s1qNKj8LlKDI1ZalnrMU3SeqN3tzQyRkSQWIHAEsnglbHLE7TgHsbtkkZc12y65r3K89G8wzVbTBFMwFrZG+DHN"
                    "C5zXe1ZhdqSCTRAe0skZJC+WKXb9UmkyOgqc9xJL5YvkLkiIqvkOnKn4lkkNDDtxhdvJJlhMmXVlJZkV1iTEZu9opVbdx4zLE5DZeJi7c7pJ8ZkehuUohhkL"
                    "AI479WQubBk68fU5zIbJjc2WIuf+nsMlrGR5j63ac5DhJ8JkJqcpMjBqatO37WWachf7FhgDnAe50OZI3ZdBYjmrSFssMjG8OvejP86MiIz+PBeO497Pwfgz"
                    "I9n+e0exoa/lJBOv69z8a+O2u5A/prfhjZ0Px3JOh/n+x1ob6dfIO9ri30q33J2SyMjSaVGlZKI+4jIy+4tKTtOtH47kKMz2OQ7AA71ruPI1oHRHggjY/cdy"
                    "NDS7MZ7kEgg7DmkAtI8Hq8g7Hk/5dz1BSxdAHqg5d0zSKzi/l566zjgIm0Qa1tvdhlXFxm80mPKxU3TS9Z4jHZJ9qfgzjxohRvpJeFO1j1dIxjLoY/UL9KeO"
                    "5625y/gUNXE81cZbOQxo6a+M5M/Rc9z+7YqGYkO3Nu6EN2Ult4CWQ3YrAvpj+srLenho8I9TbVvOcCaIaeLzThLczfEYx9kccrh1T5bARMDGfpgx1/GsHVUf"
                    "YrNbSjtd4Vm2I8j4rR5xgeRVOWYjkkJNhR5BRzGp1bYxVLWytTL7RmSXY8hp6JMjOk3KgzWJEKYyxLjvMt1VZTF5HC5C3icvRtY3J0J31rtG7C+varTxnTop"
                    "oZA17HDsRsac0hzSWkE3FYrK4zO42lmMNfqZTFZGvHboZChYitU7daUdUc1exC58Usbh4cxxGwR5BC7QOgvQQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEX5Pv"
                    "sRWHpUp5qNGjNOPyJD7iGWGGGUKcdeedcUltppptKluOLUlCEJUpSiSRmOccb5XsiiY+SSR7Y4442l75HvIaxjGNBc57nENa1oJcSAASVwlljgjkmmkZDDCx"
                    "8sssr2xxxRxtL5JJJHkNYxjQXPe4hrWgkkAEqJfqn6y5l+crA+Hb2VVUDap8LIs0rFuw7a8fYlFGKNjNqy8l2qx5KGJa3baGlNrkHvw5FZOqKSIteWzE9IvQ"
                    "uKp7HIub0YrVxzYJ8dgLLWy1qjZI/cbLlqrmltm8S+LpozbrUiyWG5DauSmLF1UfVB9aVi+bvAvRbLPrY+N09XP8+oSujs33RSuhlo8StxncFDbHCXkdcma8"
                    "17XYaavVazIXtF8fxlJpamWbKUMIbbRDrjbJolobSlLRvsklPsRG0IJtiH2pJaEoJxDcZJNPSJyOVDeqvTft5LvesNPV0uc4lzY3+HyuJ6pJgSGku6S6Ul8d"
                    "clOjsiWw3obrccJOi7WiDI3fZo7BsZ05ztdX2fa/vprUot9uiJJfae/xok6+P4Ii0ZeDMiSRGRjHQAPk7JPf8nuTv/PZ777+T2C9j43rv2I1/h6Ro6O/20Ce"
                    "wB7AnYXy3mSUWCYraZ3k7bjtRUrajRKuO99NMyS+ltPOU+N18tTEpuJLtTiyVuz1xZbdPUwbW+kRZMWqeYc+lHGZHkGXq8exLo2XrjZJZbczPdr4vHwuY27l"
                    "LUTJIXzQU/dia2u2aF125Yp46OaKW4yRvrU2Uqla1nc0JRiMYYxNBC9kM+RuSiR9PEU5nxzRxWsh7MwZM+CdtOtBcvur2mU315Iycrym7zrJrbLcifafur2W"
                    "h+UqO2bMOK0ywzDgVddHNThxaunro0Srq4puuuswIscpL8mUciU/KnD4jH8fxFPC42N8VGhE5kXuu9yeaR8r57Ny1IAz3bl2zLLbtyhjGSWZXujjjhEcUep8"
                    "vlbudytvK3nRme37bWxQM9uvXrwxsr1adePqcGVqleKKrXjL3OEMTOp8jw57/hefjU1dLtJ3d9LXRlSHibSZrX4JCGmtmenZT7jcVgjNKTfea342PoGSWpo6"
                    "8OhLPJ0RnY00/wAxe4BvdsbA6V5AJDGOGvhfKNoGj1DpaSX9z2b27b+7Q32J1oH9zoSe+jPwJKyvNc+6pMvguLRjcmViGCPE/GTXP5bkFWzIyuzjVym12sdW"
                    "L4fZVtDSTTllUzK3OryCbEywpkPVkMfrZ9Qo6tfj/pViJi1gihzfIGAOMhhje6PEVJpQWxvM1iKxkLUPt9bZoKc3U1smn2l/7On0m92bk/rTlqgIa6fiPDZJ"
                    "maczpLZOTZCAOHcF36PExzs00SRZauHOd7wbYpFdytZQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQ"
                    "EQEQEWHuoDhuj6geGuQuHsheZhwc4x6RWxbZ6siXRY/fRnWbPGMnaqpxoi2ErGMkg1OQRIjzjKH5Vay2b7Pd7qMq4Ry3I8E5bgOXYv7rmCyMF1sPuyQMtwNJ"
                    "jt0ZZYiJGQX6j56c5bs+zO/sfBw71A4VivUbhXJuD5sH+G8mxFvFzysYx8tV88Z/TXoA/sLNC02G7WdtpbPBG5rmkAiiByJgeR8d5dk+DZbUvUmS4leW2PXt"
                    "XI04qHb0k5+tsIyH2+6POYYmRX2Y9jAdfrrFlKJtfJlQ5DD7l83FOSYvl3HsPyTDTGxjc1Qr36kp013tWG9XTKzuY5Y3dTJY3adFIxzHAOaQv5weW8YzHBeV"
                    "Z3iGfgbXzHHcpYxWQiY7cZlgf9s0Lz3fXtwuit1ZHfz1p4ZPL1rFk9McGWtCEdsd1Kno2tf2me3GfGzSbKlEkt7M2lNKPyrZZS12/nx5Pfufk/PwfBOifzvR"
                    "ybC5D9RCC525WEMl77O+3TJsn/3gGyBvT/cH+FfzxbyG3xnkMxm6bkTOO8vREqs8rIsc5UuLHjm+dVmFPHbYdkvZBiLsqTKixoxG5cVUm0pOxUybXSoPl5fH"
                    "SXY4bdIsZlsaXy0JHuMbJmvDRYoWHk6FW8I2Nkcf+xmZFYALYnxvyjKYyLkmLfQPS2/AH2MVM98cbf1B6Q+lNJNNFXjgvhrY3TWHsjqWGwWXSR1222T7QXdU"
                    "qmnlCKVHsIkiLFsqm3gOtyK66prBgpVbb10ppRtSYE6KtD8d5o1NrQvaVGkiUrhj70d+o21G18bg90U9eRpZPVsRP6Jq87DpzJYXgsc0+HfAJK0PZqzVppIp"
                    "43xSwSPhlila6OSKWJ5Y+N7NAtc1wLXg/c1wIcATodfdIiPWjLWi2XaZfku4iPye97Itb0Rfjeu6Dv57gdhojtsdjo+N+R8n9x3+TQe57aH+eyfAP7DsO3+f"
                    "fRXFyNL/AAe9mWtJ0Z/50f8An58bMjPR/A+jSddh/h8nwXfJ79gdd/n414XZi20gHpHbR0R3/JAA86J1386HwSt0eizr35S6L8vdcpFLyvjDI5cFWdcbW0mS"
                    "VZM9l9pD2RY04l3txzNma9C69N2yxLjW8Eotfk9XeN0+MOY9H71y+nrjHrJjTZ1Dg+Y0onjFcihiBErdOLKGZjjYHXMc6R3U1wBtUnEyVZAySzXsSp+nb6le"
                    "Weh2T/Rl9jkHAb87ZMrxeack1JHdIkyXHZZXdFC/0D+9q7bj8hofqWxTNZcZcN6eeo/iTqh48r+SOIcmYu6qQhhm6p5CmI2UYbdOR0Pv43mFK3IkOU9zFSsl"
                    "J07Jq7aGbFzj1nc0E6utZlPvOOB8o9OuQW+Ncsxc2NyNZxMbnAvqXqxcRFex1oD2rlOcDqjmjO2ncUzIp2SRMu54H6gcS9S+O1OUcNzFbL4q0A1zoj0WqNkN"
                    "a6ahkqj9T0b9cuDZq1hjHt7Pb1xPZI7Oow9ZmgIgIgIgIgIgIgIgIgIgIgIgIgIgIuAyjKMfwugtMoym2iUlDTRVy7GxmrNLTLSdJShttCVvypUh1SI8KDFa"
                    "fmzpbrMOFHflPtMr7+LxeQzV+ri8VUmvX7srYa1aBvVJI87JJ3prI2NDpJppHMihia+WZ7I2PePIz2ew3GMPkM/yDJVMRhsVWkt5DI3pWw1q0EY2XPe7u57j"
                    "pkUTA+WeVzIYWSSvYx0KXUr1W5JzZOPF8ZKxxrjaLJeabp0PrasszdU4yUS0ypDCCWliObayqcRbcfr4rz7lrdqt7xugbxKdnpV6N4zhNdmazX6bJcnfGJDO"
                    "5rX1MIwNf7lfGl56XzOa7dzKuDZHta2rRbWpi7LlqZPqX+rnNerk9rh3Cn3+P+nUc0kEpDn18rzR3WxkVjKNZ0S08OHNc+ng2mR0/W23lXSzmrRxeBqPHEQz"
                    "bmWaSel//qNRlETjcVZ6US3FH4elIM9F29zLSy7kqdcSl1GzMhkzOHQ1T7cIPS+UEh8oAILWDf2QkefD3g6cGsc5joh1KIi1JP0vf5bH2LYj/wAZJOjIAACA"
                    "AwO8dZ0R2pX3eded6/yW9mevPzrx3fnyX9vkeQO3/wAv6f5D+g7fK9EA7/JII8juQT8HzvY7fPYjv2X2worbpvy5cqJX1kCPIn2tnYyERq+trIMdyZYWM6S6"
                    "ZNsQa+Kw9LlOqPaGGnFaM0mQ61md0TWRwwzWbU8sVepUrRmWzatWJGQ161eJv3PnsTPZDE0D7pHtHbYXo47Hy5CyyGMxRRkPlmsTu9qvVrQROltWrMh7R1qs"
                    "Ecliw8ghsTH9LXHW47OaOU3eUMkacr25cLCseRJg4ZUy2UxpP00lTJ2GRW8balM5Bka48aRLireeRUQI1XRNOyTrZU+ykpwbiLOJYx/6oxT57JmKxnLkL/di"
                    "bLCHitjKkp6evHYxss0cMwjjfcsy28g9kIsxV6uCcr5DFnLFenjmvhwOJ96LHRzsEVizLP7Zt5S5CySaNlu86OIey2adlSpBVotmnFc2Z8c1cU3CVINJaUak"
                    "tdxeC0n7lJPR+SMyLwZFsjL+TGT2pul3tgnsAXf57IHkAePwd91jsTNHex+3bsfk789zokHX9dFdPypi9zTL8O4nw+vet8jyW6p4TFXGNSJM++v5KKvFaM3V"
                    "kTDP1kqcy+6/LNEKOc+rmSX47DL7jf4cjj+M8dz/ADPMy/p8dhqFqd0pI0K1QCS3KzbgTI97BVha0+5JIJog1zpGA5NxjjWV5jyHAcQwUBsZnkuWpYmhEAdG"
                    "a3O2ISSHuRDCwvs2HBum14pnk/3biLmXT3wxRdPfDHH3D2POR5cTCqFmFPto9c3UJyHIpjr1nlOTrrGn5aK97JMkm2l27CTLlJhrnnFRJfQylxVJXOeX5Lnv"
                    "Ls/y7Kki5nMhNcMJldM2pX7R0qMUjwHOho044KkLi1pMcLSWgkhf0n+nfB8T6bcI4xwXBgnG8axNXGxTOY1ktyaJnVbyFhrSR+pyFt892wQXbmnf3I0VmUYm"
                    "s0QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEVaH1qemI8dz3HOpbFa1BVHI7LWK8hFCrXi"
                    "VFzzHak00d/YTkSZCZH9UYXXIqktnArY1Srj83nptnNydtiJZP8ARJ6ofqsblfS/LWgZcWX5njImlHVJQtT7ymOhjLBoU7srb0YD5Hv/AIjZIbHFV6lU7/tD"
                    "PSIU8lgfWPEVXfp8m2HjPLvaZqOG/A0nj+VlcHfzXK/v4qZ2mN9ynjWdTpJlXzyOpOxiLaSnTzRG7GPRJJLpEojQaiL+11vbR72WzS5ozbRqwtji3fckEfHf"
                    "v4A+Pkk9tfPbZ2q4MReNOwHucfbeQyUb2TH2+7Wt7YR1jXc6Ler7itd7OIlRaUlXbs9pURpPuJJkfd42Rn8GRkSt72X2kZ9th6Tvej42QSfuPbye3fz/AF1v"
                    "utt0rBB2CO3z5Gt9iO5Gvka7EEDffSzhwbmrdmzF4YyWW2wp2XIXxHfTnkttV99YvG7I46myX5KGmqbK5i3HMVSSTcg5dJXWNployGsjVmNZaGTFWX52qHuq"
                    "yADPVWN2fYYOluXiDI3PdPTadXNkB9Bge57BVcZfO5hgG5inJnKjQ6/Ug68pEwFxt060WhkGNAJ/UUYY2i69paJKLG2ZGMdVuWJ8jymno7z8eUhcaTGdUy/H"
                    "eSTbrD7SjS42tOtEtC0n3Fv5T4M07MvTjcyVrJInCSORokbI0dQkY5u2uBG9tPwdb1rtsAHThDmu/lJGvwD8aGtAfnW9f89AcS6Rl+fJlrZ/B/Gtn5PXnRl4"
                    "1vXn5H2B6T877fb+PIPYA9+/+fceF9WHWj0nz3Hf7QfOt77+f9NLiJKSURES9Ho9fak9FoyPey+C1/J+D2W9+foNePz576O9np12HjsCB4PjXdd2FxDgensd"
                    "fkDvojQ/bfn89j32ss9PXUxy30rcj1nJPFGRO1NnCkMlc0kgnpGN5jTNqkqexvLalp6Mi5pZJTJaijm9FnV8x1FxQWNJkMaBdQdcep3pVxL1X49YwPJqXWeh"
                    "8mMytcNGTw1ss+2zj53Ahn3NZ78Dg6vajBinY9h2t0+kXq1zD0h5LByTiN1zWvLIcthrT5H4fO02uHXTyNZr2/expd+jvR9NqjM4SQudGZIJ7lnRL15cS9am"
                    "Et2WNOMYlyZUQm3c34qsbaPOuKV1DcJMy1oZfsV72UYd9ZOYjRMjRVVkhC3ozF5T0dhJZgqpy9YPRTl/o7mjSzld9zCW5ZBhOS1oJWY3KRNJcI3F3UKeRjjA"
                    "dZx0sj5It9cUliuWTvu59GvXHhnrTg/4hgLLaebpRxfx7i9yaI5bDTSbaHuY0tNvGzyNe2llIGfp7PQ6NwhtRz1ot4xp5bmQEQEQEQEQEQEQEQEQEQEQEQEW"
                    "OOUeVsL4exhzKs2sVRIa3/oayBFQh+2vbVbD8lmqqIi3WEPzHWYz7y3JD8WBCjMvzrObCgR5EprJeKcSznM8qzEYKr79gs96xNIXMqUaoeyOS5cmax7o4I3y"
                    "Mbpkcs80j2Q1oZp5I4n4D6kepnD/AEo4zZ5XzTJtx+OieK1WCMNlyGWyMkcstfF4qqXxm1esMhlc1pfFBDFFNatz1qkE9iKDLm/qDz7qAyWIdj2QKOukyXMW"
                    "xCsccXX1K5URcSRLkS3m47tvcnBclQpF7Miwiar5U+PXwaaBaWcOXYB6femPHvTnGPkjAtZSxDG3KZmyB7tgRytlZBXhaXsqUxO2OSOpE+Z8s8UMk89qaCq+"
                    "CkD16+ozmnrvmGx3S7DcSx9t8vH+H1JDJDXf0PiGSy9gthOSzBgc9psSMjq0I5Zq9CvC2a5Nb6zR0LNUkn3VofsDRpb5oL243enSm4pLIlI8GaXJHaT7pGot"
                    "NNrW0Pdv5F9wmNgdHWDvti2eqUg7a6bRIcdgFsRJjYdEFzwHLSVaq2DT3gSWTvcmz9hIcHNjHho0SOrQc4E/aGkhc9862ZlryXnXlRFv+7Wi/wAkZ+Nfb+T8"
                    "3etkd/g9u+gf89n9u3f57ruka8+O+naB33J8eBvfj/PfZfRHYckuNsMkalLXoteSIjMtrMtH4SRmRmX+C+THzkkZEx0jyA1o+dA9hvQPzv4H9fjxziidLI2N"
                    "jS98mgB3/oHAb/HYfaSSdrVvqc5UZSuTw5iU/uhVkpr/AFIs4b5dlvkNfLbks4W05HeND1Zi0+MxIyVl1v3JGZxGapbcYsPfXcbW9KuIPe2LnGZg1PZjd/Za"
                    "nMw9VPH2InRSZ17JYtstZavJJFinMd0xYSZ91rpTmo2UflzbKfwSF3EcbI82po4xyy5E8tHu9bZo+NgxTFktelJFWsZf3WjrzMMdMwB+Fjt3NOmGly5DUdKj"
                    "71rNKj+3/bQRGa3D0Sf7E7PRkRbJKfBGRDdcjxFE6VwGmjfcn7nfDSdeXE+fPY9tg71e1vbTR5OgBs67Df761vQ/bfyQey2VhDoKibZS9Jg1URcg2lKNCpCk"
                    "aRGiNrJJq9+dKWxDjmfzIkNmfbszLyoopbtmGCMn3rMoZ1AbDA7vJIRoD24Yw+V4HiNju57LuxMBIaPAbvRPgAbOwBvWv+72+4613O83o29PBcq8vZZ1N5pC"
                    "bsKviya7BxRTyJ/0z3K+SVjxTZTDiJLcJ5OFYdZurOksY89mNIzPDb6AiDPoa2SmLH1qepLcJx7EeleFshkuZDMhnWxvi9yPB4+ZrcfUnaI+sOyWRgNl8rXR"
                    "yPGMljkMkVqQOss/2efpCcryTP8ArFmKpNDjzJeO8R96Ee3Pmb0TJMvk4XD7TJjMc+OjG7TiZMvbd1e5ECLOQrOVuyAiAiAiAiAiAiAiAiAiAiAiAiAiAiAi"
                    "AiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiwl1G8J4/1E8KchcP5EhlMfL6J5mpnvLsEIo8prnGrXEchUVZMgTJLVFksGrs5VYU"
                    "pMO6hxpNLaszKmwnQpGX8B5jkuAcw4/y/FPc21hchFZfG3o1bpu3DkKLy9rmiO9RlsVHv0HxiYyROZKxj24R6kcFxHqXwbk3Bs5GH0ORYqxS9zR9ylc0Jcdk"
                    "oCO4s4zIR1r9c6c33q7A9j2FzHUVczxO0xTIL3F7yEqsvccurjHL2qcfTJdqL6hspFReU0mShtpDsymtoc2rmrS2hBTYT5EXb5F8fG87S5LhMTnsfK2ahlqF"
                    "PI05mdvcrXa7bEL+nqPS4xvbth25rttOi0r+bnknH8pxDkec4tmoX18vx7K3sPkYnNcxv6rH2JK8ske2ncExYJ6ziXCSvLFIHkO2Nac3pTiTPrG0d0ew71LM"
                    "iPSJyfLqTVoySUhJnIQRH3LWUkyLsSRDJInAjpDt676JIJaPBOvx/i+e3buVlPHMgJoP07n6mq9IaO33VidRuAGifZdqJx8NaYe/U7axDZRe5LiSNST0akqS"
                    "pbLqDTo0rbcbUlxDha7kuNKQ42vS0KQtJKT9/wCbYIJBHhxaQdb2NdxrWho9jrpJ0QFsGjYMbo3gjYII2GvbvfgtcHMc07/leHMcNhzXNJB3MwzNneZsSkTr"
                    "BaHuVMFhx05s2hJpkZri6Vog1nIMZg3HXXrKIv2avOFtEbCLd6JZ6romR0NSxiEEf9nr7cc5wbh8g9zsS52gylZPXJLiS49IDHt6paDez/ZY+IB7qtid+t+a"
                    "8bixc8WRx8Lm4nISSdDR9zaF4NbJLQkk6GMZC9pfPiw7qkfWZJC6WzPQu2Hcc6RbL5MjPZfJbLRGXjfjfxsz+NmRfgZHo7+N9yd7J86/BB7+fJ7j/LBG7Hce"
                    "Afnz32Ok6bvXbt28kE/C4x8jLuToy8eS/u3/AJ34/Pj5M/8AG/A56PxrWu2t7138eNHue+j3Hcb7rtx9yO5Ozsdj8D8d99j5I7A9upcLMZS6g0KT2qItpUXl"
                    "Rfzvz8GXykyMvjRFtJj9+R4Hf87J7nwCO3cDvvxrWwO3o1pHRu6h4JGwfH/F+/cfkd/k+CF+mF59nfEuY0Wfcc5RdYTm2MTTscdyiglHDs6uYqNIhuOR1KJx"
                    "h+PKgzJtfZV05iXV3FVNn0t1AsKixsIEnxOT8X49zPCXeO8mxlXL4a/H0WaVljXt20gxzQuIElezBI0SV7EL2TQSMD43AjZ2Fw7lue4dnaHJuK5e3h81jn9U"
                    "Fuo8xu9t5b79WzGP7uzSstYI7VSdr4ZmAdTeprHMt2+nD6q+H9VzFVxHzC7S4N1Fsw5JwksEVViHK7MA2zckYgUyW/8At+YtRHCl2mDKlyX7CJFssjxNybU1"
                    "+RVeIVG/UF9M+b9JLEvIcE6xnOAWrPRBe6TJkMFJKC6OnmmRsDfZJBZWyjWtgmPRDYbXsvjZNcf9PX1NYL1frRYDNtgwPqBUrCSxjy8MoZ2OM9D72CfI7re8"
                    "dn28Y/dmr1GSMz1AJhMYIrqVKAiAiAiAiAiAiAiAiAiAiwDz31C4hwRQfV2nbdZVYNGeP4jElNMTZxqU42ixsnjS8qpoGHmnEyrRcaQ4tTTsash2U9JQ1bE9"
                    "PPTfN+oWT/T0gaeLru3kszNE59es0BrjBXaCwW8hI1zfZqMkYB1tlsy1qwfO3R/rh69cN9DePDJZ2UZDPX2SM47xWnNGMnmJ2tf/AHrwes0MVA5jv1mUnjdF"
                    "F0+xXZavSV6c0G+c8hci815Ou8y67VeWjiST/tsqrcfooanTc+kpqhEmc3S1DSjUiHFVMs7WSmOwu2uMgujmXEyf/GuK8Z4BiP0OIp/pIASXSSObYyWQm0QJ"
                    "Ldv2q7rdjpIEj2xVqkXW/wDT1KNYx1o6N/U31Z5160cnPIeYZAWp4/dixuOrNmr4PAUpXMJqYuk6ab9NG/22NmmkksXrz42vtWbMsbXN5OnpYtMyaGtOyXSL"
                    "6uYsiJ14yV3dqC7le0ylSdtsJUej0pxbjhmsfG7fmvP6n/ZE0n2YGk9DO2tu7DrkI7OeWj/uta37ViNetFAHBv3vf3fIQGlxHfWt6YwHuG/0J2SCuZ1rZko9"
                    "6P8AP5/BHsvPgtmRmWv4I/A6O9/Hz8+fHn/U6+f9Dtdot14JHcEjt4Pfe9a2PnR7bIO+6J8mW/J7PWtduz1v/wCdEfhJl58/AHsD/Tz89v8A5H5X47ZI2e53"
                    "2O+4Gu5PYAa7DYGw3z2G+i8y8onw7irEWjdQnkfM4b5488pRk9itCanoUrNjbSpt1E4pbUmtwkzMml30Wddu/VRsUk1Nj7nCuKHm2Yc++wnjGEnjGUb/AIMt"
                    "kNR2IsAHElhgdC+KznfLhjpYKLBFLl4rlb2beSZw3Etyg6hyDKxSN44x+w6nCHyV5+Rlpb1E1Z2Sw4R0ZDZMrDNbdKY8U+rbjM2kiMiMzItEfcpS1H+e5S1m"
                    "pS1KM9rWtSlLV96jNZmYlOO52exHcdPSAABoBrQNNA19oboAdhpugtJN6g0lxcXdTi5zj1Hqc4v33Icdknehv9wQQe20cT22CmOFt2Sn/ZIy2aY/cRpMlfg5"
                    "BkTpmZ6JCWvgx5V6bqd7I7MjJLhvt1hvcHx/J3b22SS4He9L7xN0A7XkDXYA+NDtoEE6J150ABo/ccW8jyLHKclxri7GY7trd21pTtFVQpMVqVZ395JarcXx"
                    "xhUx2LEKwnPTo6obEqShmZNtaZslJeQgx6NCxTwWHy3LMrI2tj8bTuzvsStcRFSoRPsZC03pJLgPYdF2HUPZlaAWyd8gwmDyfIctiOP4eB9vL8hyNPEYusGu"
                    "3LdyNmOpWaSB9sRmla2abXTDEJ5JHMjjcVcj6UeBoPTTwDx1xBHlfuNjjtP9TlNqUlyW1aZjdPOW2USYEh6JBkqo2LeXJrsYjyoyJVfi8GmrZCnXIanV0l+p"
                    "fOLvqNznkXMbrTGcvfkkqVyA39JjYQK+NqENc5pkgpRQtne0kTWPemPeQr+kT0k9OsX6UenXFeBYrofDgcayO3aa0N/iGXtPfczOScNAh2Qydi1aDTssZK2P"
                    "ZDAVsSMFWxkBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBFWY9Z/pz/pLlTGOoKhrlIo"
                    "OWmEY5mEo5rKo8bkjGahLVd/4OU6Uon8rwSrbehRahpytY/09yi1tEQ7G2RItbLfol9SRew2Y9OMjZc63hHOzGCY8dnYa3OBfrRPHj9HlJ22C1+3PGTAiBZA"
                    "9rakv9oZ6TOx2ewHq9iKbW0s8yPjnLJIY/ubmacLnYPIWNENAu42KXGumcDuTHUYXO6pow6BvIqJu1gyYKjShx1KlxluEZIYlNH3MLWaSNRtk4ZoeItqUwt5"
                    "CS2otz9ikLdHyNjqII772NDwfGidD5Gz2Oq4cdffRtxWWgljSWysboufA4hsrBvTessAMfUdCURu79K1gsIy21utuMm040txp1tZbW062pSHG1JSXg0KLtV5"
                    "Mtp2ky0WvR2NdjvY8k7BHYg9yf69IHzrWvO4aczXtjex4e1zRIx7SdOY9vUx7d9y17SHb0N9XcfC4alyG+wbKqbNMYktxrugkfUxyeJxcGwYWh2PYVFqw2po"
                    "5lNcQXX66yipUlw47ynoq405iJMY61ypXyFOanZa4xztAJaQHxOaWuinicev25oZWskhf4DmtBBjc9pyJkdW9RtYy+wzU7zPZna32/eiO2uis1nva5sVurK1"
                    "s9aTpI6me3KJK8k0Mm6s5VDlWNVfJWFNuoxLInnI8qqcUl2bheUMNNOXOH2i2WIzC3q959Cq+Yw01Gn1rsGwittwLKtXK8PFX5zLJicloZOiwOdINNbkKhc5"
                    "sF+Frnvd0yNaWyMc57mStfGXyPZMWaOzmFt4LIy4+0WyOj6Ja9hjSIbtSTq/T3YA89YjmaxzXxyf3kEzJa8vTNBM1nSnv+Wi1stpI/HjezPR/wAeNn/8ePI9"
                    "3voHfnXg9x2J/oBoDufGgfPZedF3I+SP/jrfn99aA151rW+KeMjIy8a8edF+defPwXwX/Y9q1rQ/QDtxA0enWh4B32/5Heyfzo72F3oxr8jR389/Pz587/0H"
                    "ztcRLaQ4jyXgy/8AgvJkWy1/JpMi8no/zsj5Dt2G9f0Hknfz3HY//HYC78MjmuBBGx/X99+R2Pb5G/we/brKlTKidHsa6VJgTYUhidBnQn3I0yDNivtSIsqL"
                    "JZUl6NKjSGm5EaQypLsd9pt5taVNoMcLMFe7WmqW4IbVSzFJBarWY45q9mCaNzJYJ4ZGujmila9zHxvBY9h6S07KynF5CxVmrW6liarcpzRWK1mtK+CxWsRP"
                    "EkViCWNzJIpoXgOjlje17HjbHBwBbZk9Nf1loFkxScE9ZGWMwrVhEeswrnq/eQxDtEOS1sQaHlWz0mPEmQo7kWND5Isfp41hCjqf5Bmou2J2YZLWR9RH0j2M"
                    "E3Jc59Laj7WEY99rL8Pg9ya9iIizrmuYRpDpLuNY8PdLj2uku0mOaa7bFVrxVtM+nX6toORHH8I9UrUNTOvEVTEcul9uChmpAGsiq5ojoioZaYkCO21kdC88"
                    "H3HVLUkUM9lAjJREpJkpKiI0qIyMjIy2RkZeDIy8kZeDIQDIIJBGiOxB8g/gqfHnwv8AQRARARARARARARARab9SvVxQcLIexbFmIeU8kOsIcXCdW4uixZp5"
                    "SEtSciejOsuyrFTSzlxMXgyY89+MhuRaTqGFPqps/dfpd6N5TnkjMnknzYnjDHkG2GhtzKFm+uHGNljkYyIOHtS5KaOStFJ1sgivTwz14olfUd9VnGfRKrJg"
                    "sO2pyb1EtQh0GFEznY/AxTB3tZHkstd7JY2nXXVxFeWPI3gWOc+hTkbfEN8udlfJV5OyHKLmyt582Qt62v7N1L8qXIUol+0wlCGoyVpSs0x4MOPEp6aIlqBW"
                    "woNZGrqtM5K1PDcRxtfF4ijXp14GdNShXaWMjZ4MkrnF8paS0F80z5bduUumnllnfPZFLfJ+T8m9ReQX+U8rzFzMZjIvD72RuFuy4DTK1OBjWwVasAIbWp1Y"
                    "4qlSECOKKNgZG7usGFGro6IsRlLLJGSl6PvW45rS3X3f7nHVbNO1F9qSJCUpbQhCfFsTy2ZXSzPMkh2Bsaa1o1prGeGtGgdA93bcXFzi49eGNkMYjja1oJ6v"
                    "u3su87ds7c4t2Qf8IAY0BvS0fVrRGZkRHrfnwez1/wBvP8+NGRl5Iz8fHe9a/P8AUEf6/wDx38a2vo07P/CCQfnQ8jWgTregTo61533C/wAPXk/BkZaIv5I9"
                    "bP4I/OiIvOy2Wy1vbyfkfPn+vbz5/wAjv4761+kjZJHToj8ee5HyBsk+NeBoHZ7/AIXGRUuCYrb59kzTj1RTe0xDrWHyiS8lvpKHlVGM18tUaamLKtHI0hT0"
                    "5yJKbqqqJaXb0Z+LVSUDnTxt/kGYpccxLmMu3uuSazJGZocVjojGLuWswiWuZYajZIwyu2aF9y5NUx7JY5rkTh61NlKrWt5zLtl/hOMYwzNieYpL9uYSfpMT"
                    "UsGGZsVzIOjkayd0MwrVobd72J21HxOi8y3LbzOMit8ryOQ3Jur2V9VMUw2TEGKhplESFW1kU1LKHU1EBiLW1UMjW4xXRWEOuyJCXpDkrMLhsfx/F0cNi4nR"
                    "UcfF7UIlk92eZz3Omns25vtM9y5ZkltXJvEliaQsayMtiGqMvlbmcytvLXva/UWiC2GCP2oa0ELGw1alaFpLYatSCKOvDE3syNjR3O3HhIEM58xEYy/20l7s"
                    "hReDQw2Zd/3HozUtRpaQf3GRuErtNJLMu/PL+nhMu9OJLWDv3e7Yb28dg0vcB8N15IJ85o6ntb37HbvH8o12IGt6J0SfGtnZaQu23txCxyls72clKodRDckO"
                    "R0uNsKluJJLcSvYcNK0NPT5bkeAwskKS06+hSiNCVEXjV60tyzWpwkiSzK2Nkjm9ftB2zLNINtL2wRh072kjqEZ7hxBPfhiMr2tBA35cOrpa1oBc49Ic/paN"
                    "k/zHQJ127bV+jz082vMfULddQeYxHZGKcRqdtor0iFGcrMg5SyFuQxjsRg5cWUw9HwytO2ymRGiHX2ePXjXGFjFklDfcYfjh9aXqJX45wzEem2HnEV7kftz5"
                    "KON/99X45QeA2OQsc0sdk8hFHG5zg4WYqt9juzndVjH+z99JhyPm2W9VcrT9zD8KjdiuOOlaTHLynIVtWbDGvb0SOxGFs/a9pf0WsrHKeiavGRanFXSuHQEQ"
                    "EQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEWCupbhKp6iOD+Q+JLRUaO/lFE+WPWkqOm"
                    "S3RZdXGmxxW8Ns0m4tmuvYsJ6cwytl2bWfW1/vNty3DGZ+nvM7/p/wAzwHLsc5/u4i/FLPC1xYLlCTcOQpPII+21TkmhBOwx7mSAFzGrAvVDgOK9T+A8n4Ll"
                    "2t/S5/GTVYZ3MD30MgzU+MyUIPiahfir2mdJaXGIs6g17lR2y3ErzE8gvcYyOqlUt9jlza0F3UT1Rjm1V3ST5FZb1ctUGTMgLlVthFlQJSoMuZCVIYc+mlSY"
                    "/tPOXt4TOYzkGKx2axFtlzG5WlWv0rUbXsZPTtwtnglDZGRyM64ntcWSNbIzZbIxsgc0fzd8h4/l+K53L8bz9M4/N4LIWsTlaZcH+xeqSvhmYyRo/vYi6Prg"
                    "na0RywOjlYAx41rHyZQFEltW7KNMWR/Ty0ESi7bFpozSs/lCUzorJq0WiN+JIcURuSE7yOvJ1Mcw9+nRG/8AgO/2/mY463s9iO5IKyfieR92F9F53LUHuQnT"
                    "QTVfJotAHdxrzSNG9bMdiJo+yI7wlNZV3GRkrWjLZEozI/OvHwez0RfjW9nodo+e2u+zogDYJ7kH/n3Hnv8Asdi1ZB0+NeDv8a8E78fPY/v47hdr4h5Oa4py"
                    "aWV3FftOPMwRGq8+pWWXZLxw2VuFAympZaZdlrv8YVJfksR4aW3ravcm1aFInvU86t8TNYyW9DHPRk9jJ0DJNRm6ulvU4N66kx6g0Q2QwMc54cIn+3K5ro2z"
                    "RSdrNYWLlGMFMmOPI1BJNirUj44mtmeGmSlNLLJHFHTvGNrHyTu6KlhsVrqbALsdjaHLscOgnNNMy2LWpsYce1x68hvNSIN7Rz2GpNdZRJEc3I7rciK/HdS4"
                    "0tSHmnWZbBKhyorzvLGZBuRqe8Gviljc6CzWkaWvrWI3Fk0Tmu08Fkoe3Th9rmuY4iRkjW6EfFLDNNDJDJDNDK+GeGVj4poJY3uZLDJE8B7HxvD43B4BY5ro"
                    "3/exzR0Z4k+d/BkR93yfz5Pt0eyPR+SL5Pf3Hsh6IPbQ2d+e3Y9gCCfjsOx7dtb7nS+8Z3rXxsADt+SAT489gPGgD+y450i1stkZfhJmevBaP4/tMyPRGevw"
                    "Rl5H6Dp3Yj/x8bPfewdD5P41+V24+299h23/AJb3/pv/AJrh5LaVpNJlstH4JPky142Rdutfz5Ij8edGOXV338+PnYIA8fI14Oh+/wAL04XlpaQST/np2u/7"
                    "77a87PjvobHSbGG4wonEGo2yMlEojUa2lJPZKLWlJSlR7JSO09nvRF2qPmw/IHkggfkjXffbvrWjrsN712CyalZZI3pcAHaI76AeCOwPxt2zsEHxoknamt9M"
                    "/wBXq+6ZU0XB/UFIuMs6fIiK+mxi8isybfJ+GKqJFeiRYlNWsNvz8iwGvS3XRkYbCI5+M1Ed4sEiTG4NfhFlB76hvpKo8zNzmHplUqYrlWpLGS4+x1eniuRS"
                    "l/XJNUJMUGMyrwXk76aN2VzGzGnIZLL55/T79VuQ4gKPEPUu1ayvF+plfH8llM1vK8eZ0hkMF/TZJsniG9m+85zshQYCQLsGmV7f+GZriPI2K0Wc4FktJmOH"
                    "ZPXs2uPZNjllFt6S4r39k3KgWEJ16NIb70rac7Fmpl5t1h1KHmnEJq0yWNyGHv28XlaVrHZGjM+vco3YJK1qtPGdPimgla2SN47HTmjYIcNggmzzH5CjlaVX"
                    "JYy5WyGPuwssVLtOaOxWswSDbJYZonOjkY4eHNcRvYPcELs46S7iAiAiAiAiAijZ6outT+mZMzj3hubEk3TJz4WS56j2ZsegnxJKIjlHjMR5l+FZ27TiZhW9"
                    "3M92qo3orNVBhXlrLtH8Pk96SehUmabByLmlaaDFObBYx2DeX158pDNGZo7d+Rj456tF7TE6tWiMdq9FKbL5qlVlcZOvT6nvrKqcKdc4H6UXauQ5dHNZo57l"
                    "LGx3MdxaeBwimoY9kkclTJZ5kvuxWHyCWhiJYJILEVu97sFGMOtpZ2QPuW1u/Ncjyn5Mx6RIeecn20qbKdnzpjsp9Tklz66ZJfmzrR5TsmwlyJEg3XJEh6am"
                    "W89ypiYY6GPhrRvgZHDHFFHGytSihjbBDE2KMNia6GKNkUNZobHDGyNhaGRthdU3M7I5q5bymVuXL81+zNcuXb1ma1eyVqy8zWLNmxO6SaWSxK90s9qV7pZ5"
                    "Huk6i9znjIrLCGmkssNpaZaR2NNtESW2yIzM0oQWiLyalGf9ylLUpW1mahjckjnvL5Hl73nb3uO3OJ7bcT3PgD4AAAGgAF6DQGsDWBrWtb0ta0dLQAd9IHfY"
                    "G9k6+7qcST33+xJ8F48aPetERmf3GZn8EXkvnZ/Hn5McCdHudnfyPH+HQG9nwfH7/sDy12ILRrRHz2AOvn8DW9a8fGtD+D+T8nojSWj3+d+NGfx4PZ/k9+T/"
                    "AD+j4/J7/wBdaG+3/wAvHYL8HkjQHfs7q/5DYI2dho1rzs6PdfbCholOurkyY8GvhRZNjZ2E2Q3Fg19fAjOy502ZKeUlmLDhRWH5MyU8ttmJDYkSHlIQyrfX"
                    "sWDAxvtxSWLM0sdarVgjdNYs2bErIa8FeFgMk0880kcUEMYMk08kUbAXPC72OoyZG0yCN0cbQ2SWaaWQQQQQV4nzzzTzPcyOGCCGOSxZne4sggilnkd7cR1H"
                    "dznywXJ2UsHUk9FwfFUyq3DITzLkaRJZkrYTZZRZR3o0aTHs8jXEiOlAlIJ2mqY1XUqSU1m0lWMkeA8QPFMS91wsl5BlzDazk8bxLFC+IPNTE1pWSSRSVcW2"
                    "aVhsxO6L1yW3cB9iSrBWwrl2br5i6ylj2FmExIfXx5e0MluSye2LuVnaWNfHLekhidDC/qdSqRVqrnSzRT2bGD1rJG1GZEf3HtR/4NXn4+dFv+0teDIyLznb"
                    "Rs+Sf+EAE/JHfx4O/Hz2HdYr2HfsBr52d735I1rX8p128nxvffqGvOFAN11CvfmEiQ7syPsbJJ/Ss/CTSpCFrcWnwpDzymzI+wjHh3bHvTBodtkO2M89z3Ej"
                    "/PyWta3QIc1jTvThvsRMIBJ1tw8aGyANN3sgggHevzrvvWsG812sy1sqTj6mhz7OU6/FsJ9bVQJFraWVpMT7WPUlbVV7cizsLF1l9cxiqiRnpNo/ZUiIbTst"
                    "tCR7vH/0uOqZHkWRsR1KdSvOHWZnuigqU6jTNkLk0r9RxxM9v2nSOP8AdivaDiGv7+7iqF2/bqY7GVZbmVytqrjsbSg26W3Zu2GVa9eJob90lm2+OCMBxJPW"
                    "1zO4JuIdEfTrH6X+nDAeMn4VTHy9yIvKuSptUw0SLPkDIkMybkpFghtt+9TjkRuswinu5yUTJuNYtRpcZitstQ49Knq5z6z6l+oHIeVzSTuqWrbq2Hgnkc81"
                    "MLTJhx8DGOJbD1xA254o/sFuzZeOpz3Od/Rp6JemtP0l9MeK8HrNiNnG49s+asxMa0389eJt5e29zQPcLrcr4YXuALasNeIBrI2MbtkNbLayAiAiAiAiAiAi"
                    "AiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAirResN0zSMN5PruoXGoCv6U5WTFqsvKJXNpYpuRqG"
                    "rKKiZKmNzCUbWa4rBiSYkNmpQ3GssPy63srZ+XkEKMzZX9Fvqb/EsNkPTXJzt/V4ASZTAGSRxlnxFuyDcqsa4dOsbfm91p6y4xZKONkYiqkio7/aE+kf8Jz2"
                    "F9YMRVf+h5E6Hj/KfaaBFXzlWuRh77g0AA5WhDLSme4gfqcfV/mktOUGuQ0DFzXTqp/aUS2VJbdM1EUeSg0vxJR9pEpRRpTbLy2yLTzaHGVf7bqtz0jl6HNk"
                    "Hfv8Ab0exA7ns5od/XYI3pVwUrslC5XuRAl8L9vYO5mhPUyeEdQ033oXSRAkbjc4SDTmAjTWzhvR3n2JDftyYj70aSzsle1IjOraea7tkaux5C0kpOiWlJGR"
                    "GRpM/a3sgtdtpaCDry061rZ7Eg7+fPbsCt1U7EUkcUkT+uKZkc0Mh7dcUrGvjfrRAL43NOv8JJaTsFdMsGUuJPei3svOu78aP5P4Mz3ovz4MyMtcvGx2Oj50"
                    "dA7PYfPf8+e37bGSVZeghw38eQd9xvvodxoAn9u5+d7G8BcgsXdejgrMZrTLb8l57iK/luJIqa9kuvypOCyH3HEJ/bb95x+RjTRkt8ruZOx2Ocx+8xOBVYvl"
                    "K0mLtuz1UOMTw0Zms0ACWL7I232ANJ9yEBjLRcQ32GMnPtMr2zPjHPOOi7XdybHxOdbrRN/jteMEutUomMjZlWN6ifepMAZkOgCN9SOHISsYKeUtz9usoU2t"
                    "lSYFgwuLOiOqYfYc7iNC0H8pMzNLja0mh5h9BrZksLbfYWtp1Cx7sUkc7GTQvEkMrWvY9utOBHz4Id3LXtcQ5ruprtFpaNUt6djpO2kbB7eP377BGjsEhzHg"
                    "td9zSFwbqTIt6Iy38mZGR/Pgy34/B734ItefA5nsQf8AUD9j/lvQ1+fG/Hnts+T/AEGv8id6+d6P5I8eFxjqfk+5JERn8nr/AOP4LXgiVvR+CM9jl+NAdz27"
                    "a7/k6Otdgf3/AAF3oz48jQ3+e/52P8v6nz37Li30bPR/gj2RkXnuMz0f+T2ZGWi2Rn50RmP0Hx+x7nXbxrR2dnufkd/wu/E/58HY35B15OvOtdjs/jwumWdY"
                    "tBLfjGk0kRmtkiIlEejM1tl+U+D/ANsy7i+Ukrf2/UE77Dv+du7d+w2CN+Nnv4/ACySjeaemKckHYDJCdg/gPJ8Ent1b18HR0TIB6fnqU8q9CmVLrozDud8H"
                    "5HblOzzi1+YiOtUp+LHhO5XhU97uZoczisRIhPLfQdPlNfDTRX7TLjeP5JicbfXn6cuNesdF+SqGLBc6qwdNDPti/uMgyPZjx+dhjAdYrk9UcF1hNyiXtcDY"
                    "rxuqSSl9CfqE5F6R3I8babPm+DW5uq7g3yn38ZJI4e7kcFNKSyGU/wA9jHu1UukFzf0tlz7D7tHTl1KcO9VvGNZy3wlljOUYpPlTaucy6w9W3+M5BVvGxaY3"
                    "lVDMS3PpLmEv25DTchtUK5p5dXk2OzbjF7ukurGn/mXC+ScAz93jXKsZPi8rReQ6OQdUFmEkiK3SsNHtW6c4BMViFzmEhzHdErJI2W0cR5hx3nOCpci4xkoM"
                    "li70TZGSRnpmgeWgvq3aztTU7kJPTPWnYyWN3lpaWuOdxiyyZARARfJPnwaqDNs7ObEra2uiSJ9hYT5LMODAgxGVyJcybLkLbjxYkVhtx+RIfcbZZZQtxxaU"
                    "JUovrBBPanhrVoZbFmxLHBXrwRvmnnnmeI4oYYow6SWWWRzWRxsa573uDWguIC+NixXqV57dueGrVrQyWLNmxKyGvXghY6SaeeaRzY4oYo2ukkkkc1jGNc5z"
                    "g0EqIrqe60rDLjlYNxHOmU2IKblRL7KTYch3eVJcJtCY1OSzRMosb9pTxPvutRr+9eW2ntpqSE+zlsyvSX0Jhxvs8h5pWit5MGGbHYbrbLUxx05/u5AAGO5k"
                    "Or2yyFj5KVEBwebV6QDF1SfU39aFjNm5wP0cyMtTCH9VSz/NoA6G5lx3hfU4zKHNlpYr/tBLmNMuZJ3SzHitQDbWU0mo8YW4Tc+4ZSSPtXHrXG+0zT2J7HZj"
                    "ZdpNoIi+yApBeO0pKUpT9MJB5DLNZ1VqLzsbbLZY4631HqZA4g9Tj82QfyYiXOEgrsr0nFjJbLGBm/7uDp1vQAa6ZoIAbrXRF2A7dfZrmLIRGfk1H+P+XgyT"
                    "rwRF3eDIv7dEZa2Rn8DGyPGh877dxsH/ABHXfv5J79vHletvYA3rv31seBrwd62Ox3s6I0Pz/vkvwWtHr+da87/B/JbMi2RfnZHr8877/IB/rvsR8jwdA9t/"
                    "sRvl47dj3A/y38AEkHsCe42e+t6X+n8a8aP4Ii868fBd297L4IiLf+fIDz87H5J8gHyenWu/Yk71r+i470DoHQ7j50fz27HuP2J8fkL/AGOy5KkNx0Ea1uKM"
                    "kkkiM/JKM9f/ALS7t7Sfx/kiH5LI2GJ8jj0hg2ST2HwP37/Hf/La5wsMkjY2N2551pvYnR2T870AO+h+PwtX+pnlNmMiZw1ik1DzMZ9tfJ1nFcS4mbbxHWJc"
                    "TA2H21KQuJQy22bDLeza15IxAxl5UN3GMmg2m0fSziTrDoOcZiFzTIxw4nUlboQ1JmyQy8iljc0OE1+IyV8KHDoGNlnyzGzsyuKnp/nMcp/AKcnFcfK12SuN"
                    "jPKLURDnVYw+OWLj0L43uDZInMisZprgZW3Y4cW5tV+PvNu6TmWy/wCGyMu0/Hjyfg/OyLW9kR/OvkvI3psEnWz31r8nvrWz+TrZ8eT+Bq7QDR89xrvofIGx"
                    "27bDj8a7A9iVylHAKwnJNZd0eJ2vySPyTizNXssmaiMtPLSanEGRGtlp1vaTUlRde5O6CIgaEkvVGz8tAA63Dv2DQdMd4D3Nc7YGjyjjD3dxtrdkt2fO9AAf"
                    "jtvRG+zvktXdbq1jUVTaXlkayg1UN6fLUlO3XG2SImozf4TJmvqYhRkmoiclyGW//wCoQ8KGF9meGpD0iaeRsMXV4a5w255PyyGMPlk2SWxsd8hd6KMySMYB"
                    "rqcG/ADd+XknYDQNucf8LeokaBAzH6RnA0jqC6p5/LuWxmJmOcMvxeRrltMl5lhzPbeVKZ4zqWIaG0Pft1dYVFxk0N1qyJMFfH9VUWMGbXXa0lpD6xfURvC/"
                    "T2jwHETmO/y+J1KVweHWIePUpGnJSSuB/wC0ycroab3OZ0zxT33MIe0Fs/8A6EfSYcy9S5/UHJ1vdwHpvBCcd70W47fKr0UzMc4OJDHPxNMz5SYNY7pvTYuZ"
                    "sjQ1zZLcAqiVzyAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiwR1L8F0nUfwp"
                    "nHEly5DhvZDXtyMbu5sORPZxvMaeQ1a4rfuQ4c6qmzIdfdRIh3FXFta1V7ROWlBImNQrWUSsz9PuaZD0+5jguW473Hy4m4ySxWZL7P67Hygw5Cg6Qska0W6k"
                    "ksLZHRyCGV0c7WF8TNa/9UvT7FeqXAOT8EzBEVbP4yerBc9sSy4zItHu43K12kj/AHjHXmQW4wCA/wBoxP3HI8GkvlmN2+LZFe4xfwH6m+xu1tKG7qpaVJkV"
                    "t3STX6y3rn+5CCW7X2cSVBeU33MuPMOey44hRLO8vj2co8iweLzuLlbZx2Xo1b1OZh2JK1qNssDtbPS4se0vY7ToztrgCCF/N9ybj2V4lyPOcXzcArZfj+Tu"
                    "4jIQ7OmWqE74Xlm9dUchZ7kLiA58UrH604Bar8uY79NLj5Cyj/asFIgWHceiTNYjn9C+SSLaUy4MdxhxRaQl2C0au56Yajy2lKC0xEgFrS9oHY6J+8dho6e7"
                    "ff4doaaBrI+H5Bz4psXK4B9YGzV00DdaWQCeIu6u7obErZGgDqLLMv8A7uBoGus9nt+Na0ej2XxryXgiP7i8GX/E9H9pEO/4HjZP7d++9E72P3Hg9zr8nZ9W"
                    "Ukb8dgfj4Pb4+P21/wCY6bYR+8l9prbMlEtDjbimnWnEGRtuNPNm26060okrbdbUTiFoJxpSXEJ7eTtOGnDbXNIII23Tu/cd9kk9wO2zonzvI6VgxujcCDrR"
                    "05oc0jRBa5jgWva4HpcxwLXNJY9pa4g7y4Pmh834S7JsX2l8tcf17LWTsJbSh/NcYaW2zEzKNHJ5+RIsWVOIjZJ7JJjtXTprajQYORYzVxsShBwN4U+/8IyE"
                    "rjSeS7VO0Wlzqh+1rREQ0ur7290LN7c6tZll0/zPjTcFebaoREYLJyPNXRLmY29p0kuLfIWR9MXSDLjfcc6R9Ie26WzNj8hMOvPp0n4T5MjNX87+DLfyXjxv"
                    "ZePHwMlHc7IAGuwPjWj+37dho/JA/GKRnbvk+Qdgg/J0d99/k9j+e64h0tb/ACWz+CMiLRnoiIvHjz4P+T3oy2AP+f7b7Ene9n+uwNk7AHbwV34ydjz21513"
                    "3/U/vvX57d+6414vJl2lsy8/n8+db8lvWtF86M97PRuw+e/f5120fBP5JB/r30NaXdYew7ne9geNfO/3/ofyO2u64l8vJmXzs/OzMzLWzURERl4M9kW/B/BH"
                    "sASCOx12/cfII8d+3ft/+KfC9GMnXka7Hx4P432PYa79u3ba6va1yXiN+OSUPpLykjIku6IzNOu3woz+N6Sei+D0Z/rXbPjRJ/8Ah3O9DsCe+u3bXcke7QuO"
                    "i1HM/qiIADiCXMJ1rzvqb48dx20T2Wc+lPrG5s6MeToPInEl89D7iiwcvwqzVIlYfnNC1JRJfpclpfeYbeW0o3lVNvEdh3lC9LmvU9nCTY2TM7Vvqx6RcR9X"
                    "uPvwvJYBFbhc6TDZ6rHH/FMLO7QLq0hLffqyaLbFCdzq0/UH9LJ44p495elPqryj0pzn8Z43ZZPSt9Lctg7ckgxWYhaNNE7WB5rWox0+xkIGmzD0Njd70Bkg"
                    "kvFdDXX/AMJddOAN3uBWDWO8j0lew9yNxDazkv5NhktT64a5cSQqNARlOITZCEu0uV1sRlp6NLhRb6uxvJDnY5X07erno3y30ez78Tn4Baxll73YTkdOOT+G"
                    "Zms0B3VGX7fVuRBwbbx9gieCQOcwz1nQ2ZrbPS71Z4t6r4RuUwM5r34GRjL4G25gyWJsPBHRK1pLLFZ7mk1r1cvrzs1sxyiSGPekalWz1wGUZTj+FUFnlGU2"
                    "kamoqeMqVPnyvcUltBGSW2mWGEOyps2U6puNBr4TEifYTHWIUGNIlvssr9DF4rIZvIVcXiqkt2/dlENatCB1Pee5c5zi2OKKNodJNPM+OCCJr5ppI4mPe3xu"
                    "QchwnFMNkeQ8jydPDYTE1pLeQyV6UQ1q0EY2XOcdufI86jhgia+exM5kMEck0jGOha6luq++5nnOYrirUul45iyS+lrTJabfLZRIb9mwyFDLjzZMxZJOKo6G"
                    "MbjEV9CbmwfsrQqpvF50elXo1j+FV483m3Q3eSyxdT5SWmnhY+p/XBRc4Dqlki6W3b7wHPaTUqtgrfqZMlTR9S/1a5j1bs2+H8NNzCenEUj4nN6XwZbmL26D"
                    "LeWYdvrYhrh7lDCN9uR7XC3mDLZdVpYnX+gxooZosLNKXZxn7keOZ97cNStK71mXcl6YRmZmvuU0yoiNk3HEpkFs/I5U2OqtVJZXA6ZJP5XTNHbTQdFkGgAA"
                    "QHvB1IGt3GYi1aQYRJMGul0CIwdtj3s7Oth8rdnZH2MPdpcQJF3AyMi8a+fOtGfgv7j8F8ntRbPW/wAGY8QEE/8AhvY7k+B3OuxAOh438Lu9tH57jzouO/wR"
                    "vX7dz4O/wh70r4Lz8GfcW9n51518n+f5Iv5AfA79x5A6fjxvtvwPI/z+F+gdvwDokg6I8dvknQ8k9u+yQe6/zfjXaXkv8edaL5LyXxvuIvB7My2RmP3R+XHs"
                    "f37Hv8f8tfI1o9wC77J3v8EbAJ8b7A+BvsSCdfI1v+Pc0W9EX5I/47v+5kXj438aLf8AOv3p762TvQP768fv37//AAX4NnejvQ0ew2dgflpPY/kd99/hdF5Z"
                    "5NTw9iKJFa+x/qPmEd9jE2NpckY7WEpTE/NpEduQ29HXXvGpjFXH23YdnlDKFrjWlVjWS15+5xDizubZl0NmJ54xhJYn5qTu2PJ2iRJXwEUjo3xyi1H0vy7I"
                    "3Nnq4mRzWyVbmUxVhenPfj4pixl3FgzmQEsXHq7ifcgdG4xzZx7IpoZom4+Tvi5Hs/T2sxXaNWYMXk6b4y3DMzUalKUpfcaluOOOuOLWozceeedUt1x11Z+4"
                    "t1xxx11w1LcUpSjUJTgg6ADWNAADWsa1jQ0fa1jGNawNaNMaxrGNY0BrWgABad257nPkk9x73uc5zyS575HGRznHZJcSdOJdt29nYIC+ZxRpL7e5ZmZaSSVK"
                    "NW1a7UkRGpS9q7SSW1GZkRGXgi5t2e50ANkuJAA1vZcQdAa7k+ANk7B2P1x7k/udd+3c6+3999hsbGyO+wDk+or/ANrgNMKJP1CzN+YojNaTkOJ0aCV8m2w2"
                    "lDKTJJJUaFu6JTi941asfqJ3SbIjb9sII0QxpOna+HPcS862fuDPAaR2Y2dI3vqd8nuPJ2QNnu1o7aPxsgeStduoLLFLdqMEgubUg2b/ACImyfW4Xe2tVFWO"
                    "NtmXuEbDjt9Jj+26Sku4/MY/3G1pTknGa8NeG/nLZbHXgilggllc1sbGQ6dftl7thjYyz9MJCQW9F5jwGkOPv4yjPOGsggknt35oaVGCKMvmllme2FrK7QOp"
                    "81iZzK0IaHF7nTRA9Wgbdvp59N7/AEwdLmB4Pd1r9Xn2Rtq5C5Ohy3q2TOgZtlEOB7+PzZNM/JppUnDKKDRYQ9OqZMqvs3cbXasy5q57kyRTT65eoz/VD1Iz"
                    "3Jo5HuxTZf4Xx9j+oe3hKD5GVHhjgDGbkj58hIwtBZLbewjbV/Q19PvpdF6Q+lfGuIOjhblhActySWAfZY5DkwyfIkOJc6RlZwjoQOc46rVIWsDI2sY3dwai"
                    "W6UBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBFWo9YLpyi4JyfSc90UZDFB"
                    "zE8VHkTfdNWiNyZQ1D0rcdDzrsGKWV4dUO27FZWsRCdsMRzfIp/1E23kPOWT/Rn6ofxLC3fTXKWXOvYL3Mlx8yOj/vcNYl3bqA6bI92OuymVpc55/T34oY+i"
                    "KqGior/aDejZxOfxnrJhaYGO5G6rguXiBjtQ52rXMeIytjpaWNZksfXbjJJCQRZx1JoBktlxhFyKljXVZPp5JKQzPjqaW6ntNbDiVJejykJ3s1xJTTElCN6c"
                    "WySV/apQnjDIWPa8bJaSQCD9w7hzTvXdzSW9hsAjW/it6pclx9yC9E0PlrSiRrHaIe3pcyWIn4ZNA+SB5+GPOgHALRu5rpVfLkwZrJMToTz0SUyhZOIQ+wpT"
                    "bxNup8OtG4kzZdQSUvsKS6jaVpMsga9r2hwO2nTwSNbBII2COx0NuBBII0Rvud60LUdiKKeu/wByvOyOaGRzehz4pGhzHPZs9Dw3XuM24xvDoydtdvpUxk9+"
                    "E/J/n8mZmZl5Mt/3GfjR+C0Z60OY/wC8e3kE7P7Dt+fBPcg78d9rJK0n8p2BofjffXn/AMR3/p8L8caya/4/yiozPF5CI1zRyVPstve8qDYR3WlRrCoto7Tz"
                    "SpdVaxHHYc1nvJxCXUS4bsWxiwpcfpXqUGQrS1LDS6OZoJe3Qkjc09ccsTjvpkie0PjOtEtLXBzCWu9aSvSylKxi8lGZadxgjl6Qz34XBwdFZqyPY/2LMD2t"
                    "khkA07RimbLWmnhk3eunaDLseqeUcIbW3iuUOvMzqlZtuSsOyiP2HbYzZFHaYjNPx3XkORzYbajSIr0WwgNIp7SnW/5WJtztfLibx/32iGkSn7W3KziRDaj2"
                    "+R2nNa4O25xa8Oie58sVgMj7kcXdwOSnxGQDP1NZrZILMfV7GRoydf6e9Wc8CR0cjWOBDx7sUrJa9gC1WtNbj54iM9aUfyfjz/jRqI/jZknyZeTSZEXcoh7J"
                    "6fHgnud/B0daHfsfyO+t/Hco+4GgPGtnX5H7b7+O3n5Plca6WjPe/g/8+fjXye9GWtfgi18mZhsDet9vkn/Q/HbsP+XwNHus+BoedDR8bOuw0Nd/P/mFxT+t"
                    "nst68mZaIi3o9kZEWtH+Pn8aIyIO4+SCRrx3dreiQP8AT9zvuO670WwO2vjuf20PGz8b7+D8fk8W74I/7t/nW/BfBaIz+T+PH5+PJmPzY7Hq/BB8a1o9j38e"
                    "djx27fK9GPfb+XRHjt/p22fj8jt8689WuK9qU2t37mnknpLiUmezLZGS9aM0lr40RkoiMlmZEZ/QHemkk7+4/gkEa2NfnbmnzrfYDz7uOtvrua3XVGfLDoa3"
                    "3JaT2DgRv9+/9V/fFXMXJ/T7yFj3JnFmXXWD51idgiwpbyne7VIURKbkRZUZ5L0G2qJ8ZT0G2prKNMrLetfl1lrBlwZUiMvG+W8Q47znCX+OcoxNXLYjIxFk"
                    "sE7AJWSgH2bdSdgEtW5Xc7rr2YZGyxSHbXEFwO0uH8pzfE8xR5JxfKTYzJU3/wB3NC4Bro3uaZatyu7qitVLLWMbNWnbJDKGtcB1xxvbc66AvWf4l6oMMnYv"
                    "ytDLBuonDsZVc2OHUbZzYfKtbTwY5X2V8dQnpHv1SIk1fv32PZBOKHi0KdBmryq1q0WlhVVQesX0p8q9Ps9TPH7MGb4lm7jocfk7E0dezhnOkiaypnWAdw10"
                    "8VetepxyDIz9MMdSG5PXqTWRcN+qvhuQ4Zlc7zGK3hMvxuibGUpUaVq/DljHG94dgXxs6JLM8cM1qTG2Zo5aFSG3enndi6N3IwY45s5v5I55yBhN8lFVXwFt"
                    "lUYNTWMqZj1HMOI5EtJf18qDTSL2apMqZCVkltU1kg6t1yLBqMfj2drWzdzenfp7x309xXvw9FrI2ogchmbELP1dhpe2aCtAxrphVrscyKRtOtNI2Syxs9ie"
                    "3JWqy1qt/Xz6h+ZeuucMc0s2J4hj7T3cd4pTszNps6QYv4tlnFlc5TLSsc4Ms2oGR4uvPJToV65sX5r3UKXH41Qj3T1InLR2uSe09JSZaUxEQry0z57VKJKX"
                    "pBl3vmSPbZZyO/kZbrugbjrg7ZFsdyDsPlcOz3/LRssjB6WAu6nv0hXrCuC5zvdmLdOmO9b0W9MYd3aO5a4n7n6Bd26WM5//AAe9/BEReTM/J6Mz/Pzsj15M"
                    "yMz0PO/zH57nsAPHYf17DW/HYDsfsAf6Akn8geNHt51vQ0fJ1ryCVrXn42Rf2l8a7TIzMiP4PRmRbPe/PyYee3nX5+fPb4/yPb47dgP09h27A9vPzrvvz3A1"
                    "sdiT3/GvzPXkjLxr+DP7v+5/OtaLetERlrXxyAPYjud9/Hj/AMgd7Ot72CF+A+Nb1v8AOgN+e3fsPt1vY8jtvv8A4atb8GX+dGZl4IyL/Hyei18Fv5LxyA3r"
                    "uD/mNHvok9vPjZ352PC/HaAOydjsQHHYB3v99duob7dxoH5/OytqXDMbts9yv3k47Rm0hEVhSW52QXcr3E1OM1BupdbOztn2nEtuONuMQIkefcT+ysqLF9n8"
                    "rVL2bytLjmH6HZPIh7jNIC6vjaEPQbuWuhjmP/SUo3sc9jHNlsTS16Vfdu5VY/1sdUrRVLudyxfFhsOxkltzXhklqzIHili6kj4pmNu5GWJ8MHXHKyKKK1dl"
                    "idUpW3RRd5rl99nuT2+V5HIQ9a2z6V/TsE63XVkGMgo1dSU8Z1TiotRUQmm4cBpS1PLQ2qZMdk2UqbLkSuwOFx/HMVUw2NY9tSpG5pklLJLNyxI4y2r12Rga"
                    "2W5csPdNYdprAXCGBkdWGCCPV2bzFrkGUnytoRROmDGQwQDpr06kEbIKlSuxxcWw1a7IomdTi9/SZJnyyPklk6qfhJb8/j878n2n+d7LR/g9H/BmRl6/keNA"
                    "9gBvWj4Gu2wfHcDtoAdiD5ZOj09u/wDk4gn/AIRobIB871sDXyufxmsVMsDmLIij1xpcSWjL3Jij2yW/GyjkSpC9GRodKMai7V7HnZOyI4PZB3JPsOP/AAwj"
                    "s/Y/MjiGNb3BZ7gHcbP0ja5z9t0en4Oxt3ct8j4Pc+dHRJ/mI7hc21fjtTa5BbuJarqWBJspZqcS0p5EZszbhMOLI0qm2Mg2a6uRozesZcaOW1uIIvFihmtT"
                    "16tYOdPZmjgj0C4Nc8kOkc1p/wCygYHzznYDIIpJN6aSfSrQyTywxR6DpXhrTr7Rvv1uLQS1rWhxeR2Y1pJ7NJXJelh0+2HVN1awc3zKG3Pw/iqWzy/niZVU"
                    "drVy7pNo8XG2HFKkSW26v6zJoSbiqTIi20adiPG+VY0qHGKWxOh6v+rb1Fh9PPTJvEsPO+HL8vj/AILUMcvtTwYavHG7M3z7Za90szXxVHuHSTNknTlxc0iS"
                    "dX0U+lEfPfVc8mv1jLxf02hhyBD2H9Nc5DZD4cLTmjLXxSe0G3MxZie8EWK9KQMe2R3TcnFRCutQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQ"
                    "EQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEWCupXhCo6iuEs+4ktnERXclqDcx61W7LZTR5fUvN22JXTqoK25T0Gvv4UB23rkqNi6p"
                    "jsaSc1Ir7GXGezL0+5nkPT7mOC5bjSTPibgfNDphFuhOx1bI0z1hzWm1SmnhjlI6oJXR2Iy2WJjm4F6n+n+I9UeB8l4Lm2B1LP499eOU76qV+J7LOMyERAcR"
                    "LQyENa2wgbJi6ewcVSdyegsscu7nHbuCdVeY/b21DeVS5TUt6mv6GzlU17SS5MYzjuWNHbwZ1RZEyo0NWMGU0Rkps0pvJwGao8gw+MzmMsMtUMrSq5CpOz7W"
                    "y17ddk8T+nW2dUcrXFh05hcGvHU1wP8ANvyXj2U4nyHNcXzdaStl+P5S7h8jDIwxltvH2JK8j29huCYRietLvpmrSxTMcWPYTqhzNji2X4+Txm9NSParLcu1"
                    "JJbloR21k8yLZq+pjNnXPLUSUIXDrGkdzslRll9GXraYiQPL49kbDSQ57QNAjTiXAdzpzjsaWUcJyQcJcRKSHxB9un5PVAXA264P8gEcrhaia09Tmz3JC7pi"
                    "C1qmNbUrx48fJl58l4SZ+DPyZfnyWy15I/RIIB7b0P3JHbuAR27fg9/Ou+itqV3jQ127eACPgDx2I8Hv58/BGuvymy7T150XaXkj8mR7T4ItGf8AB+SMvnfk"
                    "+PyO/wA7J8bAOj+d6Pj4JG/C9eB5BBG/IPyOxOh3Hb4/z320u88PcnscYZFNj5E25O43zFlmtzmuQ09IVCQg1Jg5dXsRWZEs7THVOOqfjQWlSrWndmQ2GH7W"
                    "Lj71Z5eUoSWWw2KhDMhUcZKzi4AShxBfVkc5wYI5w0Br3FrYpQ2R59oytk6/JuODlWMZFXEcebxznzYad7o4/ce/vLippZXxQtrXy1pjlneIal1kMxlhqy5E"
                    "WNkstx93HLNUX325kCWwmwprSO4w7Et6p8kuRZsV9lbjLra0La71tKNtXc3IZNcSTFfd50bcd2u2UfbI09E8bgWvhmaD1McHNaR9wcNOb2O2kB7HNGiq7+sP"
                    "6muY+KV0FiJ7XNkr2I3ObJFI1wBje0scC1wLhp0bj7kcjW9IeM+4y8GRGfkiJXk978fxr58mW9GXn47XSDobG+3l2vx286/YbB2NnwvTj6SBryda89/B8dj3"
                    "Px+xHdcY/s1EWyMyM9FstJLxrx2kRF4I9mWiI9n43t5/bY7kDRJIJ8Ak9/Hbv31regu9D48Ej86Pc9977ne+29fgDzrXFvF5PWz0Rlvwej/9WiLfbrZn3Fo/"
                    "Hgt6HH+vb9+/b9z2/wD+bPc6XoRn/Lejr414Hnx8dgT38+drhpSSJtalGhKEoUtS1mlKEoIjNalqUaUpSlJGalHpPaXcakpSZDmNg6Ghs+D3I7jXwT2/Pdvc"
                    "fK78Dvua0BznFzWta0dTnOJ01rWgOJeT9rWgF29fbvQWIL3Jf3CSxSYtVzcouLCfEqayJV18mzlWlvPlNQK+nx6shocn5HdW02RHgU8OvaUVhYSocSvOyfkI"
                    "jHwtWa9WvPPYmhrw14pLE1ixKyGtBFCwySzTSyEMZDFG10ksr3dDI2l7z0NcVtLjXFb1uxXFouquncPbqh4bOWkdXuW3dLxRghYHSWfcDp4oY5nTMrBhkUvv"
                    "N3TTN6MehjB3oNs5WdRTXPnE/Kedcj4xYtHa47yfU47yPGx3HcUuWm3m04pxXX282gqIjb07GclyV3M87KGqJnK6aFGTinIY/Uz1JykuSr/qOOXOJchwuLxN"
                    "xs7YJ8DNkuPOms2qkk72Nv5l0MNuz7UVSxWqNxmPlj/V42a1ayafnT6uTPHeMSiPAYm5R/ViRkDxmcnFj89JctzuirVpmtikb+jryttX7cXt2ZaOYdSloRUp"
                    "Lui7qLoOrfh6TlSK6txrlfCplRjvMmO1HvKqWMgnwZMujyulbcZUuuw7kaPX29rjdUqVLcx+4pcxw1MizTiJ3VrqrnnHshwLkbcTbnmvYW6y1b41bslrLEtG"
                    "GWOOzUm/vXOsZDCusVa160Y2MuwWsdkOpkl6SpV13zHhGPrVouQ8drTR4jIvbFbD4y843KdMjnUrE1erXoRMtezZsUIIRG9scU7/ANBWqHHPubJyY0iK4bUl"
                    "Btr1pJmZGSy0SvtUX2q0R7/kj8KIj8F4kcsczeuJwc350Dsd9acPI7/Pg+QSO61PPXkgcWytIcD2Oh0uHnQcSQ7to6GiN99bC+fwXwejPz//AL8b393+C38b"
                    "/wAefp3+Rsd/HgdvjXb9z+dLrnx4IAPb51+Bvv33/i18dvK/JR+DMzPtM/nxsy/gy0ZH4LRkRmXyfki2OYA3oa3rwQe3zve+3nY7b+OxXHuABp3b9hvx56iC"
                    "Bs9+wJB7HsF/JmXnyRJPRa+PP5+dGfkvJ/BkX5PRHyA8b2Ts9/yO/f5+D2H5/bZXEdyNg/17nfff4OiD32O43sjwVyNTWu2stLRPtRYjDb0mxsH3Go0Sur4r"
                    "an5cuVKfNEeO1HZQ489IkOJYisNvTX1JjR31o6l22ylC6Qxvlme5kVatG18s9mzK4RwwxQx9Usr5ZHNjZFE0yTSOZBG0yyxtd6WKxlnK3YqcAADiXzTSOYyG"
                    "vBCwyTzzTSOZFDFDE18ss0rvarwtknnIhhkcI/8Anrllrk/JYzGPm7H4+xUpMHDYa2X47lib2k2GYz4spiNMZn35NNFAg2LTcijoGIUF2NGuZWTP2kjPTvhr"
                    "+J4qaXJGOTkuY9qxm5myRzNqNj+6rg600EkkD62OLnusz1nujv5GSxOyeWjDioqmH815FWytuPE4gBnHcMXR03tZI1+Utu6WXMzMZoorAbbdDGKVaeJv6KhF"
                    "XjfGb78hZt4BWe1KP7vGjVv5/nfgvgzMiSZFozIjPzrewh/Qd+rXc9u3ju741s9/B7aWGgaAH8oBGiO2uxGu/wAdge3fRABJ7r+EpddUlppCnXXnEstILypb"
                    "jiiQ2hJfG1GoiIzPW/yXlQEta0vc4NYxpe53fTWtGySR8kbP7fjwF+u2dBocSdaA/wAX8ugSTrv238+B37k5er4DVbCYhoNKjaSS33EmaSfkr0bzxEou8m1r"
                    "+xlKi7ksoabMjNOzxOed9iZ8zu3UdMZ2Ptxjfts7HpJDe7taDnlzt9+/ca3oaRpzj8u7uJcfwTvsT89x9oBO9LVnqVzZDKanj2K4lH1jTOT5I4t5pDDcVmSp"
                    "nH6+S53H7HfMZl3c1qSlKWmouOz2zNDxLLLOJ0gDZzUrmMjr9dSq6TTQH+3137Ae/TSxkJZWZI0/a52QieQGlpybE1XGF1se4+V8jqteOMdczw5oNjojAL3P"
                    "e1zYI2jqMhkljAcR3tZ+lx0yI6bulbElXdK3Vck8rtR+SM+73LNyfFTcMe5hmMzmrZtl+ol4ziLtaze49EZaqazN5uYyIBzFWUm0sKhPqM9Snepvqfm8nWsP"
                    "lwOIf/AuPtd0Bn6HHkxT24gwkFmRvCzdY5x6zBLAx4b7YYy/n6ZPSpvpN6TYHD24Y4+RZlp5HymVjQHuy+UjjkFSRw/7Q4mi2niWv/x/oi8dnqRwaJUg0BEB"
                    "EBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBFW19Yfp6XiXKeN8+0kFace5Y"
                    "aax3KpRz45sQuSMbpu2C0iDKkFM9zLcFpylwo1LGdr4rmAZZbXBxZ91HcsrH/ou9SRcxOU9NchMDaxDpszg+sO3Ji7U0bbtVrgOkfob836ge4et7MgWs/u62"
                    "hUl/tC/SA4/NYT1iw9XVPNMrca5YIYyRFlqsUrsJlZywfa27Qjfi5ppD0e9Qx0XU2SdrZIQL+li3dXPp5qlNxbKK7FddJtDq46lkSmJjLavCpEGSlqXGJWyJ"
                    "9hsz8fM+IZXxvbIxu3xkuDd/zflp7+HAlu99ge+lWnVuS0Lda7X0ZqszJmMLnMbI1u/cgeWkOEViMvgmOyTHI4b0djQa5r5VfIlQZzJsTYL78OW0XwiTGdcY"
                    "eJJ+TNv3GjNpeiS42aHUmaFpNWSNLXtaWklrmhzST8HpI2PA7EeN62R30FIDH2YbLIZ4Hl8E8cc0LyNEwysEsewR2cGu+9vljw5jgHNIHT30mRH3a8a8b2RE"
                    "W9EZ/O9FstEfgvBpLyPwjvrx4HjuN/IH7+dfnsVkEJBI8/kn8779v6eTvRI7fgLrs1v7VfaaTPZ/JpPez7SIy/77UaVeS1oy7SIfh3sDzo99aP52da3v4G/+"
                    "920QR7Vd+tAHetaPY7H77777eCB/RbIcD8hRsgrWuE8wmpYWyqRI4nvZBJL6Gco1uyMHlupQn/wcvudlY+qU44pRfVY5GV70bCqxHh3YH0rH8Srt6o5OlmRh"
                    "bv7mjsy00DenNADZtaHiZxaHW3uwH1E4yWCXmGKh65B0N5PVj7maEBrWZ2JjnEmZh6IcoIWtAHt5WRrhLm7S7daQpECVJiSmlsSIyzafbWWzStBERnstpUSu"
                    "5KkLSZocbWhaFqQ4Rn6jHtkY17SCHjqYd9nA/GgPI1/Ie4I0W7GhrqtIJGscHdTXbIIHYgkjQ332C3XSQOl32kAt0OtSC8n58GR/BbLej0WvGjPfg/4/z4V+"
                    "676122N/kbJ1/wAXx/roD9160J2Pz/l5BH7b7jv+wP7Lp99ktdSpUh5RyZnb4gRlIN8iNJmg5Cj7kQ2j0k1LcJb5oV7jMWUkj1+tae2x5P8AlrtsbHnt3BPb"
                    "uCO5BGRYnDXMmQ6MezWJ725mkRbBAc2Juw+d/kAR9MQcCySaIkA4NyPI7O9NbUlz2YRq+2ujdyYpEXYZKfMzNc10lI7iU+pTbbncqOzGSo0Hz0QN6AA8AHZJ"
                    "J7g7HYa76OvBI/baOHw1HFBr4GmW10/dcn6XTEnf/ZAbbWYQ4Nc2EB7m6bNLMWhwmN9JzpKRNkvdX3IVXHerKOfcUHANNYx3HVzcorjfqcw5dkQZtZ9EuBic"
                    "k5uF8b2MSbKfRnsXOMgTFqLnj3D7eVGP1w5s6ewOA4qZzRJDVtcstQSBgFaYixj+OMkhnbMH3oxDks3BM2OJ+Hnx1KQWqecvwx+1yjNx8W42DWstHI87FMyv"
                    "EGRvkxeEcwNfkXmZsrGy3y+SLHzwxttVrNf9XTuVLePnhl2q9U1S1dLLayMz7+XsAUtJ+EpbOiz4i8dqjLtWpOiIt7VstGZmPE9FGhvOmNAADeO5kA/PV+vw"
                    "TtnZG99J7n4+VqLiTg+zZ6j/AO8rvbr/AIhTyrANEkaDJH78AHt47GE3pZ6h73pe5ux7lapgyrenKJIxTkPF4aoyZmW8b3M6tk5DS1y5rzEJnIoL9XWZNh0y"
                    "W8xDjZbQUrVq9/T8y8iTZBeofDa3N+M2cRJI2vdhlZksRdf1mKjmK0FiOrNO2Ih81OaOxPRyELfvkoW7Jg6LbK08O2cXNUdDbxOTdvFZeFta2/TGyVZBI2at"
                    "egndUtyV317EbJJHw15i6MOMla4Y4oDbYp8zx3K6WmynHbOBl2F5XSVGS4vkdYTqoN3j15XR7OmuYSZSGnlR58GSxLRHlsxbCEbjkSc3GnR34rUFjjrdaSWt"
                    "YhnxmTpWJ6l6jOWe7Vu1pnQWqkphc9ofDKx0RlifNWnDWzQOlgkjlfprkNK3gcvfw2WibNJSsSQuljZKxk8bCRXtwtsMjf7c8fTMI5Y4bNcudBYEdiGSNn7O"
                    "QWX0m/XOodbIyNTXcfehR+Cb2ou9KvH9r/gzPZOGWiH1bYkjPt2mOY7vp+h0lvnqGjojv5j7j5aDtY4+tG/clV3W3eywnZb2307cOoEedO877OI3vhnEmg+x"
                    "ZKbWSvKVEZKL/OvB71rf2l+T0kiPffaQR1N05pHYggg7+N9xre/nxod+y6DwWn7mkdgSDsFpG9DRA13J/p2Gl+SW1uOE2j7nFn2oIiUfcZkZdp9u/BkezP5L"
                    "Znrzoci4NYXO7MaNnehobHjfbfbt+ddlxY3rcwNG+r7R+N9t9xve9EH5Pha+dSHKRUlbJ4ZxmWkpMpLL/JtjHNJr/wBz2ZcPAkPLZV2mfZFs8sdgvJUhn9tx"
                    "GRJPvzelLYvpfxI5G5HzfKwn2IHSRcVqzdTWuLRJBPyAsDx1AdU1XDNnYQ6QWszHHoYG8nLcv/AcW/i1GRgyeTijfyCePpe+pUc6OaHDNk6SIbUxjjsZcQvd"
                    "LFF7OIkdBJ/GqUmkKzMzPZ+PyejNW/nR7LZ73538nvZGZ7LfTRr8a0O3YDX5Oj+x7k9/2Gt6rb2/m79wdkk9/A320D5Gz5/Ynv8AOaS/9Wj148metpPZFovO"
                    "96MvOv8AunQ/dka7eO520De++j37dvk6H/i793+e/j/6ngd+2zs9hsjZ357dh9X7r7ts+kyRH3GgkojT3SFINMh9PgkqJllfsI33JNch3Rk7HIk+Tl7XSxlW"
                    "PXVJqScjuRG07jZ52Otw9x2tHUbP8D+/3haOtzz26dADsNntt3jsQNka+0k+dtC7fcWVZQ1VpkF3Jci01FXy7e2kMsnIeYrq9hyVKVHY7kFIlKaaNqLFJSVS"
                    "5S2IyD9xxBDxImTzyw1qsbZLVmWOvXZI7pY+eZzWR+48NcWMDndUkhB9qMPk10sK9GvWlszw1oGl0s8jIYgXdG3SPDWhz3ODGdz3e9wa0Auc4D7li308eA5v"
                    "WN1n0EnKa2Y/hVZZTuVeTVQVwUV7GL4q5WFU4VIXZMS2p9ZkFq/h3Hs+raZRbTsGm5JOrZcKXUPWcHFfqa5+z0v9IZ8XjLDYs1n2njWIc5z/ANR02Ypjmco0"
                    "Mc1zJmQPsziUgRRXrMHUD1sifNP6SPSuP1I9V8O27W/VcV9PoqvIsqJIw+Ga1WsPfgasokAYP1+ZhkyckbHSPLKdhhjEL3Pbd2FNivEQEQEQEQEQEQEQEQEQ"
                    "EQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEWD+pHhev6g+EeQuJJ8hmA9lVIpNFayGEymK"
                    "XK6qSxc4ncSIxl3S4VdkVfWybKAlTf7lWImVq3ENS3DGW8F5Zc4Py7A8pol5lxGQisSwseWG1TduG9ULgdAWqck8HUdhhkD9EtCwb1K4Li/UvgnKODZhrf0X"
                    "IsTZoiYtDn07Zb7tDIRb7iahejr24i0tcHwjTmk7FKPLMZvMUv7vGMmqZOP5FjVxa4/f0U6RXyplNeUk6TU3NNMm1MqwqJkqqtIUutkS6ufPq5T8R12tnS4b"
                    "jMhd5fHs9juSYbGZ7E2Rax2Xo1shSnY1zPerWo2zROEcgbKwhrx1xSNbNE4mORjHhzR/NryjjWZ4ZyHNcT5HXFLOcdydnFZWuwuLGW6UxjfJA8n+8rWGdFmn"
                    "NrpsVZoJgOiVu9PubcWTGmQsritETdh7NZbqSlW/r2GDKtmuq8kZy4LCoC1F2Ja/bYaFd7ko95lj5+pjoSd9Ac9g3v7XP+8Ahw7AkO27R27yQNjLOC5Xqilw"
                    "0z9urB9uiC5vetI8G1XYD3IgsSNssA254t2HaayAFazSmzIzLRF538JIz7leNfarR+TIvn+3X50XfHjt8efgDt2HYj/nrex+StrV3dQB3vfYjfnXz2/Ov2/b"
                    "Z2T1+Ugu4/HwfaXgi2f+PJER/wAkRHr+dB+O4I0D58Dt50dgef661rudexC4/Ovydkn43oEb7ee4/Hba6xKYX3pdZW4y+y629HkMOLYkRpDLiXmZLD7Zk5Hk"
                    "R3223mH21pcjupQ80pLzZKBwa5vSR2IO2u0WvGtEaIAOgdOGjvx3BOvdrzADTgxzHMcySORgkjlje10ckUjH/bJDJG4skje0xyRvdG9rmktO49VyPV8l4jFt"
                    "Lp+ND5Ex5pmvyRpKWmkZJEShRQsihxm1JcckzCbM58aGybUGx+sZS1GrpFIk/Kq15qcs0DOp9Nx9yEk/dA466o96A6Bs9Jds6DN9T/cc/Rmf4xPxzMGKhBLJ"
                    "gMm58+Nk05zMdMX7mxlmVxe1jKwcwV57EjXWqronlz7MN1rMR5FlUxw1sVyVwWNGSpJGRznSPyRoURmmH26UZpZU4/4JaZSO82y9Pp19x2Sfknt5JBPwBo9z"
                    "30RsgDxkuGwNZgE1sstS9v7rX+7R67EEFodZPg9Uoji8tdC7pa84kfLysz2e1GtRmnZqNSjNat/KlLM1KUo/KlbNRq8756Gh433Hx3Pwe+wD30RvYHf8lbCj"
                    "PZo7dgGtA1oADpYAAAAGjQA+BoDWtLYHpP6abzqt5qo+LK2bMpMdYiPZZydl8FERcnDeNaeTDj3VpB+sS5DXkdzNn1mI4bHeaksqyrIaqbYQ3sfrb2RCwL1G"
                    "5rW4Lxq1l3Rss5CeeLG4OjKZGsyGZstlkggkdHt7KtWvXtZK9IB1ijRstiJsvrxy+/jBViZbyeSc1mNxVSa5O0sEpsTRlkdWq2v+ppvsNnuz1IJWNsVmEWI6"
                    "5t1JrdaRWyW67HMYo8ewzDKKBi+GYfQ1OL4pjFQhbNVj9BRQWauoqoDbrzz6o0CviRYzb0l56ZIJr6iY/IluvSHoRU2W5H2buQsvuZC/bnu37swYJrVq1K6x"
                    "YmkbFHFGx0s0kjzFDHHBF1CKCOKGNkTNC8t5Fb5Fl7eVtukdNZew9JlsTiJjGNhhhEtqaezKIa8UUXvWbE1mcsM1ieWZ8sr4+vVGR3dKCD8aLk/jnwf9xqVT"
                    "5qRaI9+O01HsyLRfbsyPR7b9FSP7fEDZP8DzYHcaH+/4Q+fBOwOwJ+T27r1OJkCyf+8yud6+0ObQyrwD37Et2dgE9YH7k1w1o8Gk0+DIi/t0R6M/Bmevkz0Z"
                    "K+fB6I9JEvz++wBs99E60PjtrWv8v6hZ+13gg+D2O/zokEAfHYAj9x5HaZP0qupxyFOsulPN7RBV9x++5RwXOnv6OvvXV2WQ8g8YqmzJ/wBMzW5Cpc7kTCK6"
                    "NCZV/VhckQlybO4zfF6yNGj1v4Z7c0fO8ZA8j3adTk8ULSeuNsUNPFZ0QwwueZKzY4sRlp5JC3+H/wAHsO/TVMRdmfz5XQj5JxxllroYszxekYGyWLEFdl7B"
                    "Gw016Mf6mSFhsVLVqYVoYpPftS2K9KhSs3LtiZk1CjfiupW24pt5JGZLaUaTIlHs0mZltSTPypOjSvREruL7RosCOZha9oewnuHgEbAPfZ7A+QPDhvsQQVoQ"
                    "F8Li5rnNd/hewgO7+e/YOadEa0QfwV9CLyM6gmrVrsUZ6KbHQZGkzPRG60j70kRdyjUyTnnSSjknZl8XUJWEvpvLh8wSuB3+el7vtPfQAkLD/iMhPn7frYpB"
                    "0226O+80bT9o7Ebjb1OGh9xdGHacdNiAJ11nk/PGOKMZKbTKj2ed38Z9GIsGTDrdLG8NSMvsYzzhGpqt95KKRmQw5Hub32o62JNdAvExvV4rx6XmWV/T3WyV"
                    "ePY6WN2akBka67INviw1aVjTp1voJvPjlbLSx/XIJIbNigZe3cvwcYptyET4rOWtMc3C1z0PEJ6CHZOzD7jT7NYODqjZYzXuXR7b2T14LsTYyJK3VyH3JC5D"
                    "sx196RLdmOOOzJEqU4uRIlynXTU+9JmPurlSZL61PSZLrjzjjjjhqOVcTWNjjbE2KOGNkcMMcIayGOCFrYo4YmtAYyCGNrYYomBsccbWMY1rWgLTMss000ss"
                    "8kktiaV0s0sr3Pklkeep8kpftzpJHOMj3uJL3HbnHez+Bkelfk9+T7tkW0noj3ot/wAeS+FfBkY5+O3f5I0B/wAXnsOw/wBRo/5Frexrfjt47/dodgB3bvZI"
                    "2Pj5C/hmK/NkNQ45Ep6Qs0INWjSgtKW484R6Mmmm0Lfc7SNXttq7NqMkn+PlZBG+WQ6YwdWtEEk60xvn7nuIawHt1HROgdGtLiAD9x1sH9wPkeB0nYJ7/jWx"
                    "rL0OMxBjMxY6SS1HbJpOiLZkW1m44aT7TddWtbryvPc84tWk7IxiU0j5pHyyHb5HFx2fGyAGtBG+loDWMB7hjR3OiF3gA3oaG6LQOknQ352QN+dkk6IPUO3g"
                    "g6l9VudfS1dJxvXupOZeLj5FkbbK3Sfapa6cacfgG2jSHUXF9CfsnEGfvsKxaEn21s2hKVmPCceJbc+Wka8x1eupSPSwtktTwj9XL3PX1V6srIWkHokF6Zv8"
                    "8GhlvH6Yayxk5B9oa+nVc5n2mSVjXXJfccehvs13thLXNOxdMrXR+weqyx6NXTOxwr0uQ+UbmFFTnXUOdZmr8s6n6CyhcawGZTXGdI/OdeXJtoM2LY3fIsCQ"
                    "/GrVxG+Q1U6oKjqinzap/qq9STz/ANT79OlN14LiAkwGODJOuGxbhlc7LX2gDpBlubqMILuqvSgfsF7mtvL+kL0u/wCjj0kxtu/XMPIubPHKcwJBqWtBcijZ"
                    "h8ee5DBWxUdWWdjT0/rbNo7f2e6XYRoUp0BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE"
                    "BEBEBEBEBEBEBEBEBFW29X/ptdw3kyn6g8dhpPGeWDj0OWsRa5ptNZyPQ1C/bnSZjcs3H2syw2pakxordS01WzsGym0sreXKyiBDjWQfRj6li9isj6b5W0P1"
                    "OGLsnx8SPeZJcXamJv1WAtLOnH3pGStDn9bm5FsccftVndNR/wDtDPR847M4b1jw1RwpZwQ8d5f7bWiGvlqzAMFlHtaAevJ1Gy42eQkgzUKA/wC1nPVCDkdF"
                    "FvaiypZ7W41lHcjOK0ajjukaXYs1tKTLuer5bbFhH2ZpN+O33kpBqSqe8MxY5kjXDbXNLf8AvAaBae/YPaSwga0Cdd+4rWpW7GNuVb9f/t6kzZGMLulsrNOb"
                    "LXc4dxHZhfJXlH2j25XloDhsR5XMCZWy59dYN9k6vlvwJaUkpKEyIri2HSaNZJNyOtSDWw8aUpfYNt5Bm26g1ZEHNkaxzN9L2l42PIcNgnt2PxrZI146hoSN"
                    "x9iC1BXsVj1V7MUNmBxILjFM1sjA8NLg2RrXASxhxMUgfG7TmuA6lISelb8mezP42RF5M96I9H3fG/Gj7dq/tHyO41ryQO+tHXcdvBBB76I2QO59+Fw/B8gA"
                    "ef8AkO50QD3/AOW+/X5iySR9qdq1olHszUZkXnyZ61v8fHb8meyH4T5+7Wj57DXb40NHY+deTrYC9iu3euo9uxIGwfk+e3b+vnv+y4KJaWWPW8W8rlp+thLU"
                    "ZIdNZx5kZ0yRJr5SEqSa4cuMZsvEgu9szbkMqbkssuo/N9u58jR8DtrfYka79Pg/5jwF6lmjSy2Pnxdtn+7WWgdbA0y15W7dBahL2kNnrykSRkgtd0uhkDop"
                    "ZWPzS/IrsigRr+pJRxp7alOx3O1UiDMQrtlwpJJNRJkMO7Q4REaXUmiU17kWSw65yHyO51/Qb/fsNdh87OvI773rOKC7h7k+IvjU9V7BHKwOENms9u4LMBdp"
                    "zoZotOaT3Y7qgkDZoZmM6PIhypEiPEgRJc+wmy48Cvra+M7MsrOysJDUStrKyEwk5E+zsZkhiDWV8RLsqdOeYiRW3n3m21cbE0NaF888scUMTHyzTyvbHBDD"
                    "Ex0ks88jz7cUEUbHPllkc2ONjHPe4NaSsyxkM92xBVgjdLNMeljQ1x7AFz5H6D+iKJjXSTSEARxMdI8tYwkWrejTpoa6ReDmcMtlRJHLeeSIOXcz2UJcd5lr"
                    "Imo0hqiwGFOiy7CJPoeLK2xsKGFZQJz1dkGU2eb5lXGzX5RBr6+BXNuWP9Q+UOzoa9mCxzJ8fxeCZj43tx0krJLOUkhkaySG5n5Ya9uxDLEyWpQrYnHTRNs0"
                    "rMk/j+pPIYIXt4liLHuYzE2fcvTRStkiymbihNaxYEkTWRTU6j3WI6B3M5ps2yLl6ozGurbCPK7l7M/Oi8n50f5/H58eCLZkfgi/HnRt00jX9QP69vn4H58a"
                    "Hcja0lJJs7JBJ1vY1obP76/cnuQB3+AdB/VEIldJ5GpJmZcj8bGRr0faooOVNkrXxv7+3wR+FmrZnvWyfRUj+3x7/wD0HzegPBBuYknZ8eBvZI79hrWznnE3"
                    "/wC8taD3LIQW+NAYvMbGtE63pw8HZG+kdlXGcTsvyRmRGezI+7aSV4Lzoi1sjIjPz+CEwe37EAee/bXbx2JJG9tHctA0T5WfRnR8bBJ0P6b0NDX518Dtv4X8"
                    "Q7CfTWdZd0tnPpLyhtau+oruqcRHtaK9o58e0pLyokqaeKLbUtnDiWdZL7FqjTosZ8kmbZEfVtVoLtWzUtww2qtytYq2a07BJBZrWYZILME8bvtkhsQySRSx"
                    "k6ex7mk/dpetjrlnH24LtV5ZPA8vY4SSsDwWlkkUjoJYpvamje+GURyRuMb3ta9pcSLVHSj1FVvVDwvScjEmFDzWqkHh/LGPwGno0ah5BrocSTMlwWHIsVpn"
                    "Gc1rZcPNsRRGOXBra+3m4SdrZ3+DZOqNB7lnF7HDM/awUxmfUcwX8Hcmc1772Fmke2EyPEkrpLuNlY/F5Mv6ZpbFePJur1q2Xosdrb1A4wzD5CPLY4Ndgc+2"
                    "S5Razo3jLPW8XsVabHHHFG6OeKaWixscTZaZLKwmjpTTuz29ZUOO0t1nGWm63i+NNpddZZUwUy6s3XEM11DVNyXI7MuysZTrUSM068zHbeeKTPfi1sedLjeO"
                    "K2RyN6jgcMGvy2VeWRve2Uw0arGuksZC4+Jsj4ataKN8sj2sdI9jParsltS14ZcIo16xjs5TKl0WJxkbZbTmGP3bEhcxsdWrHI9gmsTPIjjj6g0OcJZnR1Yp"
                    "5o43c0za/wA8yq0zK8d9m0s3S9iPDfkfR01bHNSa6iqifccdRV1jB+02hZ7lvOTLGalc+ynvPSdwWAx3HcRUwlFhfUqxnrmnjiE12zIAbN+4Yw1jrVqQdTy0"
                    "D2o2Q1oSytWrsZrfL5efNZGfIzARGR3TBXidIYqdaP8A7GtX91z3iKJhIPU4ukeXySl0skr3cOm6Q8lDNtGKQ2ktJkspJD7ZK/uUskm2XkyLvVHW33ESSTGc"
                    "V5HfFIxkuqSmNzt7ieSWOP7O07wewEgdry6Ro8ecZA/QeBr4d4d4LTto0NDx9o7jwD/MP2OsKS19TVSUTGlFv2TWlLyO7uUaFbIiQoiMjJqQUd1JEe+9Xgvw"
                    "WvbcIrUbon/L9HpI+0bGvI3vboy9hIaPtHY8XM23qYdg6PSXDYHYj5B2O2mu6SCdEnWh2XF6c4yHJ8ppSJUlJNMtLIiNmMjSnFGnwpDkhzXcSi+1ths9kl1a"
                    "R5eUuCVzYInh0URLnOaSeuR3YAHwWxt8EeXyO7bYCvrA0663bDv5WA60GjZ386LifxoaB+ST2OfMgVEGwt7aY1X1FVBmWdpOfUkm4lfXxlyp0gyWZd5tRmXn"
                    "ENp2484km20mpaEn5OpJCyKGN0s8skcMEbB90k0zvbiYSAekOkexpcdNYPuOgCu9BHJYlZBCx0k08giiYGkl0khDWMA+SXFre+tb0QfLdP8ApK4au+urrCx3"
                    "FrKDMj0WXZO3kOcst3hwbLEuIcaQ25axIlrWl+4Rp0LGIlfhlJfVsRRN51c45NmuwkWb8xv7etfOovRn0mvZGnbYMyKr8Pg3iED9XyLJNkJv+w/q+yGU2cnN"
                    "E95YyGH2GuP900y0+n/0sj9TfU7inDfbNnjmJYc1yaZsUnsT4bF2IJb0b5JYyxrs5fsRVI4JR7pr3JhGwx05DHfRiRItfEiwIMdmJChR2IkOLHbS1HjRYzSW"
                    "Y8dhpBEhtllpCG2m0ESUISlKSIiIhSE975Xvkkc58kjnPe9xLnPe8lznOJ7lznEkk9ySSVfExjY2NjY0MYxrWMa0aa1rQA1rQOwAAAAHgBfQOK5ICICICICI"
                    "CICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICLBHUvwTRdSPCmb8RXj0eA5kUB"
                    "mTjl9Ihvz04vmVLKZt8SyQ4cWfVSp0aqvYcN21qWbSvRkFIqzx2bJTXW0xC8w4DzLIcA5fguW43qfPh7rJpa4kMTbtKQGG9Rkf0v6WW6kksPX0PMTntmY33I"
                    "2EYL6l8DxXqbwXk3Bsz9lPkOMnptshgkloXNCWhkq7Tr/eMfcjgtRaI6nRdDj0PcDSpyzHrfFry8x6/rpFRe0Fxa0V3VykmUisuaWwk1NvWPq7UIVIrbOJMr"
                    "5C2lKaOTGeJla0ERneVx7M0uRYTEZvGyssUctj6mRqysLdPhtwMsxHY6iD0yNJB05v8AK4B4IX82XJ+PZPiPJM5xbNQOq5fjuVvYfJQub0ObapTvgc5odo+1"
                    "O1osQPOvcryxSN+2Ru9Mud8Y9mTEy6Kk+yd7NVckalLNMtiPqpmdhl4S7CjLrnl96UNrgVzZJ9yWaVZdRlBBgcfBL2fGwSesAg6H3Hr15+477NGs39PMv7jJ"
                    "8JM776/XdoEBob7Eku7lcHYPVHPKy3E0Nc9ws3XuIZA1auyUGRK15Lyki14/O9bUX9p7NJmZn438HoeiQf6AnZ+Pg60PjZ/B6dE9uy29A7x++iT57/g/5jR0"
                    "NEedFdekIPuURmR/O/BfPzpJHvx8GXx58b2Wj/COw/f4PfXb/T5P42NeNgL2oH6AI3r9/wDxP/PY34799jfASmjNSiPzszI/BH4ItHrZ/kt6PZ/BaMzIfnbs"
                    "APjsO/c9/wDyI/z+N6368EmgO+u2/G/xv587H+h1obOvpxS7PHLVxuQazprQ22bBoyJaIzv9jFmhJH9i2EmTc329LkQy0onXIsRCBB+d7O/k62PjXfX7A9vA"
                    "PbS62exgzNBjog3+I0euWm/+V87Dp0tF7vlspHXW6zqKx49tliw586vpbdLLVjdSOrPkCoUqkxOdZVHAUOamQyxc5jDddq8n5VZaVqLZ1mEvpsMOwt9f1MVP"
                    "If8AVd803W5BxjjthIjB68c6NlsXp5ibH95YENrl8kD2u9mg9jZ6HHZnNk64pck10WUykLw0nDnH1ZGT087YZF0ZcieG8fFwt6M/yClIzHe9EGT4qgJGA5aq"
                    "6zV0LFktkhpW6jxLVmaLFa/DZo26b5lZj7klx5xWu5azVovxs9ESd70kvgiL+fg/GtGQxtjaxg/la0Dv+ddyf3J8n/4rQdmV0peSGgF3ho1099DW/DdbPSN6"
                    "7dlxK/uNJmWjJOi87P8Azs9novJaP5MtmWx3mnQI2O5J76A/yGvz37/08leeR32f2B2Doa2SdaJGgR28E+fBWh3qglvpMUve1FyNxsREZkZH3RslPejP8a0R"
                    "EZb2ZkWiPexvRU/+0DXYj+D5zq3of/ZeKH5B7b3sfHnZKz7iQb+ojOyT0RDpGhreNy5JJA7gdgCd6JDTo6VcZRGotaIy7SP7fPwZGWiM/wCPP5+PBH/ccwHe"
                    "CTryd7/ffYnxrYI3/wCZ0M+adE/kEnR7d9kfA/r89/n8L4XUdxlvwZl5Px2loz0r+DMj0fhJkRFvZbMi/ASAPHk73v8AodnXjXg9tnz8E9uN/S34d+3ff5Pf"
                    "uQNHfn+uwVtR0ZdSL/THzNAyOykSVcY5o3BxHl6tZYOWv+mFTVPVeaQI3udpX/G9pJdyKG6xHkWNnjEjNcIgqif1vInMa59TOG/2vwLm1WsGdxJlu4KV7nxt"
                    "ktBgbYxc72uj3Wy8DBTeJXvr1rrcdlXQyzYuGM+iIKuax9rjuRIFW+5s1SforF+PysbdVbUTp6lmRvuODK9iKtNjpLcRZWsZGtQdcZNNhz7ypFzy/h45ilgi"
                    "Vxth63P6dmV0xuVVZpZS2Gyfz5mVEkuwLirnxVqYwezbN1LuMzJV3Ecjnl06E1rT004qcNiX53Jxf+sXII4n3GyR9M+GoMd1wYIRyxtlq2IpA2bNwbZ/1lFD"
                    "SkbK3D1p3xW53kunJzcVqGSPH8et2KtglwByGYrP/TW7sjQGiSGJwkrUfeb1CFslsMqzZC3AsB78n+PBfBmezPfb43ruLRFvZkZb2XnY2P3H9e4P79x22RvX"
                    "cgfgjR+N4MA0NA+4kHRHnwQe5APYHWxreu3ggj8lpI/uM9n8bMiMtEezNOyPWvg9f/JkXjkCddxr86I0e/bq/J+NDY15Hfs8eDs/BA8DyN9zr8bAHc6I77Xz"
                    "IW9GdJ2Mt5t4uwkmwZk4syVtLRknwvuWRETaiU26rsSaDSPo8NlY5sgY9h7lj9EDsfuHV/KA3t1A9TW7IcDvf47be+yCO/b+be3eddiPgjwR38+M+xiUllqO"
                    "+olyG2m2lvJJKUuONoJK1JLREnvWSjQRESdKIiMvgYFKR1PkjBbG5zntYe5Y1xJbs7O+lpG9nexvRXoR68OIJboE+O47k9j4BB8DWiB37rUTq95AXTUFHxlX"
                    "vuNTswdjX2RqbSwsk4lUWSirK9xZ977J32UQCkpcjmhz2MQnwn++HaKS9l3Ccb+quWcvIxrocYySvVDw4byNiDcsrQC1rjVpTdOn9Td5KORmpq4c3NuK0Bu7"
                    "lJGtcK7X0qnU9zT+pswj9RN0AObIK9OQRua9zQ116vOzqfH9k8voVdMDnHfC+WdReVU0mFlfMM9zG8NRbVcWLPreNcRsH2Z0+CqRFRdwEZxmjU457apJVGR0"
                    "GC8e5JXR3I78afLrk+s71L/tXz2rwzH2GyYjg8UkVkxPeY5+QX2QvviQB7onux0McFKNwHVDMb0fYueDct9E3piOJenVjnGRrhmb9QpYbld7xuWDi9H3GYWJ"
                    "hOiyO/JNby5aGtc5tyu2VzxBCI54hDVTTQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQ"
                    "EQEQEQEQEQEQEQEQEQEQEVbL1g+nWHgfJNDz5j0VMWh5jeKjytlK7BxuPyhj1R77EqMUiTLhxf6wwanesTqamJTwY03j3K8nmFY3eWWcs7HPow9TZL1C/wCm"
                    "mTm9ybENlyvH3SPY0/wueYfrqLAQ1zxUvTCdhc+R7mZF0bQyGo0CpX/aGejzMdk8N6y4WqI6+afX43zD2ox0MycMThgszO7Q6XXKsT8PO/Z65q2KjDS+ZzjB"
                    "vk1JEyCosqOcpSIlnEXGecbIluMK70PR5jPcRpN6DKaYmR/+JvR2iXtOyVP2Jxje2RoJcx3buAPBb0nR0A8EsPbsHHsfKrPx12fF3qeQrBpmpzNmax5+yUaM"
                    "csLw0giKzE+WvL0u6hHI7RB0VHfc1syrn2FVYNIZn1sp+BObac95gpEZakOey6ZJJ2MtSTdivaSl6K40+kuxxG8gb0vY17CSHDbDruAR2JB8HwCO33Ag/cCR"
                    "JrH24LletcqSOkq24Y7FZzm9LzDM0OaHsaXdEzNiOaME+3Mx8ZO2u106Uk/4IiMjMtl4LZ7MvBeT+CPez/B7+4x+689yTseBvfY61vetDtsDx3GtlZJXdsjv"
                    "+Nd977b7dv3O/wDkR5XBPklKtFsy34PfcZfwRfjwei/7F5LfgfgaXdv28d/Hnv8AGvnXfv3/AHHqRv7b0N/j99kdz3O9fnsdkg+N5z6XemzJOqnmzFeJKCU/"
                    "S1ctMrI+QswYZjuHg/GlDIr0ZZlMdExh+DIuSVZ1eOYfWzyKDbZzkmLVM9yPXTps2NhPqDzKrwLjNvNzRMtW3Pjx2Hx7nStbks5bZO6jQe6IOkjrhlazeyE7"
                    "Gukr4uhkLUbJXwNhd72LigMeRyF6RsGNwtF2QvzGMWC1nUY60H6b9djX2Dbte3XjgivU5bUjm061mK7ZrF1uStqcYwahx3BMIpIeO4RhlLUYxjePV5F9NU0F"
                    "DAZq6mElw/8AdlvMQY7CZE6T7kyzl/UWNg87MmyX1whibcuyWMjkbUtvJ35p7du7P1F9i3amfPYk6SSImPle/wBuCMiKtD7deBrYII2NjnyrkUudzuRyszWM"
                    "N+y+UwxDUcUI+2GNvYGZ0cTW+7YkBntzia7ZkltWZZXci6ku3vTpSHD7kmX8H/BKLe/5L/ipPnRjk09w09i0aIP7Dvsj/Tz3B7b7LHZAC3ra4EEdQIHkO7fJ"
                    "2N7PYE9xryO3GrLyRHretfgi0ZHoyI/JF8ls/BkRbLzodpvwd68/P4+O3yfGhsgrqSb34862PA337nR7HZBd+/khaE+qAaU9Jpl4+7knjnZEWu7sh5Qv/wBy"
                    "7UdxkZeNF5+SPZXosCefn9sLmtHfjdvFa8/voD/6ut5zxLZsx6I7NiAcOxA/huVaBrQ2T1DXnuNjsQRXJUoz7TI/Gi153rfyStEWz8GRFvxtJf3GZiXzmt2f"
                    "6kEa+Pgd/jRH7Egn9jsFoPjfkk/sO/8A4a7/AI7E7PZfIs0lpRaMvBn48Ge/8l8b0Rl/y35IjMyHD+Ua7kdzrff/AC7b1odz3A0e357LdnfY7P5O/P57+fyT"
                    "2GxruNrpWQ2aj9yEwZEpW/qHS8GTZIJJspMtEbizUr3jRpKEbbM+9TpN/NztjQP9QNjpJJJA2dkbP3d9jY0NnYyfEURptqXuB3iYe4c7ZIkLTv7Ga/ug4bc7"
                    "+8A6Gs65Nuh3mRWcYU/wlkEh53KuM6s7TBpTyn33LziiItmJKxr3lo9lmXxVMlxG6SCl5x1/jK1iV9VBhUXEllLf19maf8Hyotsa1mNzUxbIAAxlPOyOdL7m"
                    "mggxZxrXve93QBmoXue+e5nWBmmPXbhxfMz1Bx8fXJacypyaNroxI+6I/wDdsuGmUWZ/1teBzL744JGR2681y5ZidkasLt1zTozIj2ku4i2k/Ba2XwZ/yRkS"
                    "dlsy2Xx3cDrufB0C47G9/wD1QP8AEf2OwPEb2PLmt7a3vQHyfJII2e/fvrQPV30SB/KlGRGRl43532lo9EST2Sf8fb5+Nmf5I/0AaHbXbQ1/XZ3+T5JHb8eA"
                    "v0k9ge/jW9kAHf57aJ2T30ddt7IHJYxE+su4pqL7YSVWS1no9KjLQmMhXgy//i3Iy1JP+9DTqU9xbIuvlJfZpS/8U7hWaP8AFqUO9zR2P/cNlHYdi9utbC5R"
                    "fdIwOA0B7h7b7M1r4d26teWjfT3A7rLD0uHBjzrS0l/QVVVBnW9vPU2t1utqauG/ZWti8hsjUtiBXRpEx8kGavaZcMiJRFrD39YDGRMEs00kdevEXBnvWJ3s"
                    "grwhxGmumneyJhIA6ngHtsr0IY5ZpoYImOlnnlirwRN7SSSzvEUUTersXSSPEbQ7sS7RPckRqcSYfk/Wj1XYRhFdJYobvmjkGFS1Knm//E43ilVDkTn3okJt"
                    "ZotbPAuNMctsssq9p6K1eqxi8lvyIRS5MtvKPUXldH0h9L83nJnRySYbFSNpiSQgZLkOQ6oqzWNc0u6LOVs+6I/MFKMsA9uD7ZV+knppLz7nfE/T2m2WSjPZ"
                    "jOatQ9ErIMRW3d5FkHuAi9uGy0TVKUjg98E97F13GQhgd6CmCYPinGWE4hxzgtNGx3CsCxmiw3EaCGp9yLS41jVZGp6SrYdlOvynW4NbDjRkvSn35LxN+7If"
                    "eeWtxVFd+9byd25kb877N2/ZnuW7Eh3JPZsyumnleQAOqSR7nHQABOgANBX1VKtejVrUqkTYKtSCKtWhYNMiggY2OKNu9nTGNa0bJJ1sknuu1jqLsICICICI"
                    "CICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICLCPUdwpS9Q3CnIPEd17"
                    "TJ5VRupobN5ya2mgzCscbtsNyNRV78aRIaocng1VpLrVuLgXUKNJpbeJPqLCfAk5Zwbl2R4JyzB8rxcj2WcReinexnT/AL1TfuG/Sd1gtDLtKSeq52g5gl9y"
                    "NzJGse3CvUbguG9SuEck4Pn4GT43kOMsUXOcD1VbJAloZCFw7ss469HXvV3jfTPXjJDgC00lMoobLHby7x66gFW3eOXFxjl/V+8iSupyHG7WbQ5FSvPpaZbe"
                    "l0l9W2VPNdbbSwqbAe9nubJCjvN4znqHJsFi89jLDbNDL0a16pK0tPVDYiZKwPALuhwa4dUbvvY8OY7TgQP5sOW8Xy/CuT57iWerPrZfjuVuYi9G9pYHS1Zi"
                    "1k8IcNuq3IfZtU3acJKs8EhJLnOWmfPuKfTyoWXxWzNmalqnuktE2lDUuOyo6afpJk4pc2A07Wvr17TRVdY0W35pd2ZUJNtdCf8ACTIzZPfZHuDxrTXkSaPc"
                    "9b9fawhZ36cZkSR2cHM5ofWc+/Q6uvb4JXgX64J+0NgsPjtxt/ne65ekcTHDpuqsxBJLtLejMzItkRGR7PyXnei/Jnsy15LRkfoHyT/Tv8/HftrWv3+SN632"
                    "3PWdvWx8DfnXfX41r99DyCNnwutzHUskpxfeokERklCHXnFmXhDTTLaXHnnl9xIbZabW666aW2m1uKSkfmwAHAt+d/cGgA67lxAAa0Alzj0AN2djXf3Kdea1"
                    "LDXgjMksz2xxRt6duc86A7loALvLnOa1o+5xY0FytYdDPS2XSNwkmDk1aiJznyh+1ZRy66pbLk3HnYUZ8sV4nRKhTpsB6HxoxZWn7zKr3lM2XIWQZtJjz7DH"
                    "4+JFBgnz7mB9ROTnI1pPd43h2TUOON7+1Yglka69ngHwxytkzj44P00cpf7OGp4s+3Xt2Mk2TFvUvkcdUR8OoTObWxNqSTJysFqL9fmdMjm64JpiwQY5zJIY"
                    "S/H468yaaapkGXDjKNxbQq2vZno9mav8ns/Gz3rZEf4I/gi0fkeK0gaHfsOn8Afv+fIG9/nutGOcXnvs9vnv8A6AIHfXnR32P4aV/bEg2TNtezYUajNJ7M0K"
                    "/Kk/ntP5UWtK8GR93zxkjD9PaCJAO3fWx/wn9x8Hfbt8L9ikLAWnuwu8HWmne+oeAf3A7kdzo91+7ySM0qSW0qLZH58koj0fgz+SPykyMt/GjIx82OPcHsQd"
                    "f018ePyPPz89tL8lb36m+CBrXbY2SQCNjZJ1s72Ce3jUffqjK7uk4iT8o5T46SpXktJKuy1BpI9edr0ZFvzryZlsxs/0WBHP99xvB5otIJG93cQe+ta+3ts7"
                    "GzrW9LO+GHdtzTo9MbSQfgNoZIdv3BIA+P5j20N1z1kfjRHrsLRGZklSfB+TIyLei8pP5MtlsvtOXkn46u5PfQ07ZO971++vx3AO/K2Cwj9h9xHjZHnx2351"
                    "o9+4799rq9zYnGJUdlZJkdv3KSZmTHyRHojP/eURF2oP/wDT8KNJ7SR/Jx7+Qdnv48H8u+P+Z+d+F7uNpCciaVn9zv7WuGjKe3bv3EY2QXD+c7aD/N09Bc8b"
                    "8H/yMzPu7vnfd/cezMzPyRmfzsfPet/ds/1Oyew1oDxruQfG9/sMsYPx20B8fHx+340P/EAg8piOX5Fx9luOZ5iMuNDybEraPdUkmbHVOrzlx+9tyJaV6XGT"
                    "saS4huy6e/q/dYRcUFjZ1TriGJzuurfqV8jUnpWhJ+nsR+3KInvilbrTmyRTM+6GxA8NlrzMcHwzMjlY4va1x+tinVyNO7jMhE+bHZOpNRvwROijlkqzt08R"
                    "PmjmibZhcI7VOWaCeKver17LoZTCGmffAc+x7lbCMd5IxRDjVDlEFyT+1SZbEyxxq5iPKgZDiNy9HJKFW2MW7MmvdkGywm4rirMnhx0U9/VOu6/iE8TpqVtw"
                    "/XUntgsvZG6OOwHRiSvdrsc95Fe9C5k0bBI9tab9RRfI6xSsagXzHi1vhfJL/Hbjvf8A08glpXI45mwX8dYYLFK7B7zY3FstdzC8OHXDMJK0jWTwysj7I5/a"
                    "e06ItmRn3EWjSf3F934/kyLxsj8lsu2067d/JHbZ189yBvR7AeR3BHxvG+3jt3GtnQH5PY/k7/PxvsDvIOEwzarJM1WyXYyTSgj3oo0FTsdHb8/3SFyvJFtS"
                    "EoUR7SnWP5qYPsxxAktrxgu7eZJw17u2hshjYyNj7SXAgEkLswNGi7WuonR/DWnQGt+S4u77I1o/BJ1w6xeQm6LC67jivkITc544mbcIRIdZkwcMoprUlJr9"
                    "s0GTeT5DGj1zCjUuPLraLKq6Q0pL5EPT4fR/U5GbJShwr4oBkJ6Q5s2QuRuj6AduO6lOSSWRpDXCS3RljcTG7WwuD433bFjLys6o6H9xUL4uuN96dpDnt9xp"
                    "Z10axfOCz+9rWZqE4LXFrhLt+n26Y5dZR8kdVuSxJsVV0mRxPxsxIbZbiTKyPKrbrPcmYMlSTnMHaQ6DFKma2uvkVdnTciVEuNLbksPtwS+uL1FdkeQ4b03o"
                    "2A+pg4485m2se47yl+v046rJ36QamOkdZLdEl2RaT90YJtx+iD04GPwed9Tr9cC3yGSTA4CR7CHtwuOs7ydphexrmjI5aEV/tL2SwYivYbI4T9EVlsQIU9EB"
                    "EBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBFW59Yzp3cx"
                    "Xkug6hqWCacd5SZjYvmEr6xom4fJWOUSmqdaor6jlH/VuA0LaIv0XZVVz/HVo5MJFvlUQ7Cxf6L/AFNFijf9MsnZaJqDpMtx0PYNyUbFguyVNrzoOdWuzttM"
                    "jIfLKy7KQRBTPRU3/tDfRx1bIYf1pwlQmveZX41zT2o+oR3IY9cezErR/KySvDLh7MwDtyx4mN2gSTBTk9BEyKmtKWd2tsWkRcZMhTRPHDkd6H4lghozL3Hq"
                    "2a1GnstmovddjJbVtKlEdg0UhY9jmkktdvQ+3bT9pZ4Ouppc3xpuwWnt3rKxmRmxWQp5GuC+SnOJhGHmMTxEOjnrPeN6juVpJasp6dMjlc4HajZvIUuuky4M"
                    "5hcadBlyYE6OtPlqTDfcjSG0+DNxCXm1+2si7XUpS43tKkrHv9TSA4O6mua1zT+xaHNJ+B2I3veiRs6ICldjbEFqKGxWkE1e1BFYryNI++GeNk0Lj2GnFjmd"
                    "bT9zCSx4DmkKVT0q+lJOb5k91U5/BUWDcT3RR+JK+Uy82xl/LsAifdzFleyYl0fD/fFlw0KbU1M5Pn0b8KYmTxtktVJjf67c7dXhHp7iZAbmYrsl5JNG5j3U"
                    "cDProxTiHiSKzyAtfHN2eGYGO7HNHH/F8bYblWQy7eIcdlyzXMbm8rHap4eNz+melXMTG2coI3Y+duxHPqrLHkMbbgmdTnjjyWNtXmV53Zs1cyQbrqk9pF2o"
                    "T2lpCCM+0jItlrSj+DPR7Iy0Wjj7BA2FnQ0EnuXHvsu8O/B+B+PgqM1qw+xJ7jyXdtMOyBoFxDe3zo7PnZPc99rjjUWj8kfb514PXj58eSLx5M/j4/Ox2wPy"
                    "CDvv513Pfz20fwCd78Lpb7Dt4B2QCD1AHf76Hb+vz20vndX9qvKd7LZnotl8kRbLRb1otER+TM/BEPoxuyBokaJ7fnvs+d/Ozsn47FfMk9967HYH7jWj2BI0"
                    "PJO/Pjv3/wAZmJbPscUXYovB/aZtmezUoi0R9ngt71oiI0+e7f5JCXAOYD1Df5+8AgAHROjo7Gge50fjRsvTtrjtv+HZ8EbBI79gPOwe3cjS0N9T7/c6VHe3"
                    "av8AzS4+Mj8Gkm/osq+7aT35M0oI/BH3kW9HsbH9GARz2PfkYTLAjvvq/V4kkaPbTQC7R8a38aGwOGvDbUhLv5+ljQCPuaaOTkIA+ektBPSewBPjYNcO1nlD"
                    "Sppv7pC0/O06Z8F9yiPf3bMuwlJNJmfeZaSklS5kPST22QTrsf8AUHx8kjydAbJHZbOoU/1BD5NthDj2BO5QCftGv8PY9buoOHZgPcluO5CjNS+5XkzMzUrZ"
                    "qVsyNRmoz8qPRn3d29mZaItqP4HuR2J357/nf7fHyT2+drM4gAAAOw0ABoAAdtAa7DwNeB+e/biXDIt6Pf8Af+SLXnZmfg/ktmevt0RfJeS4OPcfGu4P5Hzr"
                    "xsnx2J772NgLvRj/AMtef3/1/Hz3/wBVxTqtJPRkSdkXkvnwe/yZH8b/ACrZ689x74k/j/LYP77B35343+w3rS7zADoHz3/5f0Hz4+Nflbp9DvNaMF5AkcWZ"
                    "LaFFwXlmdFj1smdIlFX4pym1HVCxe2aYQZxYUPP0Jgcd5bPdKOw0s8DyS6mt0uBOtrxbktNzY4s1Xj6p8dG9uQjjbCJLeGJbJaL3lvvSy4ktkydGAGR7x/Eq"
                    "NWNk+WkkGtfVzhQ5dxg3qcJdyHjMc1qj7cTHS5DEucJMjjJHvliB/R7mytFh6xGf4rFBXsXMjCwS6vsPoUuO6y6xJJ02DYfaU06iSR+z7LrTqUracS6RtKbW"
                    "klpcSpCkd/cSfMZIwtEjHsfE5nue4x7XsMZHUHskaS1zSwh4c09Jb9wd06JhG1xO96/mLda0Wv2PtI3vq3sEHWydEADYzTCZj10JEd2UxEgVcP8A8TYTHERY"
                    "cWJBYNcuynyXDJEaM0y27MlvrNKGm0PPLUlKTMsKsTOlfJOGPfJPIXMhjBfLJJK7UcELANySOcWxRtGy5xawd/PoxxlvQxoeSeloaxpc9ziWhrWtZtz3uJ+x"
                    "jducSGgbOjEVYzsh6nuoSHCxeKcuw5GzDHMD41pLh2wZjJjWtvDxbAqOe7Vw7GbVNXVlYQ5d9IiwpDFbbX1zaPaipffRsPKZKh6dcFy2YyckP6fjuJv5fKSR"
                    "+2z9VcZE6eSON79B7prHtUKZe7rMQqxE6Y0iTXCuE5DL5Hi/BMMwuy2aylLFOmiDJRHdyE7Bk8o/rIjkgxdcTTuLWh8uOxsYjhfOWxP9C7gLhrFunrhjjXhX"
                    "DGIrVDx1idZjzcqLWxKlV7aNNHIyLLLGFCL2CvsyyKRa5Xkco1vP2F/c2VhLkSZUp59yh/lHIsjy3kWa5Nl5nz5HN5G1kbT3vc/pfZlc9sLC7uIa8ZZBAwAN"
                    "jhjjY0BrQBfdxvAY/i3H8NxvEwiDG4PG08XSjG+1elAyCMuJJLnvDOuR7iXPe5znEkkrLw8Fe2gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgI"
                    "gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIsMdQ3DdP1A8L8g8Q3b6IbGZUZx6+zWz9SilySrmRb3Er9cTvbKaihy"
                    "irp7lcBxaWZ6IKoUjbD7iTyrg/K7vB+W8f5ZQb7ljB5KvdMBe6MWq7XGO5Tc9v3MZcqST1XvAJY2YuA2AsM9Q+FYz1G4Ryjg+Y22hybD3MXLM1rXSVZJ4yat"
                    "2EO7CxRttguV3bBbNBG4EEAqkrmuKXOH5FfYpkcBymyPF7++xfIaVyXCs3Ka/wAauJlHe1SrOuccr7I6y1r5kErKCpcCx9gp0Fx2FIYdXedxLkVDlfH8RyPF"
                    "S+9jszj6mQqS9Loy6OzC2QB8Lj1RyMJMcsTj1xTMkicXPjcB/NfzTieY4FyvkXC+RQ+1muNZa3ibzGhwimfWk/urlZzgS+pfrviu03Au9ypYheXfcStbZfTH"
                    "lXUFzjx3hGAvNVczkWW5By3IHmHZcLDqrGoRy77O7Fg0nHdYgYlHcRAqHHoLF/klTU463NauMugm53OU8spcM4xlc9kY3zx49sYqU43Fk2QvWn+1Sx0Uga8x"
                    "GxcdGJLZjcynWkmuStNepJ07b9FspDebYw2SnZFBhmy5OOaaSvC04p5mtW6kYknrSWJoLEU84hrSTZGWC7K2jVsOotgdZCo8Xw/jPDsX4r45pGccwTAaVnGs"
                    "dqGvYU6zBivPvOzbWVHjxU22SXtjJn5DluQOsNSslyu4vcjmkc63lqXB+OXIZO5czuZtPvZbL2TfvWnGTT7EjGN9uuyWSV1ejUhjgo4yn7kjaGNq06LHOjqx"
                    "kfnOOUzclzFy4+P2IBqrRpjp6aGMrOeyhS+wNjc+GAh1mSJleKzdms2oq1Zlj2Gf6pW/O96L4LyZEWtf4Itb/O9f9tjvtGjofk62AO5Hf8knfbt8/PdYI4kn"
                    "e/O+x7b7ne+k99A+O+tO/ffyEo9GR/JH+N6/BERnvwfnW/JErf58j7ED/l/8v8j/AJfH+fyLgTsgjRO9/nz233A8n57Ht40vmcP+7RkZdxnvx+TIy8/we/jx"
                    "sy3v+Pq0b158eO58djoHwR8nWgCfK4Pd2/cbb5+QfxrZOjrY8eBva+FxZbUZbLSPx4IiLuVszMjLXjRloiIu4+4dhoOgO3dwB+T5A8fn5H57aHft1i4kk6Gi"
                    "B58bA2ST2A/PYaG9EKOb1OciZR0tqhxpHvOf6q4Fo2jM2ydTW5kvsJfd7bifbQ6sjT3pJbSUmW1GpG0fSalJFzSO0+Mxg4LKgdXZ+jbw7Q4t1tpDnNBD9Etc"
                    "T4A3nXpvI21nDV7SxtjlkkIIDWFtDItb0npPUD1uaQ0jXVvqJBaa7Ek1mXeovuUklGZnsy8kfkzLfgzPu0Wy/PjZCT5I+CSdn/8AC8ef6615J3oHZOlvyANa"
                    "eka0DrWgBrf7ePG/ntsfC4J9XlRKIz+PBmoz+df5It/brW/t14P+4fIg7P8AqDvpAGt+f2Hk6Gvj8L14x4G9dyfHzsb8EeNHvvuNa1pcQ4Z72RkZH3eSPX+f"
                    "Ok68mRH53oz3sz8D5O6d61332+PA1/r3OtAkgEAEbXejHY7P7jtv42d9/wDkf32uPd3syMta14M960WiNJ/OtaMlERJLZa8EOB0e/k9/u3v5P7bPfsdnYI76"
                    "PZd1nffydef38ePHb8DWu4/pxclhqQy9GeQ26xJadZfZcT3NusvIU0804nx3odZNTa0kZfb3f+ozHHejsEgggk+Na2Pju0gkne9+fAB13YJZIZY5o3OilgkZ"
                    "JHIw9L45I3CSJ7Hd9OY8BzT8du50p9OjvmBznvA6ezv5RS+QuOTh4/yY4+qOiTkMtpiWvCOQkspkOyX15nXQm2conyEofm8jY3nFmuLBrLugbe1nk4H4P9Zj"
                    "B1iraa6XCuDZC2Ok98bLuNc922NdjHvkjqRN6GRYmzjq8TXvp2niF/rLwmrxTlEGSxVeOvgOSvnu1K0ZIjx+QgbH/Escxoghjhignkhs02MNnox9yi2azJbd"
                    "YazsXWLyaWJcaIwSvcWi95RVKgSlJQ24mHglatr+pXHlmfey5kUl+Bi8VlbKkWVRKy5KHUu1mlc+JYw3cy25KwPqYgNmOy8NfkZR/ucbdDoeasYkuv27qhnb"
                    "jnhvRN28z0+xQs5CTMS9o8MY3V2h5a+TJSiT9G9oBD+ikIprpkZoR2YqbJPtnAfuB+n+6X5uc885N1O3kWczi3B1PNo8Pl6aRCuOTuQaWxoZLTEhDi33Dw3j"
                    "qfkLl9VSo0dp1XI+B20OQ6qFIZKLX1x+pDaGBxHpnReBaz88Wezoa8fbisfZeMbA5gIdu1k4XWQX/YBQGg9x6orP/oo9Pv1+dz3qVdiaYMHHJxvBF0b+o5O5"
                    "DFPmLbHvZ7ThDQlq043wOMrX2bsMpjbts1wQVjKyRARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARA"
                    "RARARARARARARARARARARARARARARARARARARV0fWJ6al49mNF1H4vBeKmz9TON8hpjxYP0kHOqWnZYoL6U4wz+7qdy3D6ZNTJlS3Dx+pd49qmEnGusvIrKx"
                    "D6K/UoSsyfppk5x11mS5njpkkkBNZ0pOTx8fU8xf3FiVtyGKNnvytt3pHdUVUdFUP+0U9IhWfgfWjD1NxzOg4tzP2mElr+7uOZWRrSGNDv8AesRanc0Oc9+H"
                    "hMhAY0Yf6LuOI/G+MNZBcxfYzjlmsjzl+80/HlY9hZIO0xbG5CXXloTMyJ1qLldv7LbfulY4lT2LDNlir6hs71l5DJyfKyVqj/cwXE7MleLTo5GX8sXfpcrk"
                    "oixvV7FBr5sXV63BzRWytuJ0lbJxAQe4rCzjtOPFe67+JZ+BlnKuaXtZBEwSWcbinH2me8JHNq3nAOt1jPPTMcsE0NiM5+ydP01s+ZkZFJabdLykvJmaF6Ml"
                    "GRERp8n8pV/P/LC8Q73abO+zE5zD5PYfc3ewN7B/zHf9l4+ZaYb0mgdSNa8eNdwWu+RsDpOvzsjZGiOBNwla8F5P5MyLyXnz9xn+fzsvn5I/HodJH/Pton8j"
                    "8Ed9dj2Xkk9vPcd+5+P5u2wdN7b7EEfneivwUvx+db0ReCLwR6/Pyfx41st/HnX1a3/XXc/18k9vj9wvx7wA3ex57a/O+w38keXdydDwN64O0vIdelwlmbj+"
                    "/EZpRGvW/HvOF3IYSe96c24Zfc224RGQ9CpQnsOYWjojI2ZXt03trfQ3e3n4+37QeznN8roWbscALSS5+xprfJO/Bd9wbpui3q0SO7QfCx1YXUueZpdWbbCj"
                    "JJx2T7Wz+SI3FFtbxmR+e8ybPRKQ02WiGS1aENcAsaHSDbhK8beB4IYD9sfjt0jqAJDnuPdeFPblm2C4tbsDpb2BOtnqO9kbI3vYOyAAVoD6kTai6YFOGSe3"
                    "/U/BTSRF52VdmCfg/BGZGeu3atbIzSlRkM+9Ong8uY0b2cNkwd60d2sOdflw23ZB0NjY2Qtl+kbT/aft2aKVk62dE/prH4IBOjsbB3rfbtuA5xaTSW0nv20l"
                    "4+DLfb5Mz1vf4VrRGRGZH2kcge+v2BJIPf48jXn58fO+/nUk2NcDsEb6nHZB87Pbx/TWj+exJK4h9BHs9H+S8Ek/g9fyXj4MiTsvnZ7MyH4SPA7HewQTvRH/"
                    "AJaG/wDw1tehE5w7Ejt/Tz+37d+/7/GlxT7akbMyNSdn9ydKIy3otH+NH4Mz8lru32mRH8ifPnydd+2tHewdnQO/xsDtvXb0IiD2BHVrxvRBB8fGz/8AHxvZ"
                    "4pwvJ/ye96PZH5PZ+Na/PaRGZ+dfyY4n4J/J8/BB8d9fnR76J2fI2u8weO4+P6/Pbeu2z57DsQP2PyqRsy8EkiJXnZGZn2l8/wCNkRkfgvtP414+fVr42fH+"
                    "oO/n4Gzrz4/K+4dodh+/jv28eewW4fSnyA5wXmtPyDKfcYpbFh6nz6IlZJTN48s1xJNgwtKSWpb1NIi12cVKEe26d7QVcV5z6F6ay75Wew8WXx0sZ0LUDXW8"
                    "fM5jSa96Fj/ZcC7+WOyx0lK04acaNyyGdMha4ak9RK7uVVLfHYWe87qZ/DWtLt/xuJz44ZGBr2B3vFz8U/3DJEyCzJbET54K7m5B5r5LY5Z5NyTP2pjacWkf"
                    "T1uISPpZKERMBokvN0856Ao3JLEmxYXOyq5r0KW41eXttHjoJv2W0fuJqwcdwEDb0wrmGGxk8tZdMHxRTSRtnuSe6AwPr04mNqwydDXOqVIS4l5LjiuBwNmh"
                    "Tx/G6VSWfJxytrzU2iMT3M7dkZHNAwsfJG57rLYcfWl6nMfBWryHbSSb+vp69LzHSL0ocYcTTIVVHzl2tLMeVZlShamrLknKGY0u9bXPfZjzLiPi8NqqwOkt"
                    "p0aJMl4ziVH70KD7ZQ2KOPWDnsvqV6i8n5aXSilfyMsWHgmdt1bDVXGDGwkBzwxxrsbNKxr3tbPLL0vcNE3kelHBoPTj0/43xKItksY+gyTJ2GtDf1eXtk2s"
                    "pZIAboSXJZegdI1GGAjYJO641mtiICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICLD/P3EFJz5w1yLxDfJrSiZxjE+phzLanj38CnvUpTNxrIXqWW4zHtDx3IotXeswXX2EyH65pv32DMnUZPwzlN/hXKsFynGvl"
                    "bawmRr3Q2Gd9Z9iBrui3T99gLo2Xaj56krgHf3U7wWuBLTiPPOHYz1A4dyPhuYaTQ5FireNlkYXNlqyTRn9NervYWvjtULTYblWVjmvisQRSMc17GuEANJma"
                    "7azVGt4SMazKK4l+TUfUPOssWzCylyYMN+S2y+cyplNOMT66YyzOivRnmnUPJYccKyWXFMjxde1TlORw1uvGWWQwB5qTRiOOWdsYLDFZiex0FmHcMrZGPHtu"
                    "kaxfzq3xbx+bymNyVY43OYfK3aNyB3UOjJ4y7JBZY3qfI4OhswSB0T3F0ZYQC9o6lmXLkt2NbX3kMkmzJbbcI2j7ibJ9Kvejb15VEltuRV78odZcSfnuMYhh"
                    "S6ratY+bqD4XOb9411mM/ZNr4E0LmzDR0WPafGgvU5BHHarV8jE1pZMxrx0HbYxI0e5D5HeCdj4X+CHsd8grGz01iGjukuk1vu7C2pS3DIi2SWy2pZ+SJWi7"
                    "Ul/epCdmMpjglmd0xML9a6iAA1vny7fS3xsbO3a7AnSw+SVkYDpHdAJ+0HQLtbA0ARvXydEDY27Q79OssifkEtmIlUdrWvc2ZyVEeiMyUnZR/jW2zW4REZpe"
                    "LfaPbqYyOLpfMRI//g7e0D37EHvLv/vADfYxkja8ie/1jpiHQwf4gdSHf/eGw0aB2B1EjvvuupPnsj+TUpRmozI+41KM1KM9mZdxq8mZmZmaj35Lx7UY8fgN"
                    "AAB7ADsAOw7a8DwNDtrW/Le7s4kaO9kE7Lie/wCwPnWz46gD3C4xxWlb8/8AwX+fB9vyf8kfnu8FrwY7bR21oH5H4/Pz4H7/ACO/yV137J770QOx2Bvt57aJ"
                    "2B/l8fC0k9SBO+ldKtK+3kvCFLPZESC+nytolGWvyaiR4SXlR/KvJZT6eDXMtbH/AM6MgB2HzPiXa7+SAPk/6bW2PSRw/tNHvy6nb0COxIr2Ce/Y7AbsDfYb"
                    "AHY7gCcP7ddxkeiV2pItqPwZaPZmXjz4Pfn4MhIRx0fjzofbvWux+Bsb/wBfJIIKkwwaP8ux4/Yf5/PyRv8AK490zI99ytlo/wDHk9duvzojIzM9mSj8mfjf"
                    "z+O4Gm9+wHnXY72NbIP5770NbK7jGjQIA3/4Ab8efH+S+FSlIUam1F+SUlSSUky0n7VJ0RKIiM9+En8+fB64k9J2PGz8edb35I89gdEHYI3279kNa8AOB+Ok"
                    "hxDm7I7g6Oj4/A3vQ8L4XPpXtkrtiuGatKI/9hZl5NKjWRe2rei+7aCPZGo/HdxOt6I0ewHbwSDrfnuf30Rv5+Oyz34wCNzs13Gj7wGtbAG/cB3sdO3EDu0b"
                    "Ov8AYFJJnWsKAtBqaffT7zpHptUVpPvyTSpJa7jjNu9h733qQREZHocRsu1+4Lv6b79yAe+i0bI2SN9+6/LWShrUbNsP1JDEfbaR93vv1HAHNcd9JndH1bH8"
                    "ocT4JXcc3szbJrHohkhLpsSbFKO0kojkpC4kJWy2k3FITKUjZ6ZRHLyzJUQ5Su+3oHbey7wNgdx2H7nZ+ddOvwsd4zS6jJmJ9l0YkhplwIJm+5liydnTi1rn"
                    "QNfogyPmJ0+EFTK+g702Tef+szFbS6rJ0jjzp2Njmq1sW4EhytTkdPYRv9Osbk26XUNVdjMzlVdmFTCdakHeUuEZnCaQlhmU+xFH6v8A1BdxH0plwdSyI8rz"
                    "e1Jgq7Q5olGGjjZNn5Q0HqMZhfWx7nAdIGSPU7sxr5F/TnwKDlfqpW5NZg9ylxKi3K2y5oMMmaLn1cCHDo6TKzVu7vfUx+NrPa3qPWy/gKglZigIgIgIgIgI"
                    "gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIq5nXbx6rjX"
                    "qDzJUV+rRW5k5B5Oo4tZDdhS61nLpVoq6O5eJ72513N5FpM3vGrGC1ES1UWNLGkJfsokyfLsp+m7kn9o+AUak4kfPgXT4CyZpGStniqMgdT9ph2YqseKuUan"
                    "tSb9yarZcz+7cGMox+uTgDeFetmTy1UdGP57SrctrNjYWsgyLy/G5qEFrWB0smQpnKyOY4kHKML3F211bhfOZ2bYhmWH2cyL+/YycO3pXDR2zJ9HbOSGprzy"
                    "VGbbiaS+YiKmyuxK3lZZCS4SnULdeyjnHH6+BzeDzdSCU47LCaleaCTDXv02xugYwj7gchjnzCvETpjcPO5v2uDWRzwF92W45l8bLMz9fipop6zXN3JPQvdf"
                    "uygaDQ2hkGRe+/s6V+XhA7sc49UkrdW66p1TnvGrThumalkojJJoc7/uI0aNJpUekmfb4LRH7cTWBjQwN6GjbQzQaQe+29PYgg7BA35KwWUlzj19XX36w7QP"
                    "YlunDyC3RAHwPGtO3x61efzryWy87Mtns/wRH8ERfhO/HwXaaCPIG/Pb4Hj9ydb8/k6/ddZwGwQSQdHRB79tkDZ32AG/jz476+N4zMjLu+fJme96L58n2lsi"
                    "+d635/kfaMdwdHt4H+nbWj2O+3/NcCex772QT233OgC4/GiNEHXcdu4XEvKNKiL/AOC8eNnst68+fJ+dEZ/BedH3o27HfzvRPz+D5AH/ACJ/5FdR5G9HR7Df"
                    "YbPnyNkAeO3jtrXYhaWeo6fd0q+T8f6jYMZkSSMz03lGjV5Ij8efBfP/AGMyyb07B/tk7t2GLyI7E/8A3bF/5+SRvv2Pja2x6SHXKYB1eadwjt5P6O0O/wAe"
                    "Orf/AOL/AIQGqv4vfj7u3RJMzI/CdfGvj7S/Bkf8l48d0gXDz23+x7HWiPOxr861+P8AOTzO53r867dzs7Ov/Ejz38bJXwukfcetK3/gvjfx5Mt93yRfPjwe"
                    "y8ce2gNbIOjokjwf2/yOz89+xXbZ4871s/n9+5Hz52PwRtfCstGZbItmfnyZmfjfjf5Lfg/k/wAeNjg5xJ79+nXjY7dwPka79+2+xB+O/ZjJOvz4BOv8t7B7"
                    "/wCp/wBVxrxeD13fg/8ABEZeTSne9GZl+daSZeDMiHEnY+CRofv235I2Pnt5AB+T3XeYNnf+uu34/q0AA+ex87+V2/B0/TrtbSUavoK+O4kzT5JJIaOZOWlP"
                    "d2+60w2wZEkiNz6g0FvexxjI+8u/lB7HySB3P/Jrd7APgn4WPcnd7ooUINfq7UrXaI0dueK9UbH+CSZ0ocTsNEfV2HddYJ9+ymv2ErapE2Up97Zkrs7zQTbK"
                    "VH27RHaQiMyR6JDTLaTIzIh8xshziBtwcdd9/sPA8dhvR2PBPk+4YoqVeGnXOoq8AiYe46i3fXKRr+aV5fK899yPcTolXxP09HTK3w90c2vNVxU2NdmXUpla"
                    "7pKrRD8WUnjXj160xfBI6IDiiZRCsLqRnuZVFqhpD93QZfTSnHHYTValqof6wOd/2s9VJ8JVmD8Xwio3Bwta4OYcpIW2c1KHN2Cf1JiqkbcWimGkh3U1tiX0"
                    "0cT/ALPenFXKzxht/llh+ald0lp/QO3DiWHrAf8A/MbWzkEBvXO8sBaQ98+IimpDICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICKLz1ReO3bjjLCuToMGC/IwW+lY1dTpb5R11eM58urS1NibUSZs+Rm+OY"
                    "VjUSEtJqSzlFhJZcaNpxuRKH6VuSsxnNr/H555o4eQY/3qsUbepk+SxAmlaybtuOFmKs5a057S3qmqV2v2NdMC/9oJwZ+f8ASTGcxqV2SWuC5+CW7IG7ljwO"
                    "f9vF3HM+5uxFlP4LNJtr/brMsyDoAe5QwcL5I5ivLuLvqeisQcpdXgV0qUpDbSq7L3YsOvUuQf2sMw8sZxW6kyF6JMWpkJWaW1rE9+dY1uX4XlY2xzSWcO1v"
                    "IqLYg50jbOGbNPZDI/Mj5sPJlqMcY24y3Iy0F4aqeuF2zS5NXrGeOGrn4pcDadM7or9ORMX6KSxL39qKvlosdcdNtojbXJkJhMrHZ8zusXV38gzQpLM0lSUm"
                    "ZEX/AIj3DRKaP42tK+x5e9GRye0tkRkMA47bbcxsenBz4CInDuf7stDoXgf8JbtjR337O+2wRwz1N1O9K3WhIS8HWtP6nCRujshzSOokga6xvuCuimryf5Pf"
                    "kj0RnvW+0iPx4/7aPevhIyIN8fj8gb7+O5Px+dEj/Xa8AnXz38+D+O4G/BHfR138d++/leMu1ZGWvJH8lv8AG/jfkyLR7PR6MiMtmQ+0e9jufBB8/wCRPbY/"
                    "+WydAr5PJI3sa2CBrxppPk62O++29b7LiXll3eS0XgzI9ERkfg9fyaTIt+fjwe9ee6wdh/UjtsnwT3/AI8fv3+djpPJ6hvRA7997OydDtrWiNk62QO7u5Wlf"
                    "qNqNfSioyUkkp5Fwcz8eT85GjRF87JSiWpRn8F5PfgZJ6fjp5l3B2cZkG/6yYx2j+xALQQBo/lba9ISDymt21unc123/APYVw78DW+kj9tDQHYqv84ei0Wtd"
                    "pfxtP42Z+T34IiIyPWyIyLWy38daO/zv/Qa0d6HfevjyT5IKlEweCe435Pzs71+O/fWzr8HRXyOH5PeiLZn4Ivu8kRqMzL8H/Pboi3+D3wJ+deCPO9Dt+djz"
                    "22PG970CCOy0d9fvv+m/HfXgdtgf0A2V8Dhlo/Gj86326Vo/kjPZePBFrRb+TMvJ8XfvsjwfnRB12G/6E7J7eO67Td72Pjz/AEP/AMv/AD1oHXFyFEnaj0kt"
                    "mXnREnW9mZmatGR7MyV+CP4NI4E7A2N9jv4J+R40D+O2/jewV3omnY+e3by7e+2gOx/Gh++l22zT+0YpWVXtqZl3Thypqd6WltCm5sgnCToyMjcr4C06NDjT"
                    "b5f+si4v+2MN2ep5JOu+gC0kEk+ewb579x3+fBpO/iGdu3upslfGsEFY9O2l7mvrQuZsa6SG2rbT5ZI+IjuQD2fhfjDJ+Z+U+O+JcLjuPZVyPmuLYLQdkY5L"
                    "TFxl13XY/WTJraVIQ3VxJ9lHk2kqQ4xFg17UqVLkR4zK5DeM8w5JU4fxbkPKLxLauBwt7JyEl3d9eF5rxDTXEma0YIGta1xc6VoDTsg5jgcJPybkWE47VG58"
                    "1k6lAAEAtiml3akHU5v3Q02WLAPUBqLfV2K9TnjXjzE+I+O8E4qwOuVUYTxth+N4JiNW5KlT3YGN4nTw6KliPz5zr86wkM10GOiRPnPvzZrxOSpb70h111VB"
                    "eRyFvK5C9lL8pnu5G3YvW5neZbNqZ888hHgdcj3HQ7Deh2AVulOpXoVKtGpG2GrTrw1a8TQA2OGCNsUTAAANNY0DsB4Xdh012UBEBEBEBEBEBEBEBEBEBEBE"
                    "BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBFjLmfjmFy5xRyDxtOZr3Cy/"
                    "Frapr3rOMiVFrL5cZb+N3yW1NPm3MxzIWay+rZbLSpUCxrYs6IaZUdlacg4pnp+L8lwXIa7pRJh8pSvubC8xvmggnY6zWLgW7jt1vdrSsJDZIpXxv2xzgcV5"
                    "zxalzfhvKOIZBkUlTkmCymHl91jZGRuv05YIrAa5kgElWZ8dmJwY5zJYmPaC5oCqJ5NBlspm1tjXyqOyaQuPOq31oOfTzOwik18h9hSm/wBwq3jXFkLZcWlq"
                    "ZHc9t1fZ3Hc9x69XuVqOQrTtuVbMcVmtOQfbsV5dPil9skgwTxOa4Nd3dHJ0ub+P5luRY29g8vlMRfjfWyWEyFvF3oT1RyV8jjbLqdyIA9Ja6GzBK0OGu7Nt"
                    "JBBO+9/buckceYjyEtEVNnc49EvrluCXtxI2Qx1yqnPIUZruP2osPJq69aismRqNivjkRnoaDxdJvF+SZnjQdMatDJTY6k6yfcnlxkjYrnHp5ZC0dc0+ItY5"
                    "80g7e7ZmPbutg8td/GatDkLRATmKEGYlMEbYIm3ZeuHORsrtLm144ctBkGRQbc1kcEbWEsDHHDhKSfwZ6Mt70fx2+DUZ62ev438b8ERjO9EHXgAg/HkHwB8b"
                    "+Nn877la1cNgdhvbex2dd/PyNgkb7eAB56SvmfMyLz9vn7deDM/BGet/yZl8fO9l5IfaMd/G9Dvvv/5b8fv3Hfewuu4kN7fPjv22Ow1rQG/J79vAHc64eQZk"
                    "oy8K3vZ/kj+4/BaP8dp/nWj2ezMd2MDtrY8efx2Hc7/fW9Aa0PGl0nnbtd9HQ/l122RvZ8AEHz48jWyVpd6i+/8ApOfUReP9Q8H+75Iu5+/3o+3aSNR68dxE"
                    "fx8nrI+Bj/1zA2d/wzIdtDvo44Dfnvog70Cdj/Pbfo9s8qqnRJNa85wJIAaaF7RI3o9wD3IHffY6Ar+qUR6/JGkjLwZmZH+CI9aLXn52RaJRaUoj34d67nWt"
                    "6J0fHz37+Dvzo72d9tSoa0j4AIJ/YAb8nv8A1Hjzsg9u3yLMy/lOiI0n9xF47j2ey35NX/8AkrSd60OB+ADs9Xft3+D8HxoeOx23ts+Oy0ft32d/nto6/rof"
                    "5bBHgr4HT0Rn5Iz14/BH4Ij8ePxrWzP58EfgcTvbT279t/nyXdiO3knwRr8b7dpg2Na/f/6mh/59j4A38fxWQ0WVlFiqSa2lOe5JQruSZxmC910j1syQ8aUs"
                    "bI/7ni8+SHHW3Nb26QQT41oEnp8a0f8AnsHuQSv29ZNKjZnaemQM6ISNEieU9Ebh37mP7pdED7YzrtsL7Mnl/XZDJ7jPsrWmq5OjM+5xlS3pazSnRpWUx99h"
                    "Rn4Ulgi+UkPnI7qkcRr7ftb8d293d9ednuPwAfOwevg6/wClw8JHd1x77jgR3ayTojrt346TWiimGvDpnd9lWAP05vTaxyx1lXfN13WWMjHemnB5uQVkxh9i"
                    "NWNcmcjNWWD4dEtorpqduYj2Iq5TuokeOx7NVkONY/ayJcaVHrmJkKPrc5s7CcFw/C6s/Tc5fkRavR7cHfwPCuE3wWtP6jKvqAl3UNVZW9Jc4OZKv6W+LfxL"
                    "l2W5RYj6q3HKIp03lrHD+KZMFriCWksfXx7Je7S17hcb9wj6myXqhVkp7oCICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICI"
                    "CICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICKsR19cTxeKuo/PGaumcqsdziSnk6l9p6XLhyHs9kzrHK5DkuW477VlYcmRu"
                    "QLJ6oad9qqq5VOmKzErJNfGatH+mbl7+R+m+Lq2LIsX+OyvwNgSdDZGwUWsON01rQTXhw82OqRSHZllrT7cZGPcaIPro9PRwz1xzOTr1/ZxPPKFXldIsDjGb"
                    "0jf4fyCMOLnASnK1X5GZv2NjGUhaIwx7C7h+mG9bs8SzfBn46Tcxi2h5bWuLcJTsmpypoqm5hNRz2f0VHd0sSwluEZpVNzvtWlJuI7sl9VaL6ubwOfZIQ3L0"
                    "p8RZaGgRxXcQ/wDWUp3SAg/qb+PuzV4W9iK/H+pu+l2o/cUlF/i96g/2Xy4O820wbc63Lj8uGwSuezWv0dK9BA0vGtT5kMeQHRa+a1r/ANpspkDSjTHdUTSj"
                    "13LjuaXFWZmW198ZxsnVf+snEkaTQoh2qdk3KsFnY6pGDrAB+2Vm2yt8nWpWu6e/ZvSfkLC7lcVp5YgDpjiGEkglh+5jiew7tIDgPLy5uu21wzxno/g9kRF8"
                    "/HyRfj+CL+NF+D+O/H8HvoOJJ/ft313+O/5386Pfz5B2cD5IH5HfwCR8D4Hnse/xrhpB/cWiIyIi2Zdvn+fgzP48ns9b3rZkY78fj58/O/n9tAf8t6/bz0Xa"
                    "3rvonXfzryNf5jR87I0Bsd9N/UY7U9I76i3suQsHNOz8b+qyAvjRH4SnWiVvfk9ER793gRJ5to+P4ZfIBHfscbofIP8AMdfHjzoBbf8ARwdXJ6x+7QrXQe43"
                    "3pXu5PfwdNG/g+O2hXzUe9FsjM06Pt7D7vgyLwRb34PW9+O0zIyMb/ce/wCRvt+2t67d9a7huvOx8EalOwdvn47EED+vfvvt+B3/AG0V8qlGRmovJfjZp8bP"
                    "x58FvZ62WvJ9xFovu+ewSTrv/wCOtb13/YHWxoHbT8DstHnffRPn87/8fn/M6Pc6+BxRER/48eCSnWyIjMzPfyZFsyMjIiPWzPuH47e/3Gj8jsDvQO9fG/A7"
                    "nwTsDsxjY+e/+n4+P30N7+B+wHZMZ9qur7vIpDCXkw2jRHJRkknFNJS4pgleO362a7AioPWyWnt2atji3bWuf/N07I/ca76Hcd3Ab/JBH9fGzfVct43DxSOY"
                    "bLw6V2t9HWXNbN+5r12W5yPljt60ulRfcWaVOrU886a3H3D8rcddUbjjhkREfeta1uKLejUrwZlpJdckhp1skuB/qepoJJ3o77jffv3/AK5NP0NHTGwRxsDW"
                    "RsHZrI4wGMYPA0xga0f90Deu+/Qe9BHpxRwV0A4bl9nVWtZmHURfW3MN0i5NlT6cblG1jXG37STcZhxjGbzB6Gq5AqY0h2Y4UrOrWaiQhmc3Ei03/VZzYcy9"
                    "Yc5HXmZLjeKxQcVoGN5kYX44ySZSTq63t6n5ixfaQzpDWRxsc33Gve6yL6f+LnjPpriHTR9F7Pvl5Bd3H7btX+gUY3AtDnGLGRU4y93UXuDnNd7ZYBNUI3Ld"
                    "aAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiA"
                    "iAiAiic9V7ig7/jbBeX66BOk2HH17JxbIHYsphmDGxLO/pERLKzgK7JFxLi5zTYnjdAhhTztYWcXr6Y300udJjSu+kzl5xHNchxaxKxlTk9ITV2vY4uOVw7Z"
                    "po2skHaGN+Mnycs5IAldVqhztxta6Av+0H9PWci9J8fzqtD15HgGXYbMg6ARgOQyV8beDtt6z7OSZh7Del4DGNsHocXgshb4JyZOJcw4e/ISp2uyl97j21Sl"
                    "4mElGzVcWurJD61qS2mJV5c3i95NWvXbEqHST5V3FP31Axjs1wnNNiLY7OIjbyWm4xmQ+5gmyWrcTGtDne9bwzstQha3ZM11vY6INPfBbbafJq1OUNfWzsUm"
                    "BsNdK2FgfkdNoTule4NDKmUbRuODyGvbVdEXM6+tm1/I9UplyLPNC0OMuuVkr7FJNvRuyI5qLRaNLhSmzPatqeabI/CSPUvF7YkbLX6gWvY21D9wPUNMjl0f"
                    "n7DE4AaOo3O0dkjscjqOika8jpex7q8zdAd29RaTvsAHB4Gyf52gdwAsRyD/ALkn8F8nsyL51vZ7/wDT8Ee//fW82i1sEd/xoA/Hj4+Pnxs78d1iMv8AL/Me"
                    "w+ADrev3GyfA7/AHY9hwz/8AcZ68EXx87Mj35P8AJ/zvfk9nsknruxn43r/LwPntrt/5fsdb6DwQexB8H586Ot+NEjR2SB4G99lp16ixH/0hSFb0Rch4P9m9"
                    "GZql5Ekj+POjVsyL5Mi/kjL3OCn/ANeDsbP8LvEEDeiDjDo7/Pfu78+d9juP0cP/AKzU2/Bgu6Oid6o5F2vgjuwgeO7vHhV8VHoi3/j8mZIIvhOjIvjZkovJ"
                    "qMjMiMjIb/PzseAf27fPf8nXjx+S1SnaPnvrvs9vk9j2OjvyBreh5Hx8rh77jVo9kXnfcReDIzMtbIy2X4+dGZdpjh8dtdu432P7jv8AI1vY6fz50vuwfg+S"
                    "e2td/wDP/Tz+T37Lj3lKIjMySfz5Ij8kXwejIiP+S7fjZF8lo+DgBs9wddh+N6350Sdb7HW967tOx3Im7IG9kn53+e3zsfHfx3799hc9kZqraGkoUoJDskzs"
                    "J5EozMkx1e6bLhb2aHrGQt1HkyI61PntLRcZC0RxRgHf8ztnvsAHwOxBc46Hj7fyF5WGH63LZPKlxdFABUqnXYmUdAkY4+HxU4Gsd/3bp/JWQOm/hrIeoPnb"
                    "iHhDF026brlPkPEMGZnUdQu7nY/DyO8h1tvl7lcgyS/UYTUPz8xyFx5TUWHj1FazpTzEWM8+jA/UfllfgvBOU8tn04YTFT2q7HPaz3rzi2HHQFz2OH99ekrx"
                    "npY52nnpY46athcO47Jy7luC41HrWWyMcFh7g9zY6cbHWbz+mMtd1tpwz+2etrfd9vrexpc4ep3iWK0WC4pjGE4tARV4zh2PUuK47WNuOut11Fj1bGqaiAh1"
                    "9br7iIdfDjx0uPOOOrS2SnFrWZqOhy1anu2rN21I6azbsTWrErzt0s9iR0s0jj8ufI9zifySrZIIIq0ENaBgjhrxRwQxtGmsiiYI42NHwGsaGgfAC7COuvqg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIsXc2ccM8u8S8g8bOS11r+W4xZVtXbNpS47SZAlr6vG75ppZk0+7RX8att2o0glxZLkJMeW27GcdbVkXEuQT8V5Ng+RVmh8mIyVa46Ekhtiux4bbqvIBIj"
                    "t1XTVpC0dQZK4t07RWI8+4jQ59wnlXC8nsUeTYLJYeZ42XQm5WkihssAcwmSrOY7Ef3N++JvcBU8b6ukMLn186vtqKW2cmHMqbaO9V5HQyCNbMqptIbhJkVl"
                    "9UO+5DsorhJfg2UZ9hwkuMmQurwWTgv06ORqSw269mKvbgsRETVLccjGSxTRvG2zVp2lr2HuJIng607v/MtncRkMBl8jhspXNXLYbIXMdkK7mPY6tex88tW1"
                    "E5jwHNdFYhlYQ9vx93faktRaHyZx7jGYKbUwvNscZmzO4iQ01lNXLlUuTIaJJn/4OHmtJcsR+4kqdistmbaCcIijYKv9luRZXChwk/gGUfWi1svdibUMN7FG"
                    "Qu/9/Pgr9KSXpLg2V7wC4t77M5BH/FadHMNYWx8ixsWUY5wb0ut9c1TIOYG7AjjzFO+xgJLhFG0ODXAga/SUmj3EOIUhxLikPJV4UhaFGlxCtn8pURkotaJS"
                    "T/JDZcJBLXNIc1zA5hH8rmub1NI7aIILTvfg72R3Gq5m+exBDukgu7hwPcHQJ3sEdwT2J7dguDkb7j0Zn4/nRHv5P5PZfO9eC2e9H4P0I9AAa8HwR+NHvvXf"
                    "Xwe+x2/bzJOoeRs+CR27negOw3s60N/gaA2Vpz6jSjT0iOGnylXI2DIc0XwX1GRrIi2WvK0/nRERqLZEZj3OBaPODve/4VeIOt7JfiwQe539pOho68kdtLcP"
                    "o40f2mqf4T+mvd++uo0L++wO9kAgdyNDe97VfI1Ef5M9l8HvZ73oz8n5LZeNEfzrXje/XHXg+P6f/ED50d7Pjv21Kkb8HW/+R7/nt/y/ovjecTpRlvzsvBlr"
                    "/wBiLR6ItkRGevwX5McDvuAfHfXne962Rvv2H50SDsdguzG09hrQ/f8Az2N70d/P+vYdl/dREOwt4bBpUtonilSO34JiLpzR/jTrvsxlaM/L5HrwRjiO7mge"
                    "P8Xjw3fY+P8AFr9u/ck7XHIz/o8falBAkLDDFvXeSYdA6f3ZGZJgD/8AcivjyKd+4ZFZOoJXsRVlVxtq7yUivUbTyyNR/wBr0xUx5J68odToiLST+TyHSF2x"
                    "2Ib58Ab3/q4EjRGt9zsbXZw1b9Jh6TD/AD2Wm9KNa062A+Jh1rToqwrxOH/FGfPdWKP01fT/ABuRerfkTne0S6uF05ceoKnbJlCSbznmJvIcOo5xy3O43WG8"
                    "GqOVYUmHHaS6T9hVyXpbLTRRZ0GPre5p/DeG4Dg9eUixybKnKZCP7h/1XhA0wNOnDYnyc8Uga4FoNLqDS4tMcsPpb4x+s5NmuVzRAw4SgMbRk7H/AH7KkPs6"
                    "aQekw0oGt6mkEi2W76eoG8eKwlOhARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARAR"
                    "ARARARARARARARARARARARARARARARARARPkEVWPr946a416quUKmPJcdj5guBy7UxFRj7oVPyW/cOT3ZEtKEsSnZvJmO8pLYZbQ29DqI9Y1JJ9xX1sq1T6Z"
                    "OVO5H6XYiKw5hn4/Yn47KRIDI9mLjqvqExEufGxmKuYyCN7j0yzRziNrGRdIoc+ufgJ4Z675jJV4DHjOc4+ry2tJ9oY6/O+annYgGtbrV+sLchIOzfa4vfIX"
                    "9P8APSrfFY4tnWCPqlPy8et4ec0xPLU4hFNkDMbHcggxkGWo8KpuqmjslpI/bdsM1kuaJ11al5B6t491bL4DkDPZZDlKM+AuljWNc6/jZJcljp5CNOknuUbl"
                    "+qCfuZXwUTdlrGgaC4hI2/xbJUPac6zg8iy/G8OeT/DsrHDUsCUOJbHFWu1qQhLD0PkyMod0vDTLzeeVqIN0t9CSQxZJ+pQRGZGUps225qNfhSlm1LX/AJmG"
                    "RHtJmXV47bM9FsbiXSVXGJxI8xODnQHe+4DQ+EDv2hH5AOJ5isIZ3ODQGTf3g76AcCBJ20dHqIkOjr7992+MZyU6M9l9pF53/g/HyRmWy/8A3eS/j5yuI/jz"
                    "8fg7Gux/bz37LG5Wjez+B+5G99ie3V+Pz5IPyNOvUZ7S6QpJqIy7eQ8EUSj/AOJnJvy7jIvBmZGpKdme0q0WjJI9vgh3zjs7zi7/AGPYkCTGDW/233HbvvZ0"
                    "e+4fRsb5NT8ncF3fx2/h+S0dEbGzoD56QddtBV7DUZJ8qPRls9aL+D8f5PZeD2Zn/JF4324nue3Y9iddu3yRoEb8OGtb8jY3KtrfyPn9z+fJ3rtr8f8AiN/G"
                    "8r+4yVszIy2fgj7tbJRaIz3vyZ/J+D7dDie3415OvHbsB8a6QDryO/juSe1E3eu3z8g/B8fIH7eNkEfC52gX+11VzkSyPbLBswz8kk3FONtskfgv9p6ykRml"
                    "bLRGwr+DJPEkta+XW9N7dh52G7J//CLR3PkHv27+TlGi9kMdiGHQllElgaB0wNc+QHfh8dOGxIzQ2RK0dwQV0FhHtoQ2gzWfhO+4lKcURduy2X3KWf8AdsjU"
                    "azP52ai6/cADZ79I1sb2d9vjZI33P+EH87GWyuDnOLtNBBOgOkDf+Q0AO2h8D86C9EP0IenqNwR6dvFt28cxeRdQE+056vXJTZtMHXZazAqOP/2htyLFkN1M"
                    "3jPHMOvU/UKllItbq3sIcpddOiNtUyfVBzN3MvWLkr4p2T47jjo+L4x8ZaWGDFGT9W8dD3s3Jk5rzzoh3SWCRrZA4CyX0N4z/Zn04wccsDoL2YY/O5Bjwfcb"
                    "PkQx0Mb9tadw0o6sRHdocx3Q5zOlxmREe1t5ARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARAR"
                    "ARARARARARARARARARARARARARARARARARARARARARARRKerZxwu24twDliPIJkuPslk4zeNPJJEJVLyCdbHrp70hDK3DsmMwo8bxqkYedahuKzSxSo1zFwm"
                    "1yy+kblIxXOcnxqUtEfJcYZqxLgHfxDCiayI2h3b230J7883T97nU4NbDVAD/aF+n/8AaL0rwvN60Tn3eBZxotOY0ud/AuSurY6852jrogyUGGlJeNRx+88S"
                    "RjrD4T+Dcvbw7l7EJcyY7Epchlu4LkBM9vtvV2ZGzVwFS1OKShFfT5Z/S+TT3O9JtxsfeWRHpJLsC59hXZvheZZBFHNexcTM/j+suL2WcKXWrLYAASbF3D/x"
                    "bF1m6PXNkYwSQ46p+4NfbS5JWpzOkFTOQy4OyY5BG1pv9Lac8+y1r4KmSZRvTRkkuFYey0zNYFvFn9KqRXWDJoUmXUOnLbSbZk6aGCUmW2ojLv8AvjqdWaC/"
                    "ueZYIy7iLWi+OX2xWqzw4GG4wQuId9gMmjC4aJH2yBgDu+o3yEaGwfWz1Iuinje1zJa7w/x94IOpRruRpm3EEDbmtHwda5SkkR/4SXjZ7JRFvwRfjRb/AD92"
                    "9fHktoQnY+CT50D23ruSdA77b8a8nue+vp26do/4QPxsnX7d9jx203e9AFaY+pAfb0hqTsy7uR8CI/HjSl5IszMv4SSSLx/gyLe1D3eB9+cEnuP4XkR5II1J"
                    "iwCCN78ne9aOj+AdxejLd8lr9yOitbePOz/uOQZo7AAaOskeTsD42q9Dij2REZfaRfg/tM9aLReDMtH/AG7LXgy+Rv1x7Ht4J7du3ca8fnwNnR7Dt8SpYAdf"
                    "v/T/ADO9A99Ak+SP9FxjylGatI7lGZdqNeVq39qEmXj7zJJFoj3v5P8A48T4J1vzrv8Aj8bIHk+AQdjtonZ7sYAGy4Nb8k+GtHlx130NH/ILnst7a+po6BBr"
                    "SszOfLNKjJLiIqfaQSy2e0Sp8mVISRbInYKTToy0OMx6Wxx/JPU7se4aCB3I79Ty52h+PkDZ8nj27mQymXeGFoAq1+ofc187hIda7B8FSCCE+C5lk+NlZM6U"
                    "+D7TqX6kuEuA6iVPr5HK/I+M4dKt6qGmxsseorKxYTlWWQYDiktSl4biqLrLXWHVpjnFpJCpLrUZDjqNceqPMG8B4Byjlx6few2Lnmose/p93JTNFXGxNcWu"
                    "PU+9PX39ruloc8sLWydO0eEcbPL+W4Dj3Q50OSyMDLp0S1lCHdm8ZAC09Bqwys11s6nvYxjw9436n9FSVWM0lPjlFBj1lHj9VX0lNWxG0sxK+qqojMCugxmU"
                    "ElDUeJEjsx2W0JJKG20pSREREKJZ5pbM01ieR0s9iWSaaV7i58ksry+SR7nEuc573FznOJJJJJJKtVjjZDHHFE1rI4mMjjY0BrWMY0NY1rQAGta0AAAAAAAD"
                    "S5UfJc0BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEB"
                    "EBEBEBEBEBEBEBFinnPjGPzNxByHxi89Xw5GX4xYV9Na2lW1dQqDJ2UJn4lkzlW8ptE13F8niVGQxWkusPfV1jC48mM+lt9vIeJcgn4rybB8iriRz8Pkqt18"
                    "UUroX2K8cg/VVPcaQWsuVTNVkPcGOZwcHNJacR59xGjz7hPKuF5JrDT5PgslhpXyMDxA67VkhhtNGiRLUndHZhe0dccsTHsLXtaRTaua2U2UuFOZXWSTJxmT"
                    "DS8SpdY640ZuQ3nU9ponwiX7DulJU1KbWRmlSfN3mEyUF6pSyFY/qIbEEFqF7mj2p2SNZIyTXh0UgIe3t0uY4faQ5fzHZfGXsDlb+KvROr5DDZG5j7sbm+2W"
                    "XcdZlrTs6WlwHRZgeOznHQ31H7SZRsezD+vsJw3PXJC5U/JaKM9kDq2UMGrMa152kzYm2UaQmKeXVlycAiSgnK1cN9KEoeJJRcu4b+z2bzXHmxtir4rITRY1"
                    "rXmQDC2GMvYLqefuMow1qkLB2em02aPrcWdR2tmbTcpHRzjDPIM1RgvWJLTYxM7Iuc6vl3EsLmmN+UhuGu8BjpK7oJHxxOeYhgXJK06u1lQ0pMmiMnYneojN"
                    "URzuWx2me1H7ZGphS96N1lw9kZaGycVaNqnFOSOvuybWgBO0tEmx4+7QkA0SGPHx3WqchAYLMsQ0GgAt33+xxJZ5B7AfbsADYd2OiDo36j7ZH0fPq0Wkcj4A"
                    "ZkXnzvI2zIiIy2WlGpRnrfbv/vlPBHAc5770cVkh48AzYsjY79vHbfb43tbY9GWn+0tYg6HsWgT1a/8AofkXa7bHcgEjsd6OjrYrwO60X2+e0jLZbPWvHd92"
                    "z3+d/bvx+Rv4nsPPyD89x2A/GwOx1/XyFKtg7nuD31vx2B7+O340PnRH5X0UsNM23jNKJSktLOUrtLZGbGjZLezUaVP+wgyLWyVo97Mi/Gfc8d/DgSSP2B7g"
                    "n8n9x28nyfjk7H6XHzO2GvkHsM3/APpSWyEfnphEp89tbHdcVkk9NlfWLqFm7HhulVxjMy7farzWy77ZF8tvTPq5SFbIzS+RHvwY68jtyPPctB6W+dgMPS46"
                    "12Bf1HQ+Xb/K9DC1TTxNNj2Bk1hn66cDuXPthkkfWD3D4qwrwPB3p0Lu/lWS/wBM106ozrqV5Y6kbVVe5VcD4FHxfHoEmvRInnn3L7tlXRshqrNaVogHj2DY"
                    "pntBasNE1NmMZ/A7X0w2prEqB31w80ZU4rxXgtaST385k5M/kGtc0MFDDxurVI5g1wc4WL12SWJpDmbove8B4gcJc/S7xo2M5yDlc0TTHjKUeHoyOH3fqcg9"
                    "tm4Yx0EAsrV67XSBwcBYLGlzXygXdBWepsoCICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICICICICICICICICICICICICICICICICKqr6hHFsji3qm5JitQqeHQ51Ihcq403UreU+qBn7tk9kMvIW3WWkMXk7lGm5JfbZhuSI6sfTRy"
                    "lvplyZMaNa59L/L28n9McVVlmmff4w6Tjtz3i0kxURE/Gfp9d/00OHsY6qOtoJsV7I7tDSaGPrp9PHcH9csvlYITHivUGrByykWNJiF94/QZ6EOI0JTkqzsh"
                    "LH1kgZONwDWvY0cZ0t5U3MxvPuPJsxxc+kkwuQcbiG0at01g5W4vmhJeJKkNR66zawqS1H2SnZeS2j/aazfUrJfVbEmtl+O8irxNbWyME3HcnL16P6+qLeVw"
                    "ZDOxe+xVfnYpJNEtixtOPqLRGBHritp2Q4tlsa73H2cDbZlK4EXVH/DclJXpX3vm3/dGG+MSK8A2H/qrk7WNdHYc/JWf1/1MKNYtp/3oJ+y+ZFtZw5KvsUZk"
                    "Wu1iUpPb4LtTKfWeyQesf47ZEViSs8kR2B1s76AniHcDffckQ7nwTFGBouG/Ey8JkaJQCXRt6XkaB6CSQfgHoed7J19xI7qN/wBSIlf9Hkg0mfajkjACcUZl"
                    "pBe9kbSj0ejPbnalP2mZEoz0Xgy2TwQgc6bs/wD0KyXg73/eYt34/wCHv+x7d1sL0Y7cmrA7G4rjekb0S3HZJx0dEEbHVvY3ojfkGu26pJEX3GpJlotb2rzo"
                    "tfao07Ik+D+D2RlrZnIB2+5PY7AGu/je96HfYB147+N67SsjH3fy6O9n8du//wAAdaP7b0ufppCaeou8gU4pt9llbULubJSfq0ajwD0RbWh6zlMNu72aW45r"
                    "PwRmXFpLGvk3vQHSD2HUOwB2PJe4fA/YeV5OTidkcjjMQ1ofG+QSWul2new4e9b+R0ujowSvjO9l8oaNk98cx2zQhKCNSln2/cajNZkZ9hGavjuP/kozL7iM"
                    "9bNQ6vgtbrTifkeT8FwOvJG9Eg/BJ7bzOVwc4uI6W7J0B9rQNnTR+B8Adx2HYAr0XvQu6eXen/05+IJNnGgx8m52eseoS+XAU840/A5EjVrfHTrq5LEZ9M1f"
                    "ElJgDltGNomYl25aR465DKESpFKv1H8zHNvVzk9yGYS47Dzt43iy0/3f6TDl8EskY/4bF91yyCdl3vbPSNMbZV6NcYPFfT7B05onRXb0TsxkGvAEjbWR1N7b"
                    "9EjcEHs1wBoNbEAdu6numAGi1tJARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARA"
                    "RARARARARARARARARARARARARARARARARARARARRC+rjxId/xxgXMNZTImT8DupmJ5NYtyI0NVbiOYtsyYFtNJx5p24KDmFHS4vUVzbUuRBc5CtrCMmLFO2e"
                    "XLn6QeXjEc4yXF7E72V+TUGz1I9Pe2TKYgTSti7HphbJj7F6xLL263UK0TidtCr4/wBon6eu5F6U4TndSJht8Az0YyEgaPcPH+S+1i7Lers9zYsyMJL0jYZH"
                    "78mgA5wg64fyqLhHKWJ3VjPTXUUqwcxjK5ThbjtYxlLLlHZS5hEZLOHQvy4WWmgiPcnHIaiQZoIhYZzbEzZ3ieWp1a/6rIV67MtiYmkNlfk8U9t6vDCdlomy"
                    "EUU+I32AiyM42A8lU88MvQ0eRVorcxhoZX3cPkJSHObDWyLDXFl8TNGcU5nQ3WQkOJlrxuaPca0iQm1rzadn1k9haFIJ+FMZcLtWky9xh9tWiSaTL7i8aUhX"
                    "lP3Fso807PWytbryNLXe1Ygkaepp30yRuG9tcPB+QR2OwQD79+q+GaxTsRuZJE6WtNE4EOa+Nz4pGuHfTmHY3oOB33DgCosfUsiqg9I1zCUajONyXgbRr7f7"
                    "0on3q0r+0yT2rQtJpIi3padkWz7tx+nszbHM60wA1LiL7gBvber+GBzdeSWlpaT3/lPnQ1lvpC32uU1IiO7Bfb+7v+rMuQTs+CO/Uflw13OzXHeUr/inaj7U"
                    "pJOzMz3pJERmaj7j0ki0fkteNEJDuB/bez52NntretbBJPjWhrsVKuIDuSdAAlxPbQ/m2e41oAl3j899rlsue+gqqWibWjvdUc2YgiIlmzER7Udxf3F3IlTH"
                    "pbxH+HoBnszSRFwmADWsafkP15GmghvfQ7Ek6Gxvp+O5HQ48z9XfyWUkaS1g/S13H+UPncJJWt7dn160VeMga/u7midErK/R/wADWHVH1QcF9P1cRn/qnyLj"
                    "2NW7rL6WZMHDzkHYZ9eRCN+Kb8rGcDr8lyhuK1JZkSipFssOIdWhZaz9WOYM4F6fct5O+SOObG4W1+ga8ncuTuNFLGRMADiXm5YhedsLGsY58n2RuI21wLjj"
                    "+V8w4/gwySSG1kq7rhaARHQquFu46V3U3UboIHwkg9RdK1rNvLQ71Q66ur6ivg1NTBh1dVVw41dWVldGYhV9dXwmERocGDDjIajRIcSM03HjRo7bbDDDaGmk"
                    "IbQlJUUSSSSyPlle+WWV7pJJJHOfJJI9xc973uJc973Euc5xLnOJJJJVpLGMja1jGtYxjWsYxjQ1rGNADWtaAA1rQAGtAAAAAGl9g4LkgIgIgIgIgIgIgIgI"
                    "gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIsS88"
                    "cYw+ZuG+SeMJjEJ5zL8Tta6odnpI49Zk7TJzsRvkqNmT7MrHMpiU9/AlJjvriTq2NKbaccZSg8k4fyGxxPlOA5JWc9smGytO69sZIdNWjlaLlV2i3qjt1HT1"
                    "ZWFwD4pnsJAcViHP+JUuecJ5XwzIsa+pyXA5PDvLwCIpLtSWGvYbsOAkq2HRWYnFrg2WJjul2tGmZe181gplZd00+jnG17FnQW7SW7OrcksIXIqbJpC3Et2E"
                    "VLqok1snV+3IQ633GojIXh4PI1shUx2UoTx2atuCG3WsxOBjmY5vVHMwlpHT362uBGwAe4BK/mIzeNv4HM5bD3m/p8ng8tfxlxgD2uhvYq7LTtN08Ncz27Na"
                    "Roa9oLSzRG9BSc4RmL/JHHeE5zOnN2F5aUKKvLn0NttuHm+NPPUWTyJDTWkMO3MyCnKY7CUo7avIK10kJQ8kRhy+GZxnkOd4/DA6vSpZJ9jDs6nOZ/A8kxl/"
                    "FxxF3d8dKGf+EveS7/ecdZYCSwlbTzVluXix3IGyPmdmcdFZvOkDfdZl4uutk22Hs2x9ixPE3JOcCCYchEZGRPJao9fVRjGz0k3b6CLTvIPHnuFot7Zk2mlH"
                    "sv8A+9vZl8mW1f52H6VzdXMK7D39rF5Mt77GnyUe2/8AIkdtAdtdwVkHpbCG8toPd2LxcaSNd9YjLb+e32+dEk6GyC7vW3o4iptjGWaTU3GNDp/ak+1ST/2k"
                    "H8ltSyIz34USTL+TKT2urQ6Rvq6vHne+3YbG+3jZ0PwFIzKWBWpygOAdOHRjuQeg7Ejt9+wbsfGi4d9BddyWd+531jJQptyNFUmsiOI+FRoClsmtJkR9zT8w"
                    "5ctJ+PtlF8pHXe4Oc53c6JYBo7AYPzvvshzgPwfjWx7OFqmjiaULmlksw/WWGu8ia30yAOB8OirCvXcB/ihHzoCzJ+mN6amsv595m6oMgo62bU8NYTCwPAbC"
                    "xitSpMTkLk96V+9XWOuqStdXbY7gFBcY9ayNsPvUfKhRGFSI0ycTUAPrl5iauJ4nwatKWvys8vIsmyNw+6pRMtOjBOA/+SS2+aywFh6n1WP23oAfMT6X+ONn"
                    "vZ/lczNilGzCUi5pAbYsCK3cljcW6eW1/Yh21zegSyNc1xcCy7IK3FMpARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARAR"
                    "ARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARVTvUM4jRxP1Rcjxq+omwsbzt2PyrQyHlyp8SQ5yBKs52Ukiy"
                    "kLfS3JVyLX52+ii721UdG5SMRY0aofqUqtk+lzmp5V6YYqpZtCXJ8YfJx603QjcKtCOMYlxjaGh0bMTLSr++N+9YgnfI503ulUIfXX6c/wBhPXXMZerW9jDc"
                    "/qQ8tpFpAack8so8hjbpx085SI33NIj0MowBvRp7+rdJWUrU3yFxtLmt+wtmPyPi8FbSEGmdH/bMUzo0SDIlOSJsD/T9+LESpXZGoLmWltKUvuDKvV7ECOxx"
                    "3k8MB6nGTjWVsMe47hd+qyuC64wdNjgn/tBHJMQ0ulv04XPd/dAR/wCIXRe47l8NLM4uxFhmZowuI9pkV32KGTEO9PNiZ8eLe8d2fp6kkm2Oa73cMeqkz39H"
                    "d0lKEkr/AFD49+4tErS5tggk/jya1t/lX9p7MyIzLoek7iebxbce+JyQAO9bEtE/8wHEdgO/jfjYvpt9vKMc/wCGS2y/euzP4Hmu7RruSS06B8Ndsf4VW6jE"
                    "vHsbs7cjQ3MQ045D7zIzOW84iDV/aZ/egpT7TzhF49vvUZ9qVGcrHEsjcQQT36d+AT9rSAT37kE/Hc/C3bJ05jM0scA51cyMbY6ft/uGNNm938gmCN8TCdff"
                    "oDZICw/HaPsaYb8qIktII9mo1HtKE7PW/B6T5PajPavnXQIcWFrSR9pA6T0nZ8ADQ0dgedEHt+d7GleC58jtAfc5xHgAbJOh4HzoD/wC9HP0PemtHTd6dvDq"
                    "LGhZo825uTN5/wA5NL0l2VYP8hR4CMBk2LUpxR1tnG4ep+OK20pmWYbdZawp7MiKVkqwkSKUvqK5q3nXq1yjJV5/fxuNsMwGKcDtpqYdv6WSVh0Opti623ZD"
                    "9uDmytDHujDCrKvSDjT+Len+BoTsMd21A7K3w4AObaybzaMZAJ17EUkcIB0f7suc1r3OaJcxpBbNQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQE"
                    "QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEUOXrB8TLu+OONeaa9q0dl8d31phe"
                    "QoiLYTTsYtyGmskxL69bNBS5MurzHE8cxTGPZdWiLI5FuEnDdOeciLMD6OeXuxPO8lxSaWNlTlGP/UwNeQHuyeEjszRRRFzgxrZKFnITTjRMn6OBu/tANe/+"
                    "0W4B/aH0jxPOasHuXuA56v8AqnCMucMFyOWvi7m3BpIbDlBhpz1uZEImzucXO6GmBbj7Of8ATXkTEc6feOPU0NoZZOv2DfM8Nt4kuizVCGCJRuyixWztnq5O"
                    "ldlozXvoInWELTZDynBHk3GMvhY2mS3cqtkxg6unWYpTw5DDOe8ac2L+KVKjLJH89R1iNwLXua6mniGUZiuQ0p5XSmpYc/H344GRyyzUL7H1LcUTZXBjpHwT"
                    "PbH1ODmPLZInxStY9mXfVLqiR0n5XBWZPe1n/HqI7rZ97Ty2snJk3WFloltqZ99aVEZpNozWZl4NOkfR6x+p5jQlaCDJicg97D2cwOjqu6XjsWlpIa5pAIf2"
                    "7+DvjiUBxnJ68TzowW8lCX/4XmHAchjLwSCXBzgOneuoOa4AkdIrB8gSjZi1FGgm+xSlWUrW/eSmGlyDXoIiPRtvOOWDiyUXlyEwrRK2ZyzncQGsA7dRefJP"
                    "bs0+NnZL/wBgWg77je9OIVxJPkco4v6mtFODsPbe6yWWrbiT/jiYymxpBaOizK3vsLJHR3092/VZ1RcHdPdTBs5yOT+RKChyI6WaxXWVVgTEn925KyaJNlkT"
                    "Ud/EePK7J8qb+12Q+qmTHhxpc56PFf1f6tcwj4J6ccx5M+Vsc+PwtmKh1OLScteYKGLb9n3kuv2a5b0lu+k7fG3b2bq9PuOu5VzLjuCDSYbWRhku6a15FCp1"
                    "XLu2nbdPr15I9kODevfS7XSfU/rKyupa2vp6iDFrKmpgxKyrrYLDcWDX10BhuLCgw4zKUMxosSM00xHYaQltlltDaEpSkiKit73yPfJI4vkkc573uO3Pe8lz"
                    "nOJ7kucSSfklWjta1jWtaA1rQGtA8BoGgB+wA0F9w4r9QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQE"
                    "QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEWI+e+MG+Z+GeSuL1SkV8rMcStaymtXI6ZiaTJUMnNxW/+jW9HbmnQZJFqrlE"
                    "N59piWuCmO+smXFjJ+Fckn4hy3jvJq7S9+Fy1O9JCHmP9TVjmaLlQvDXFjLdR09V7g1xDJXEAnssM9ReHUPULgnL+EZMA0uU8eymFleR1GB96pJFBaYA5h92"
                    "pYMVqIh7CJYWFrmkAiltNbcSt+PMiTYMlslsy6y2huQ7KBJbI0yK23r5KEvwLKE8S4djXSEJehy2Xoz6EOIUhN6OMuwZClVv05Ypq1uKGxXnhkEkU0ErGyRz"
                    "QyDp645GESRPAHUxwcPJJ/l+zWKyGAy2Qw2TgdWyWIvW8feheOl0NujZmrTsGxsdE0L2jxsaPYd13zq6zf8Ar309aGU/MXMvMazvBuNch20aTK1xWwjt0Ly3"
                    "FH3SXLDB5uKTbCcf3PXljObVtSVKLVPGMCeP+tGWqMibFSyWMv8AI8YGP3urmJonXI+kACIQZyHK14IOzWUa9dwAa5oEhOL5QZeLjOVLpDM+K7QyLpIxG39X"
                    "iuP5WJr4CzqM8bsd+kfPPLqaS3+q93qcPcfWlyOcm1vrSahTamCeKFDW0vubXCgmqOw62r7iUiYpLtgWtkS5atb2ozkBL90jn+R3aBvfYbA8DsD513HcgeAV"
                    "KHC0/wBBiqNZ3UJTELNlr29EjbFoieSJ7SQ4PrhzKncFxbXZ2GgBZj/TIdNjuW8+cx9UF7TW5UnD2ExePcItJMN9nHZ2fclvvuXr9TZJJDU3IcNwegkwLqrN"
                    "x1FfUcq0s2VH92wqpDUAPrj5qK2L4zwKtLqbKWX8hysYcw6o0uqpjGvZ9zx79x9yXrPtj/c2hhk6pWxzD+l/jRmyGf5ZPGPbpxx4eg89fV+psNZZvkOGo3NZ"
                    "AKsYaOt3VK8vbH0xmS7AK3FMxARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARA"
                    "RARARARARARARARARARARARARARARARARARARARARARARVMvUY4pY4n6rOSokSew5Cz19vl6prvZWiZX13Ikq0lW701/20sSzsOS6nklyAbKdx6tmDCkEp2O"
                    "qTItz+lLlx5T6U4qnL3t8Zmk47Yd7jXOeyk2OShqPrMkcUeJnx8Ic8NbJLFP7W+lwbQF9enAHcK9f8zk69aSLFc9xtDl1WdsfTXkyUpmx2drsIaxhnju0Y79"
                    "prCSG5SCV7vcmcVFN1NZhY490y8k0saW8TNxlXGtpBimZnDj3lXdyErs0J7le1LegLQw6aS1KZgRkOHuKylUhJMcx3J8Jl2xQixVoZilJO7Xuuq2f0hiiPSP"
                    "7yOKwJJWB51E6aRze8jgtP8AojNHf5FUwdkyOrCe5kDGNECH+CZWrcY0l32e+ZabXdDQT0N32YCISWo3stssITv2222mU7UajJJJQ2R+DUZnrtL8n8kZnozy"
                    "d7i1gOtjqGwPBOuwG9a2SCd/0IGipvPl92R8rzrre57yQGjbi5ziddgB50NADtsfHo4+iD09xOn705ODkqYUjIea4cvqDyeUa1ds9XJzUKVg8hEZTrv0Bs8T"
                    "V3HsCVEI21HYw5suRHiy5cmO3Sp9SHLxzP1i5hfilMtHF3W8cxzuouH6XBMFCR7fgNnux27LQ0AanG+t3VI+y/0b487jXpzxunLH7Vu5U/jF1hA6mWcq43DG"
                    "8ju50EUsVfbi46iABDA1rZaxoxbQQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQE"
                    "QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEUNfrK8YSLviXjbl2HJZZTxzlNhi9808hLTS6PkZmvZg2KphJUapsPK8coMfqoT5tx315lOU2"
                    "+iWTEeZMn6LeWtxPqDluKzb9nlmKElchw+3IYMzWmMDCP5ZMfYyEj3NId1VoQWuaNsry/wBo56ejkfpHh+dVoTJe9PM250xjjc+T+CclFfH3iekj+7gyNbDW"
                    "JC5rhHFHNICwNcH1MOsZa08B2/Yk/wD7ywtLii2fY2p20PuMz8kSnEtp8dxH3f8Aq8laKGgXq/fxBaP4/wDe1OxH8pA8eB30PO1VD9PwJ9R6vUT2wuYc0HQ3"
                    "/uw8aGxoFzu+tEAeOx0c6ReAbHql6muEOnmqmzq57lvkGlxefbVsP62xpMaI3LvOsmrYykm1ImYngNVlOVsNSFNxDOjP6p5iJ7rzeE+r/NGcA9OeV8qbIyO3"
                    "jsdIzFgua33MtbLamNjaXtLS5tmZs50122ROPQ92o3WNenfGTzDmnHsC6IyVbF+KXJAhzmtxdTqsW/da0tPRJHGK5II2bDQXMBLm+p5AgxKyDDrYDDcWDXxI"
                    "8GFFaLtajRIjKI8ZhtOz02yy2htBbPSUkWxRS975Hvke4ufI5z3uPcue4lznE/JJJJ/cq09rWsa1jQA1jQ1oHYBrRoAD4AAAC+ocVyQEQEQEQEQEQEQEQEQE"
                    "QEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQE"
                    "QEWHOoPiiJzjwnybxRJRVe/mmJWddSS7uGU+sp8sYbKxwzI34pturWvGMthUuRRXGEfVR5dYxJiLblNMuJyjhPJZ+Hcu45yiv7pfgsxRyD44ZDFJYrQTsNyo"
                    "HgjpbcqGeq/Z0WTOa7bSQsO9QuH0/UDgvLuE5DpFXlPHsrhHyPHUIH36csEFpo04iSpYdFZicAXMkiY9o6mhUBetBl+HwXfRJMV6FJbzPDUSYL623H4K0PWi"
                    "lRXVMKW0p6MvbD6krUgnErQk9mZqvbxdyHJR4/I15Gy1rmNNuCZvYSxWHUJI5B3I6XscHtIJ6m6IJAX86voti7eD9XpcJkmOhyWGh5TichCQR7eRxpdQvxHu"
                    "Wn2rMMrB4GgSA4aKkJ/TLdOiMp5+5n6lrSYj6Dh3j2uwDHKV6sRIKZmHLk6TJk5LDuFPIcrJeH4jg11QPVyIj52cHlA31TIya4mZsIvrl5nJWwvEeC13kDK3"
                    "LXIck1jwGmvjQ2pj45W9O3+5Zs2pm/e1rTWDnxuLonsuH+ljjrLOS5LyyeBu6EUOEx8rwesSWt2sg6L7tAGKKmx7y0ucXlkb2gTNfdUFbKmogIgIgIgIgIgI"
                    "gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgI"
                    "gIgIgIgIgIgIqP3rq8KO8Tclczxa+HUV+OZ/k2C8z40zW/VKWqJyJc2MLKV3ByEIYau5nK8DkCxOPCU9Ej4+/jyjUzJlrjtW8fSXzP8Atb6XYmlYlklyHEXW"
                    "+OWjM5pc6vVdBbxTmAO6hBHjLNWoxzwwulq2AC/2y5U2fUJwFvAfq8hy9au5mJ9SeOZbktMRgiJmVgxzq3IIY9N0HvsY9uSlhBc4SZX3dticImWGvRJ6dP8A"
                    "p19O/hZme1ARlHNUZ3qCyp2vaeZbeXyZX1T+FtSkyWI8n91ruLKvAam7JwnWkXcCxRCfdgFFWde/1F8yPN/V7l2RjeX0cZdPHcZ46RTwZdRMjAC4dNm0yzb3"
                    "vbjOXODS4sbaF6NcZ/sr6d8eoSNIt3K7sze6jtwtZZ36wxuPS3/5nhkhrNGvtbCBt5Be6WcaPW0UBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE"
                    "BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBFEJ6rPQXP61S6aGKSiK2R"
                    "T8wY1iHKDjcyNWlX8J5JlmJ5PnWSTnDXGl27uMQMFkwsbqGpRmi3zB55ppsnpMuPIr0F9YW+llT1IgmndGc7xKZ+EZ0SSsfyig58OKic1vUyKOWLIWpJpnRu"
                    "B/SQxvIYSDHn1x9IpfUnOek+YqMj93iXL7bczJqMT/2Tz2FvUsu2vI7pc2aO7DipGmORkgh/UtiPVJ0ul0ixY0GNHhQo7EOHDYZixIkVluPGixo7aWmI8dhp"
                    "KGmGGGkIaZZaQltttKUISlKSIo7ve+R7pJHOe97nPe97i573uJc5znOJLnOJJc4kkkkk7KkI1rWNaxjWsY1oa1rQGta1o01rWjQDQAAAAAANDsv3HFckBEBE"
                    "BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE"
                    "BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE"
                    "BEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE"
                    "BEBEBEBEBEBEBEBEBEBEBEBEBEBEBFppyz6hPRnwXnl/xjyzz1iuF57i6axV/jVjX5PJnVpXFNX5DWk8usoZ0VapVLa11ihDMh1aI8xhTqUKWSRnGF9NPUDk"
                    "eOhy+A4fyDMYyw6ZkN7HY6xbrSPryOimY2SFrm9UUjHMe06LXDRHcbwrMepHAOPZGXEZ7mXG8NlIWwvlx+TzFKjbjZYY2SB7obE0bwyZjmujfrpeD9pOjrHC"
                    "PVk9PB3ftdT2HO68mTVFnbhpL+TSjE1KIv8AOteDMz0Rj12+ivq28bZ6c8wcO523B3j2HcntF4A+V5T/AFl9JY/+09SuDs32HVyfEN2T4A3bGyfgDypAae3r"
                    "Mgqau+pZ0ezpruug29TZQ3CeiWFZZRmpkCdFdT9rseXFeakMOF4W04lReDGs5I3wySRSscySJ7o5GOGnMexxa9jh8Oa4EEfBC2PG9krGSRuD45GNex7TtrmP"
                    "Ac1zT8hzSCD8grkRwXNARARDMiIzM9EXkzP4Iv5MEUdr3qy+nYw442vqlwVZtqWk3GKzM5MZZIUpJuMS4+MORZDKzSamX2HnGX2zS6w442tClbMh9GfVexDF"
                    "Yg9PuVzQTRMmiljxFp8ckUjQ+ORj2sILHsc17XA6LSHDsQVrqb1e9LK089Wx6hcOgs1ppK1ivNyDGRTQWInFksMsUlhr45Y3gsexwDmua5pAcCF3Hjn1Jeh7"
                    "lvOcX40466hsRyjOs1sVVGK45DrsqjzbuzTBmWa4UJc/H4kVTyK+BNmLJchBIjxXnFGRIMdDMelvqJx7HT5fOcNz2JxlZodPeyFCWrXjDpGxN3JKGgudI9rA"
                    "0bcS4aC7uI9S/T/kGQixWC5lx3M5Kfq9qji8rUv2XBjDI93tVpZXtYxjS573BrGgHbt9lvCMCWcICICICICLVbm7rd6U+m/MIWA8381Ytx5mNjjUHMYdBcM3"
                    "b857GbO0u6avuNVlVPZbizLTHL2HH911DrrtVN7WzQ0ahl3H+A815XUmv8a4tnM7SrWDUsWsXjrFyGGyIo5jBK+FjwyT2pY5Ol2iWvBCxXPc64Xxa1BS5Lyv"
                    "j+At2oDarVsxlqOOmsV2yOidNBHbmidLG2Rroy5gcA8dJ0SN4XL1ZfTo9tby+q3jhlhttbrsiU3ksSKy00k1uuvSpVCzHZZabSpx1511DTTaVLcWlKVGXuP9"
                    "HPVWNjpZPT/lbImMdI+R+HttjYxjS973vMYa1jGNLnucQGtBc4gAleNH6t+l80rIIfUDiEs8skcUUEXIMY+aWWVwZHFHE2yXvkkeQxkbWlz3kMaC7spERrVb"
                    "D8oCICLSfkv1G+iTh7Ncg465I6hsLxjNMVnIrMion4+Qzn6mxXGZl/QypFVSzoRSm2X2/qGG5K3Ir3uRpKWpLTrKM9xHpb6jZ7HVsthuFcjyeMuNc+pep4u1"
                    "NWsNa98bnQysjLXtEjHM20kdTSPgrBst6m+nmCyFjE5nm3F8Xk6ha2zQvZvH1rlcvYyVomrS2GzRkxyMkHWwbY9pG9jfHYZ6l/QtyHlONYThXUViGRZVmN/T"
                    "4vjVLArct+rtb6/so1PUVzRv46yyy5NspkWKh2S6xHbW+hbzrTRmsvtkvSb1Lw9C5lcrwjkePxuPhdYu3beNngrVYGEB0s8sgDY2guaAXEbLgBskL44/1V9N"
                    "stkKmKxfOuLZHJ35m16VClmqFm3bnc1z2xV68U7pJXlrXOLWNJAa4nQa4jeka8WfoCICICICICICICICICICICICICLQ7LfU76DcEyrJcIy7qSwuhynD7+5x"
                    "bJKebX5YT9Xf4/ZSKi5rXXmsdciuu19lFkRH3Iz70c3Wle28tOlHsSl6Sep2SpVMjj+B8ouUL9aG3SuV8Rblr2qtiJk0E9eVsZZLHLE9kjHMJBa9p+QsBueq"
                    "nprj7tvHXuecSqX6NiardpWM9jYbVSzBI+GaCzA+w2SCWKWN8ckcjWua5jgQCCuy8S+oV0Zc655RcYcTc+Ylmme5MVodBjVdDySPOtf2aosb+zTFcsqSFEUu"
                    "JS1FnZONrkIcVFgyHEIX7ZkOhnfTjnvGMe7K8h4jnsNjWyxwOu5HHWKtYTzEiKH3ZWNYZJOlxYwEucGuIGmuI72E9QeC8lvfwvj3L+OZvJexJZ/Q4vMUb1sV"
                    "4SwSzmCtPLIIY3SRtfKWhjXSMaXbe0Hcwz15GFLMPCjnY9W705ZUeNLi9VOBy4syOzLiSYtbmUlmTGktk7HfYcYxlxLrUhpSHmFoM0vMuNvNmppxC1bMi9Gv"
                    "VaeNksPp9yuWORrXxyR4e49kjHta9j2PbGWOY5jmua4OILSCCQQtdz+rnpdVmlr2fULh1eeCR8M8M/IsXDNBLE90csU0UlpskUsUjXMkje1r2Pa5jmhwIH7o"
                    "9WP07nCM0dUeDr1+EVWaLM/7S8EnGDM9moiLW9mei8+B9Geinq28gN9O+Wd/Bdh7UYOzoDqkY0bP43vXfWl8JPWb0miG5PUfhgH7chxjz42ezLLj2/OtfAO1"
                    "uRxHy/xxzxgFHynxLk8fM+P8lOyKgyaHCtYMG2TUWk2lsHYbdxBrpjrEe0rpsP6goxMPORnDYcdbIlnr7J4y/hr9rF5SrLSyFKV8FupOAJq88ZLXxTNBPRIw"
                    "ghzDpwPkLPcbkqOXpVsljLUV2hciZPVtwO64LMMjQ6OaCQfbLE9pBZIwuY4fyuKySOgu8gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIqnX6g3pyexzlLifqcoaOQqi5LqJXGfIFumzhlAgcg4bAet8JIqV0ytZFrnGAIydqZYwlu1MGs4br40tiDPtGXrewj6Jud9bOR+nN2YbYW"
                    "8kwbXHpLGF0dXNV2aGn9ZdRsgHTmdExb1tcfarw+uTgTQzjXqVSYWOY4cWz7mMLhLHK6S1grUsm9RCrN+vpNHS4SvyMXeN0QMldhh5TKkvN/7bqFpUhZH2ml"
                    "aVmtKk/B7SskqToi0f8A7iwlg6HNc3QIPUCCO5Hgnvo9/k72N6Vc0sfuBzHkua5rmvafBD2kOad/BaSCPkf0V5D0Zef4vNvQ9gWMyrh21y/p8ku8G5KmRFKI"
                    "8xT41Ar7Xi4mlHpy2jscS3eE08vIVpUq1yWmyREl+RZQrFwUw/UZwl3CPVjktSGuYcXmLJ5BiSA8MdVyhM80YLy77q902YHhri0dDXNEbXNjZdr9OXOf7fek"
                    "fFcnPKJMpjan8AzA643PF7DH9EJnNjDegW60cFuNr2Nc1swaTIW+6+VwaMW80BEBFH16onUdL6Y+izlzNMevZWOchZfAY4o4vtq9tty0qs45CJ6oayWq+oad"
                    "gIn4HjhZHyIydilUNxOIuMKZmPPMQJW0/RXhH/SF6mcW41KzroTX23csdBzW4rHNNu6HNPkTMiFYAbJdO3wNkau9Z+cj059M+V8rYQbtLGvr4qMlzfdy19wp"
                    "46MOZ9zT+pmY/qBBAYdHelQYNDDKW2GG0sx2m0ssstGZJZZbSSG2W0kZ6Q02hJJT+PGzL5F3zGNiYI42MZGwBkTGgBrGtADGt7ABoaA0a/HbelRQ580z3zTS"
                    "OlmlkdJJK/RdLJI5z3SPOgS55cST277761qf/wBADptj53zrn3UpkVNAm0PBFKWL4DLmxlPuxOXeRKyxr7i2ppKXyRAt8R4nk3VLaNvx1rl0vNcJyMtJIeEC"
                    "fre5s2vjOLcAqTtdLfmk5Hl2NcHvZWrF9TGwS7/7P3rDrNhobt7hAzqcxjumWwn6F+DOkucs9RLkUobBFHxfDPftjJDMYL+Wnj+0mdjfbo1g4vbHFLHM325J"
                    "GtfBbtFcisfQEQEQEQEVNL9QSok9deHK3o/+kziHz5I//wCcXU2ZHv8AB+D0Wy35/G9Wd/Q7v/o75ho6/wDXRoI2AdHB4/v+fg/tvW/gGr/67m755wb9+I3R"
                    "3J1oZmT41o/zHfz3G+2lAPyCtJ4Bne+0jThOXH3LIv8AbM8esTJW/JkRGW1bURfb8JMtnLblDW/2cz2x/wDQTLho332cfY3o71vsP9d9wVEXhDS3mnEAN9+V"
                    "ce23ROy3MUuxHzretDuQSP6+pKKDlf0gItVutbqco+kHpq5M50tSrZdrjtU1VYFQ2rkxuJlXJeTSWqLAsbk/tzb1imusMjnQXMgmQmVrpsZjXV9INiDVypLO"
                    "Z+nvDb3qBzLAcSodTZMxfignnGtVKTT7l224kOaBXrNke3qBDpAxga5zw04hz7mFDgPD8/y7JfdXwmOnttiGi+zYa3pq1WN6mdT7Fh0cQaHtP3E9Q1seeTeX"
                    "1tkl3d5HkNrMvsjyS7t8jyO/sXFP2WQ5JkFnLucgyCwddNanbG8uZ023sXSM/dmzH3O4iV4vJw+Jx/H8TjMHioW1cdiaNbH1IWgNDIKkLYYwQ3TQXhhe7Wx1"
                    "knQ3pUQZ/M5Tk+ayvIMxM+fJZnIWslce5znETWpXSFgcSXGKBrmwxF7nO9qONrnFw2rFvoH9GzOVZXedamb1JuUeDSLvA+DimRorkW1ziQxMo+SM/gLdaVIT"
                    "/RtU/L40p7CG+3HfuL7kyqsIyp2MwJKa/wD6y/VMTWanpfhrQdFWMGT5TJBI0h1kN6sfi5OgHf6drv1ViMvaWzGJr4iGscrEPos9Jn0KFv1RzVYtsZKOXG8W"
                    "jnY8Phxwe39bk2B5a0G/KxscD/bk3WibJFM0TSsNrIQDU/kBEBEBEBEBEBEBEBEBEBEBEBEBF5zfWMsz6veqsjMtF1K88oQRfaokp5WyxJ78EpRn2no1bP8A"
                    "BbItleR6PNb/ANFHpu0Hpb/YbijvwNuwtPq2AR1bdsedgDx3Con9ZwT6t+pmndhz3lewCSATm7pAPbt2Dd/G96+d7Xejc4lz1Ium5P5TJ5ZWlPktH/0/8upM"
                    "/PyRpUZFr7i/PjZjTf1jAn0YmJ2Q3lPHyCfgubfB+dgH9x5+Bra3X9E7APWWUkHZ4XnwDvvsXsH3d23oAkfA24eeyvbCptW2+V5bXF6t8X8Y95dxnxvgiSUr"
                    "wo0JxapSRK1o/JJ15LZFrzo/F+HEu3FeM78jj2EGjog6xVXetefz87O+xB2qBfUMf+0DnffQbzTlTdAeD/Hch28dyNnz38Dw3SlS9LHpU496xerem4x5SfmO"
                    "YBiWCZPzBkuPwlvMLzmFhWSYHjsbBZc5l1l6rpLm0zyBMyKTH92XOoKeyx6J+3v3rV1VaM+qP1Eznp56cwWuOuFfJ5/LMwjMiHASY6GWlZsz2IGFruuw+OAw"
                    "xEloic8yn3OjoO8PpP8ATjB+oXqRZZyNjrON45iP43/DHMaYcjaNyGvWjsuPdsED3/qHMYNyua0dTR1E3waamqMcp6rHsfq66joaKtg01JSVEKNW1NPUVkVq"
                    "FW1dXXQ22YkCur4bDMSFCistRosZlphhtDSEpKoaaaaxNLYsSyTzzySTTzSvdJLNNK4vkllkeS98kj3Oe97iXOcS5xJJKuGiiigijggjZDDDGyKGKJrWRxRR"
                    "tDI442NAaxjGANY1oDWtAAAAC5IfNfRARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARaLepJ04SeqXo15l40oaSNe8"
                    "hV9D/X/E8R5cKLJe5KwFz+o8dpoFvPNLFCvNm4c/jq0uPejlHxzMLth99EOTJJWxfSbmj/T/ANQ+L8p2f01DJRMyLNyBsuLt7q5BjxFuRzRWmfIGND+qSNm4"
                    "5AOh2vfVXhUPqH6fco4jKQyTK4yQUpeljzXyVZzbeOsNbLqMuiuQQuHWWjsdPjP3t8+lEmO+2xIiuIfjSGm5EZ9JKIn47zaXYz3b9qy91pxDnatKVJ3pSS0Z"
                    "C8eGaKxFFNXkZNDPEyaCaM7ZLDNG2SKWNxGnMkjcHMOtODtdie1FFmpYqz2KtuGStbqzS1rVeQBslexBI+KzBIDsNfFKx8b9EjbexPYqdv0Geotrjbqst+FL"
                    "65nMY/1E4a/SUVaUd2ZXHydxsxd5zjLpvdpt0LM7B3OV41hYuOtt3Nu3hVG4l6adS23Cf61+FfxLiOA5vWjDrPHMgcVfcWkyHFZXrkhcHRgAiC/EA4TNLWts"
                    "AxyREPZPN76G+a/wzlXJuBWXFtbkFEZ/HDraG/xTFmKrdY9shBL7GPkruiEGnapTGVkjOmSvckFZas7QEQEVSz9Qj1CLybmDiXpspLe1Kn4oxiRydm9Y0hn+"
                    "nrHPeQifpMOTKdUpUk8kwTBqvIZCGUNsxE0vMiFG9MfcU1W2I/Q/wgMh5T6hW4gHymPjeGe5gJdEOm3lpI36PTuRtKBzOoF2iXN6egurs+ufnDj/AGT9Oqjw"
                    "WPdLyXNASOaW+1urh4nxbDXske+5YbKQQySu0Md1e4BXaVKZZacfkPtxY8dtb8iQ8oktR4zKDcfkvKPSUtMMIceecMyJDSHFq0kjUU/pXsZHLK/pZHGx8kjn"
                    "E9DGRtL3vJAJDI2gvc470B5A0q8o68k0scMDHzSzSMihiaAXzSyvEcUUY2NySSObGxv+J7mtGydK/wCemN04v9MPRfxDg13AbgZ5lFW5ynyWhdG3RWjGa8hk"
                    "1eOY/fMEZyZ1tx7ji8d4vVaz1FNnQMIgLXHgNJZrYdIHrLzd/qF6j8m5IHvdSluuo4ljpDI2LFY4fpaYi/wxxztjdbMbBoSWZCS97nyPvS9G+Cw+nPpvxfir"
                    "GsFqpQZZykjGdBny18m3kJX725zhPK6IOcTpkTGNDY2sY3fwavWzkBEBEBEBFXu9VP0r+pzrX6mce5h4bv8AgitxGt4KwPjWZF5Nz3kHFMjTkeK5/wAy5TOk"
                    "RoGK8Nci1kukkVvIdI1ElPXUKcU2LatPViGERJUuXP09fURxz0d4xnMFmcBmsvPlc8MtFNjJKLIooRj6tP2n/qp43+458Ln7Y3Qb0/cSdNiX9Q307Z71n5Fx"
                    "7M4nkeIwsOGwtnGTQ5GnctSTyT3v1TZI3VZGNYxjPtIf1Eu0QAB3i0yv9PB6gN7i2SUMbMejxp+8oLqmZkv8zc0m3HXaVsqC28ptvpZJRk2uSThpS4jaUm2S"
                    "vv707ry31scIyOMyFCPh3KonXqNyn1vs4pzWGzXkga89Nku033Oo6aT27A/OmOP/AEScqw2dw2Xm51x+dmKyuOyXsRYrIsfMKNyC06IOfYLQZBD0DqHSOrZ7"
                    "DSuuF8F//wA+P/YVtqxtARV+vWJ6Y+vjrHzzjLjHgri2os+n3jukXnMnJLDlTCsaRlXL98q4x5xmwxu1mtXkZrj/AAsn4tDZJaXBtHOSMpaeaNdVBcOUv03+"
                    "o/pr6V2MzyTlQyNnkduL+G4yOlRdOKGOcA+1IJXBjBNdkDI3dMj+mGL7mtJBMYfqR9NvUb1XxuI4xxSfGUsBBZ/iOYdftOhfftwjVGFrI/c6oarnPnIkjj/v"
                    "2xOY946gyG3EfRF9RW5zHEqPJOK8UwvFrvIqyuyjOHeWuPLlnCMdemNndZIqjrLGwtbyXX1yZTtRT1tZP/dLg66FYKr6uROtK+Uma+s303jxWQkwlXM2cyyp"
                    "OcbDboGOtJb9pwgE8nUOmP3C3q2CCNtdppKizhPop54/L40Z/I4aPC/rIBlH0bTn2zSEgM36drw373MHSTtrmBzpI+p7WNddS4i4qwvg7jDA+IOO6z9ownjn"
                    "F6jE8dhrWl6UcCoitxvrrOWltpVld2ryXbS9t30fV3FzMnWk1bkuY+4usnM5e/n8tkc3lJ32cjlbli9cne57i+ezI6V+jI57wxpd0RtLndEbWsB00KzPEYql"
                    "g8Xj8PjYGVqGMp16NSCNkcbI4K0TYowGRMjjBIb1O6GNBcXO6RtZFHmL0UBEBEBEBEBEBEBEBEBEBEBEBEBFUf6gfQv63uUee+buTcYyjpWYxrkbmTlLP8fb"
                    "vuWeXK69aoM0zu9yWlYuayv6bryBBto9bZx27OJBvbeG1OS81EspUdCJDth/BvrG4TxbhvFONXOJ8ptWuP8AHMLhp7NaTEtgmsYzHV6U00IluCT2nvhLog9r"
                    "H9BHWAT0tr1519GfKuW8z5XyipzPj1KvyLkWYzUNSxj8lNNXiyN+e3HFLJE5jHyNE3S/paGgjsX/AMxzt6evo7dW/Sz1e8Tc8cpZN06TMIwQ86Xbw8A5F5Ly"
                    "PKpK8l4wzjCq5utqsi4MweocS3Z5PDkzXZGRxjZgMSlsMSpBNMLwP11+privqrwKXieI47yDGXH5jG5IWclJj31RDR/Vdceq1mWUSv8Afb0ksLdNIcQekrPP"
                    "Qf6Y+R+kXOncryfJ8LmKjsHksT+koU71ewJL09CZkpksExujZ+jcHN0Hbe0g6BCs1CFimcvLS4zXrjLjNBHsk8dYMktEZGtJYvVaUZq2W1F9xEeyIvjt+Dvx"
                    "4n/9a/Gvj/1fwxO99icbVJHjZ7k9/wDP+lBvqC3/ANoHPT/+uvK3f/52/oD9wPAGj2Ot/M936fp0z68cmTrwrpZ5ZT3fn7eUOnsyI9mZkfx43rx5LYif9bo3"
                    "6b8ZPbf9sYif6/wjI6/5Ej86A791K36G+3qJyr9+IsHx85Wu4/8AMf8AP+iufir5WioCICICICICICICICICICICICICICICICICICICICICICICICICICIC"
                    "ICICICICICICICICKgT6rPTtE6Zet7lzEqmPXQMSz+UzzhgFZWRVwma/EuULG5lzq4oiWWq2BEpeSKjkTG6OrqVKh12LUePJ9mGclENq376WudjmfpXjK9mw"
                    "6XLcUeeO5D3XF8z4oG+9i7T5C4vl92i9kRe4B/uVntPUR7stQv1W8APDvVS/ka1YQ4fmMIz1J8emwMvbZXzFWOJo/uzFZbFbcCC2Q5DcZa1rootLOLOS8q4i"
                    "5EwblLCZcuJlfHuW43m9B9HKVE+stcTvYGQQqea4ky92lvJVc3TZBCcNcaxoLCyrZjT8SVIZc3Lzri9fmvEORcVtAGLOYy1SaXNBDbBb105QPIMVtsEge0sc"
                    "Cz7XsJBGkuA8pm4JzbjPLoP5sFlq1uXv0k039VbIx+COqTHT2o2h7Xs6ntJY7s0+khxhyNivMHG2AcsYLNcssK5MwvGM+xKwejuw35mN5dSwr+lkSIb5JfhS"
                    "Xa6fHXIhyEokRHjXHfQh1taSopyWPtYnI38XdZ7dzG3LNG1H3IZYqTPgmaCQ0kNkY4AkDYG9BXu0bkGRpVL9V3XWu1oLcDttPVDYjbLGSWOewkscN9LnN34c"
                    "R3XeR0l2lx1xb1uP1FpfXMxmup6Sum29tYSVGmPAra2M7MnzX1ERmlmLFZdfdURGZNtqMiM/A5xRvmkjhiaXySvZHGwa2573BrGjehtziANkDuuEj2RMfJI4"
                    "NZGxz3uO9NYwFznHWzoAE9gSvN36kucbDqM575d50s/3xt3lPPb3L4EDI0x27ykxqW8iHg2LWrUV6RFZm4VgUDF8KW2xIlNIbx5ttMqShCX3LzfSvh8fAvTz"
                    "inFWxtZYxuIrSZDpb0h+WuM/V5R/SGg9ZuTStJLQ49A6z1Ek0Z+rnLn+oHqTy3lIlEtS7lZ62Kd1+6Bh8c40sb7T9u3FLBD+raGu6GusO9sBmgtgfTS6d2+q"
                    "TrR4X4ztY0+Rh1Tcu8qckfSwI9hF/oLi9yFkE+ruWZaVRjxzN8pXhvFl+s+x5iByGp2Gf1aYvdgX1Lc9PA/SrOTVZfay3Im/2cxOpPbka/IMc2/Yice5fVxw"
                    "nkAYHOL3MBDGF0sef/TF6ft5z6r4b9XAybEcXB5JlGyRGWNz6p1ioJG9g0z5H25Y3v0z/dZD/eOa2J91LrJ66eAuhvCK7LOZbmzl3WSuS2MF41w+LCt+Q86e"
                    "rXYDdw/R1U+xqKyFSUCbSvcyDKckuKLGKl2fU1T9seQX+OU9xVZ6d+mXL/VHMnC8Tx36l8LGS379h/6fGYuB7ixk9+2WuETZHgsijYySeZwd7cTmskcy1b1A"
                    "9R+JemeFOb5ZkmUoJHuhpVIx72QydlrDIauPqNIfYmEYMj9dMcUYL5ZGN7qsLzf+oD6wM9nqj8P45xvwBjTcpTjRwq5HKeeONtuyW2EzMrzavi4acWXGcjPS"
                    "q6LxRHmQ50c0xsnmw1qS5PLif0T8NoRRScx5Blc/cLR7tbFOZicexxEbtRyOjsXJSD7jXOMkYc1zeljXN6zA/ln1s8ruzSR8K41jMNUH/Y2s4JMpek2Hh3uV"
                    "q08FWFo2wxls0x20mRune23VaR6xnqTyVpWvqovkmlwnEE3xb09R9Ggz7CJMPh6I0tsjL7232n219ppdQ8Rn3bPi+lb0NhYWjiE8hLC0mfkPIJH7I0XbbkY2"
                    "B4/mGmgB2ukADQ1VN9VfrpNIJP7U0YQHNIjr8dwzIwWnfSRJUme5pBLdGT7gd7347ph3re+ozjFnDsLPmmg5CYiyCdcpM+4l4sRQWDPYptUWcnjbFeNMlU2X"
                    "cT6VV2V1cn3UIJchTJrac8TMfSB6NZGuY8fRzWAnIcBaoZq7b6Se4c6vln5Bj+nXYNdE3Z7/AAF7uH+sH1lxs4kyU/H8/WBBdXuYWKjK4N39rLWMfWEYeSA9"
                    "z4Jz2HR0klT59AfrV8X9U+VU3DXNGOQuFeZsgmKrcOsGLUpfF/JFgs4TNfSVFjaONW+H51dS35bNThN5+7wLU48KFj+c3+SWzWNxobesv0ycn9MKk3IcVb/t"
                    "RxKJzRYvQ13RZHFB5eA7J1Ge5H+mGmA3oZPaD3kSRRNY56mV6M/Uvxf1SmhwWSrDi/L3sLo8XPYbPRyRY0OeMRecInTyD7z+kmijshjC8NeCCZwhGJSYVaf1"
                    "hPUh6r+kjqpxfing3PKjFMOsOn3j/kGZEmYPh+SyZGS5HyPzXjdnJOfkdPZSW2f2vBsfaYhsOIjtONyXjbJclaxNL6avQTgnqxxDPZvlUmcZdxvI3Yqt/C78"
                    "NWE1f4XRtj3WSVbBdKJp5NPBaOlwaWuIChv9SXr1zj0k5Rx7EcXrcfsU8tgJsjZOXp27M7LUWRlqj2XVr9RoiMTW9TXteeoEtd30oiMq9c31IabF8ovIfNeN"
                    "qlUuNX1tFaf4g4sXGXLramXOYS8hGLtvrbJ6MhLjaHG1OoU4lK0KNC297Zf6PfSbHYrKXoZuUyT0sbeuQsmy9f2XS1qs00bZRHQZIY3SRhrwx0bunq04O0To"
                    "7AfWF6rZfP4LE2cfxCCvlMzisbYlgxeRE8cF6/XrSyQmXLyxCVsczizrje1rw1zmPbtrr54q3VnKr09d/ruYdwhllvxJ0q4vjHMeb49LtafL+Tcqn2K+J8Uv"
                    "oPvRFVWPVOOya+45Vn1Vuw7CydMLKcEx+tW0cOqyu+tG7WBTS29HvpS5B6gY+tyPlGQk4px20WPpRCt7ubydd3udU8FeboiqQO6Y/Zlsh3vtkdJGwiLUkUfW"
                    "D6qeNend+zx3j1Acr5LV6mW2MstgxGNsD2i2vcuRiWWSctfIXw1YpHwujaybo91rmwhZV6z3qN5S5YLT1ELxaJZHIJ2nxDjDh6rqYKZLrjn09LNtcDvc1gMR"
                    "EqTHiOSs1sbNLCEqkWMmUS5a5eYn6S/RbGRwfqcJkszYZovnyWbyPTKQGgn9JRkqVxGTshjmyaJ0XkaUP839W/rTlZJW0sph8HVeS1kGMwtZ87YtnRfdv/rZ"
                    "HzaPSZIIqsZ0HNiYSQejYx6s/qJYkSiquqvPJCFrM3P6ix/jHNXNKcJa0IczvBMnW0RkZpL2VNLQjSEG2lKdezd+mH0PvsLJOGMrE+JKGVy1ORutj7fZvCNx"
                    "HkiSN43okHZC8aj9UHrjRkDm8wZa0CTHfxGKtxP77Iduqx48kbZJG4eA7WlIT01fqDuecNsq6n6mcKx3mnD3ZyUWWV4fXwOPuU6evcKS7ImNQIzjPG+bSWHP"
                    "pIlfRlVcWtFHORKsMtkvoQ07pDnf0UYaxXmt+neftY+4xpMWI5FILVKy5rCQyPJwQssVpHuGtzQTxAOBL2dP37z4J9a2VitR1PUbjtazSkc0HM8aY+GzX6ng"
                    "OdNiLM8rZYWR9T917ZsFzehsEnWDHao4O5z4t6j+M8c5d4cyyDmODZOw4uFZREux5UKbGX7NlR3tVLQzY0ORU0slwrmjtY0WxrZbamZLCNoUqv8A5JxvN8Rz"
                    "V/j/ACHHz43LY6Z0NqrO3RBB+2SN42yaCVunxTRudHIwhzXFT649yHDcqw9HPYC/BksVkYWz1bdd4ex7XAEtcB3jljJ6ZIngPjeC1wBWWx4a9pQldevrWcNd"
                    "K19ccUcO0sDnvmmhnuVeWtRr5NXxlxvYRZX0lpT5NlNfGtJd9m9W8TrUvBcaiKTVzYsupzTKsItUxIk2SnpB9M/L/U+CLN3phxbikg6ocrcrvmuZJvUBvFY8"
                    "vhdPER1EWpZIq5IZ0OeJGuUcvV76k+HelssmHrxu5PysB3VhaE8ccNE9DnMdlb5bKyoC8MaYI457fS57m13mF7FX5zX1uPUYzKfYzIPNFDx1EnOuk3Qcb8U8"
                    "cMUcOK60THs10jkGg5GzSIaSI3ifezOdNbkurW1OQ2llpqaOB+kH0dxUMYydPM8iss6dz5DLWasUjg4kPNXGGnGC/Ya5jnzRdIA1vq3C3kH1ger+VkeMVLgu"
                    "OVX9WoaeJju2Yw4aEf6zJOtCRrSPtlZWrygk7Ow3WJ8f9WD1DcalLl1nVZyA4+6tSnTuqXjrLGNr+0ybgZhhF/WsJ7S+xLEJCGu41NpQZmQyq19M3oldiEUv"
                    "CYIddvcq5PL1Zdd9EvgvRnYI1sDTu4dsDZxOp9TfrjTl91nMnT99+1bxeJtREHyOiSl1Eduw6u3fpI2St9enT9QZ1J4LZ1tR1D4hiPO+IvW0FFnkVTDg8Z8o"
                    "1dM+6abedGXRRC46yubBZJLtFjbmIcfMz3yXEt85gsvosYWk+b/RVx65FNZ4Dn7WHuAPMONzkhv46Q9Lnxs/XRRC5WLn9MQfIyy1senkOex/ubu4R9aueqyw"
                    "1vUHjtXKVCGiXK8dYad6M9TGlzsZYmdVlaI+uZ/tWonukaWQxdMjWw2lOmDqp4W6vuMIHK3COUJvaRx1muyCmmtIg5Xg2SqroNpLxHM6VL0j9qvoEWxiPKOP"
                    "JnU9pDkRbnHba6oZ9day4C8x4ZyPgectcd5Rjpcbk6p30OIfDYgcT7VqpOzcdirO0dcUrD3adODXAtE9+Jcw47znCVeQ8XyUOTxdrYbNGHMkimZ2lrWYJGtl"
                    "r2YHbjmgla18bwWkdlsUMWWTKIr1FPVt4t6IJyeMcWoGuXuoCVDhWErDmrhFPivHtXPXXvRZ/I1/HYsJ8W2sqaa5dYzhlPWS7a7jMRXryfhlFd0+Ryt/+jP0"
                    "+co9XZH5BszcBxStKYrOetQPmNiVh1JWxVUOj/WzxHQmcZYoIXOa2SXqJaND+sfr9xT0ihZUsxyZzk9qIy0+P0pY45GREHos5Ky4ObRqvcNNeY5ZpAHmGGUs"
                    "LTXAzr1vPUVzK1mz6vmbHuNIUtavax7jfibjduhhRiZJom4r3JdFyll5LWpPvPSXcxecN9bq4xQ2SZjsTkwX0gejuKhY3K1c3yKw1rOuxey89ONzwS4ltbF/"
                    "pGfdsNc10kjS0ABjXFxdBnP/AFhesGVlecS7Acaquc4thqYwZGy1rgAI3XMi6Rrg1oc9ro6cUgeRuRzdAYdieqv6hMG4K8Z6r+R1TiJSCRKr8CsKoiUv3D7c"
                    "cssPmYypffsiW5TrW0kybQskESCy9/01eiMsLq7uD1GsLQC+LIZWKc+AHCwy97o3sdg/7v5iNlYZH9TPrlHMJxzOVz2nsyXHYyavs9yXQOpmNwG+3j41pvna"
                    "vhn16OuDj26bkckWeEdQONrQlVhS5pieOYHfLNpLyiZocv4pocXrcdXJW40mXOusB5AS2zHSmJTpeW48rWfKvox9OMrXkPFb+a4reIcYfesvzWP6xoNZNBb1"
                    "d9p2iS6O77rXO6wHtb7R2hxT60PUTGTws5dicJyWkHtbO+lXdg8l7JJ6pInRyT0ZZwP5I3QVYXgac9hd1ttB9EPqC8D9deJTrLjadLx3kDGIkKRyBxHlC4rW"
                    "YYkic8/GiWzCorjkDJsQtJEZ5NRlVI67FWoirryJj+Ss2GPwIDeqHpNy30nzQxPJK0clex1vxmYpF0uNycDXEdcEjmtfFM3RE1aZrJY3B2g+MCR0+fTP1U4n"
                    "6rYP+M8ZtP64HNhyWLthsWRxdroa50NqFrnAt+4GKeMuimYWyMOngLegayWyFS/6kPWI6+OOOozqD4+xblmihYtg3OPLOG4tWu8Y8bzHK3HMU5ByPH6WAubL"
                    "xl6dNXGq62I05MnSJMyS6lyQ884pzZWW8B+k30t5Lwfh/IclNyduQzvGcDlrra2Urx1m2sjjK1ucQxvoveyMySvLWmR/QNNBKrZ9QPq39TuMc55dxvGUOJyY"
                    "/AcmzeHqyXMXkJLT62PyNirA6aSLLxRyS+1E3re2GJpd1H2xpbGemh6p3WZ1GdaXEHDXLHINJkuA5o3yCzfVbXH+E0MjdBxbnGX1cmJZ0FJW2EaQxcY1AJxJ"
                    "ynI70V6Uw7HWpTTrOvPqB+nT0/8ATH07m5RxuTPyZJubxWOAyeSr2a4gu/qnSubFDSrOMgEDGhznkN2ftJOxsH6e/qN9QPVP1BPF+S1eNV8a3AZXKh+Ix96t"
                    "adYpT46GJjpbWUuxmHpuvc9rIWvLms/vA0ObJa0EFVOVeV9xq8X+mvG20dndx7hRmlJ71vGao9J0fkknvRkRbIjMj1vV9/ET08V4uN/y8cwnnsRrGVh9wJ7E"
                    "jXgkfg+N0M8/j3z3nX4/tlynex+c5fLjr8bI0Njse4+FPp+nzWlfXtkv8o6V+XP58d3KHTzrevzpPju8mWz/AB5if9bZH/RvxoD/AOnCM6/H/VOQ33+dnWz+"
                    "w+NFSp+iFhb6hcrOhr+ykYGvn/rOLx+w27f7kftu6aKwFZ6qxHqaetNnPHXLb3CHRvkmPxGuOrCdXcpcpPUFRlarnN4Uh+vm4Bh7OQxLHH26XEHmn2cvv/2q"
                    "0mW+Ummgo59Czid65k03/Qb6V6fNOPf2t9Q35bH47KMaeP4ujJHSt2KhLXfxexNNDYc2CcAsqQ+w33YSbHW4Pj9uE3r39U9rgufHEfT6HDZTLY0k8hyOSZPd"
                    "o0rBaQzEw16tmmZLbARLal/VarENhdE55cFGEv1ufUSaQt1XOVcpCUmo1L4n4eIiJJGZdpN4CalLURGXaROKWZEltvvMiVvWx9HnozVgfYsWeUwQQRvmsTzZ"
                    "ys2OKGJpfLJI/wDhoaxjI2uc5xGukE9td9E1frH9artiCnWxPCrFu1LFXr16+Cy75ZrEzxFFDGwcgLnySSOY1jR0guIHbatien5E6zZ3DTGeda2bsW2e5+xT"
                    "XmO8ct4VjGJWPF2NuQnJDNdlzlBjuOvyc6uTmtSMlpZUT6bEFwomPsLk2Ue5sJtdvqUOAQ8ltUPTmvkP4Bj5Zqzcnk7362fLyxyFhtwBscMUNEhnVV0z3ZWP"
                    "MsjgCxjLEfTn+303HKt71GnxQ5DejjsS4zDUH0qeJjkZ1tqSGazcnnuMDg208z+02RhbEwAFx0n6+/Wy4v6V8vteG+FcTh86cwY7ZN1mdT5F8VNxfxxKaVYM"
                    "W9JY3NazZW+XZ/RzIsSPaYXTRaqprFSpsK+zyjyWnkYy/tn0e+mPlfqdUhz+RtN4vxSYn9PfsQOnyGTa1waZMbSLomurb9wC3PKyMvj6WMka4PGqfWH6l+Je"
                    "ltmbBVK8nJuVxsBkxdWZsNPHvewviZlL3TIa8j9xO9iKGacRSiXo0OkwLZV63vqMZNZT58HmXHMDalv+6zSYFxLxiWP17JNoa+mrm+RMa5KydDK1JN5Z2WX3"
                    "EpLzq+ya2x7TLUyMN9H/AKO46NrcjTzmfma1odLfzNqo17gPuf7OJND2y7ew33JGjQI13BhpmvrD9YclITi5eP8AH4C55bHTxEd+dgP8rHWco+2yUMHdrxTg"
                    "c4/zBwB3jX/6v3qRtSSlNdV2Uk939xb464LeZI/9wvvhyuKpMBZaWvTSohoI+1RI2232++76W/RCSL2v7HuYHdutmbzrZddtkSfxEkE6G3BpP9AXLH4/qm9b"
                    "2yiX+1UDiNkslwuHdGfGgWNpR/bsntvXcfIbrYDiL15eu7ArtuXn+QYLzpjy20lY0ecYJjeK2iiYJ5wk0GS8T12BxqGVNUplqVYXmM5zFjMNmuLQKeNSl4Dy"
                    "b6MfTfIwSnjeRznG7ZG4DLaGWodegGtlitsFr2yducWWnPBHY6+1bB419Z/qLjrELeT4fA8ipGRrZn1K8uHviIuBe+N0Us9SWZjARHGa9eJ2/wC8kB+4Wceg"
                    "r1IOEevbGrQsNYn4HyriMGBLzriTJ51fLuIEaRHr0SMmxC0hLbRmnH6Lmaqkj5KVbRXEaW3FbyvE8SmW1TBmQQ9VPR/lvpJlo6PIIYrOPuGQ4rOUfcfjshGx"
                    "7x0B0jGOr22xtbJNUlBfGHjpfI0OcJ3el/q3xL1YxD8jx2xJFcqCNuVwt0Rx5LGSvaDqaNj3slruf1MhtROdDN0EtI2ApChqpbQQEQEQEQEQEQEQEQEQEQEQ"
                    "EQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEUCP6gPpylcm9MWFc8UMOym3/Trl7zd6xWw/3BS+LuVn6LHsoluRWWjltIoM1puMsktLlLpwqDDKrMrCyjlC92yrJU"
                    "/SLzscU9TBgrc4hxnNan8Ik6nEMGUrufYw7zs9DQ+YzVS5zd6s9LXsBO4vfVpwb+1npdYzFWAS5ThdkZ6AgAyHHdH6fMxM00vef0TzYZC06kmrRfa5zWBU0E"
                    "SDQflSiMjMjT5TvtPXxrwe0nvwf9vgvBatiHgbB3289iD28DtryPJ7fvpVNOhBOmtafka+4HeySCD3BBGu2tHv8AKukegV1Hlyx0pZLwtc3FlZ5f06Zk5BZK"
                    "1flTpZ8acmv2+V4Q6ie+gm011ZkkTkfCKKmQ84qixvCqWG21FrV1jBVLfVvww8Y9V72XggbFjeY1os3XdGxscbsg1kdfMNaGEsJN1v6hzgGOf+pDntLyZZLa"
                    "/pM5iOTekuNxc0vXkeHTy8dssMhkc2lCTNhnEv3J3xkkEbuovb7kMgY4NHtxzqiLqk2ogvW96hW+D+hPNMUgSrWNlfUXc1/BdI7VMRn0xqC+iz7/AJPeulvv"
                    "NuwKS14vx3LcMRaQ2JcmLk2XYywhpkphzYu/fpp4SeberXH4pYhLjePudyXJhzSWmHGOY6pHsAgOlyMlNrS/TddQBL+lrtF/UdzQ8J9J+SWq8rYsnmYBx7Fb"
                    "e1rv1OV3BNIxpIdKa9H9VYdGzbiyJzndMbXvbRgdkm464tbinFLWpa3DPy4pR96jV+O7ZmZqPXzvxoXIF51skdjvuO/Vvv42QSdEgedjWiNqmxkAaxjWtAAA"
                    "a1oHYNaNAD+g/rofuFbz/T1dOJ4dwhyZ1RZJBs4dvzBkK8FwddhMgrq08a8ZypMa6vqqEy2U6rl5ByVJyjHr/wDdZBnPh8Z4tOgQokNX11xVz9ZnOTn+f4/h"
                    "1Of3aPEKDRajb1knPZXpsWmuDgGl8FIUIW+2HDZkDnuOmR2kfRzwZnHfT25y21C2PIcxumeKVzQ1wweM669BvubJfHLYN22C7Qb7waGjTnvrSdaXVJkfV91I"
                    "8lc73doqwqcmuJNZxrCbftfoMc4cpJ09rjHH66FbPOu1Kjx+WWSZPDitQYM7kHJs1yFqsgLvX4jU5/RL08pem3p7gcPBWjjyt+lVyvILbAwy3MrdhE72zSM7"
                    "PbTilZSgb1PbHHEGiSTfUYQ+ufP73qR6i57ITWZZMRibtrDcfqPDmw1KFGV0EksLCd9d+zBJbklc1j5Gvi6mR+21jfl6T+kbnXrR5JPjPg3GotpYQoRW2T5P"
                    "kM+bQYFhdKmWxEOxzDKYtTfLrDmPvKj0tRWU99k18+xPcpcesa+kyOfTdr1U9WOL+k+Djy/IHTz2Lc0kGJw9LoN7JTQ9DphC6Ue1DDXbJG6eeU9MZljaGvL9"
                    "LoelvpFyb1ZzcuJwRgp1aUUVjK5q60yU8ZHP7ordcEb457U9h0Mgirxui22OR0k8P92Hzfx/00HJr7Pvy+urA66U4RqKtj9J2R3MWJtW0x3LhzqnoHrTtTpK"
                    "pqKWlNZn3pgs9pIOFln64OTGzManCMG2mXu9hlnIZCSwxh/lE8kRiilc0/zGOOAPA0BGTsTVq/RJwttSJlvl3JpLgY0Ty1o8VBWkkG+p0MMtGxLCw9ixkk9g"
                    "x+HumA7xodfnph819AMfEslzDJcX5J4wzS4TiVNyTi7UqgV/XBU1vkH9O3+D28uysMblWVFQXltSOV+QZjVTItLaosbunsirKy0kl6KfUdx/1auTYGbGT8e5"
                    "NBUkuNpSTi7SyVeEgWJKNgRRPjkhD2SPrTsc/wBsyOZNIIS50a/Wn6bs96UURyKrloeRcYluR1JJ21DSyGKfY02s29CJpobUUspfD+trmANf7IkqsErnRxvx"
                    "LF+JIZkxZ0+DJYcbejTaufKqrSJIZW27HmVllCdZm1ljEebbkQLGG+1LgTGmZcN5t9hlxMiLdWpeq2aN2CO1TtxSV7dWVofHPXmY+OWJ7XbDg8O0AQQ0nY7l"
                    "RyqT3Mdbp5HGzPq36NmC3Tss0XwWoH+5DKA49w14aSw9nt6mP21zmr0XOgHnu66m+jngPmnKPplZdlGGuVWayYRe3Ass5wO8t+PM2ua+PtSoVdd5XilxbwK5"
                    "x192thzWYD0mU7GXIdo89T+JjgvqDy3ibHdcOFzNqvVf226lKW2qJOnO2f0c8HUd7J2SGnbReF6b8qbzjgfE+WBhjdnMJSuzxuBBjtmP27kZ21veO1HMzYHS"
                    "enbHObpxq5fqIFqT154SW/H/AEi8RGREXwZ8ydTpGZ/BHv7db+NHr52U/vog/wDyc8uJ8f21/I2QMFjS758NGj37bI/dQK+t1gdznhRDQSOJ29k/A/jEmv28"
                    "77DZ/wAlX5zp9tWB5ySz7UnhmUksvOiR+xWJKPxs9JLzo1H42ReTIS25I5zuP5wE7H8GyoJH70JwTrz30f33oDXgxR4dCf7Y8T+3euTYE/5jK0yPJ0O/7a7g"
                    "k/K9A31jeoy76cOhTka0xDIDxnPeU7Wi4Xwy2aKYmdHezIp9hmztNKhKberMgi8T47yFPx64TIjLp7uHAsmHHJcWPFkU8/T3wqtzz1X4zhshG2bF1ZZ83lIn"
                    "EdMtPDwut+y4Fzepliy2vA9oDtsld1McwOCt59eOZWOC+lvKM3Re6PJSVosTjJGgExX8vMyhDNtzXNHsCZ8w2BsxhrXNe5rhQWJ5KEtMtJQ200hpmO0wgmmm"
                    "WG0EhlhllBEhDTbZJbbQgiShKSQgiIiIXNNDGMbHG1kcTG9EcTWtaxrWtDY2taNNYxrQAxoaAAABoaVND2SSvfNKXySyvfJLLI8uklkkcXSSSPP3Pe95c9zy"
                    "dvcS5x2SVP56eHog2PVZxji/P/OvK97xnxjmRXy8UwXj6lqHuSsiqYE1ypq8xkZrlachxXFamwsINnJraFfHuXy7yhVUXn7/AEarBdciEHrT9WF/hnJ8nw/h"
                    "WIx9y5hZP0mSzOWdPNXGRYd2KtelWkrGSOsD7Mj5LLSbAe0MYInNkm/6MfSliuXcWxfLua5PIwQZuJl3HYbF+3Ue3GSsJgmtWpopJmz2R0StbHGxkcPb+8Mr"
                    "ZI99eRP03XAkvFJ7PEPUnztjmdbYXVWnK8DjHkfDCS3JaclRrLGMIwfhy9cOXDKREjToWYxTrZL0exegW7MR2qnafxP1qepFa1G7MYTi2TphwMsNWrfx1ktO"
                    "wQyyb9xjSAS5vVXd92gSWEtO38t9GvplarPjxOQ5JirJaQyd92vfja7pPSXQy1Iy4dWuoGUgt2GhrtOFYjqT6eOUuk/mTKeE+YKeLUZliyoMtLtZNds8fyGh"
                    "t2jlUWV4nbyIVY7d41cMtvtxLBdfAlRbKBa0FzX1WR0tzUQbBvTf1EwHqfxenyjAPeyKdz693H2HNN3GXoe09Oy1oDXaDmyRTMaI5672SNDQekV9+pPprn/T"
                    "Lk1vjWdayZzGMsUclWje2llKMoPtWoA5zzE4EGOxXc58kEzHsL3N6HvmB/T/AHUplmEdUtn06y7aTJ425zxTJLGJj8iwmOx6vlfBKZGS1WRU1c7MKshLuOPK"
                    "HM6XMJUWu/dr5NLx21KmlX4jHjiLv1pcHqXuMYjntesxmUwlyDEZK2xjQ+zicg+QUo53Adb/ANJfLxC973CNtt0Yb/ebbJ76LuaW6XIs5wGey9+My1SXOYyq"
                    "8n26uTpe0zIOhOy1pu03wPkiAaHGo6bYcHh023rN9b190g9NcDHONbVym5o5+sLnCsNvYUso1xg+I1ECM/yLyNTISpMlVtRxrjHsWoJ8cyPHsszrG8ldRNaq"
                    "FVVjFX6cvSuL1R57FBk43P43x2GPMZ1oaC21GJmx0sa5zvtY2/YBbIXB+68U7QxxPaVX1Cep8vphwGxdxr2N5FnJXYfj/Wde1akifJYvdI+5/wCgrNfYawFn"
                    "XKIme5GXB4oqJWhppEZlHtsR0NsMtI7uxDSCJLbSSWtazJH2kSlrWtZ7UtRrM1quChjirQw1oI44YIIo4YIImhkcEUbGsjjjDQQ1kbGhjWjuAAB470+WH2LV"
                    "ie5cmksW7U0ti1ZnPXNPYmeZJppX6AMkj3Pe86A6naGh2EvHQ16NvUN1q4c3ylPzCk6fuH7RJniWc5Pik3Osizw4s6RDnSMS49g5Pha04ySoshhnNL7LaliV"
                    "JSxIxvHctpXzt2In+sn1UYP09ytnjHGsezk2foyezk5X2XV8TjZmOc2ao+WNr5bdxhDeoRBkMLt9b5Tpqlt6OfSrled4qryjlmQsccwd6L38ZRhgZJlshXkZ"
                    "G+vceZx7NOrIHOcyN0cliVhaSa56miUvO/02PHcjFYrPFfVbyHj2dNLjrnW3JOAYvnuH2DbaHTlR4uNYpZcX3tQ5KdNn6aa/mV6mubSs3oFs4pJpj/S+tnnc"
                    "VrrvcZ43apl+/YhdkK07WF4JH6g2ZmucG7APssBJBcCAWnflv6LvT6WuGVs9yStYDD/eiTHyMMvQQ13tyUnAM6+kuZtxLQ5rXtJ2K8nWL0dc09EHLCuLuYqy"
                    "v3awJWQYJmOPzVWOJ8h4mzYLr13GPy5EWDMZsKqQqJEy3GrKviW+K2FhW/UtzaG+xTJMjm76Ter3GfVzCTZHCe9Sv48xxZfDXHsdboSyN6o3se0MbaqyacIr"
                    "cTGND2mJ7GSDToS+rPo7yP0mzFfH5h0WQx2RbLJic5Thkhq3WxECWGaJz5v0V2MFrhVfLKZYR78L3NbPHD2foL60cy6IeoLGeWaKZKewiyk1OOc0Yk1HmT2c"
                    "y4uO1N26ZbqoTiH52XYnFnWuT8cSEH9RDydLlT3FQ5VlNdceX66+lGO9U+F5GsKkQ5PiKti9xq+0MbYbaijdMcY+Xpe51PIEGJ8TuwmLJYvbkLnr1fQb1VyP"
                    "pfzOg+W3J/ZXM24KPIqT3uNZkU7mQRZZjHSMjjs0CWukmae9TrbK2YRwNZfx5h5oouNOnPlLqIrHoeSYzgHCmb80V8iE79TX31FiuC2ecRHoj8dxByIdpXwG"
                    "1sOsOo91h9C23E9yVFT3isTNkc7jcE4GGxey1PEuDyGuhmtXI6ZDiQ4NMb3kOJDgCD2PhW9X8hFTxN3KtIkgq46zkA5oLg+KCs+yCBtpcHMbsDYJ3rYK817L"
                    "czyrOcnyXM86vJOT5tl+Q3OU5nkctiNDlZDl1/PftMjvpMKA2xXw37S3lSpioNexHroCXihV0aLAjRo7V7HHuPY7imFxnHMPWjqY7DUocfWhiY5g1XYI3ySd"
                    "TpHmWeQPlle9z3ukkft5JG6MeS5/J8vz+V5PmJ32b+buTX5nueXiNsznOgrQu6W9NerAWQQNDWgRRNJBJcTI76cvpick+oJaZNdxszgcT8OcfWdTWZXn0zHX"
                    "Mvtchv5ao1lLwLCKFnIMdjx7pjGnEWVxlt3ZPV2KHdYo4ziuafulhCqdA+v31AwekQx2GxmNizHJ8rA68IbMz4aePxwe+Blif2Xe/JLYnZK2vEOmMsilLnkg"
                    "sG+vQH6fj6tDKZrMZGziuN4yf+HsdTiY+5kMiYmTPjZJYifXjr1oZY3TdIdO50zA10A6Hyz4yf06HR0uvfjxOZeqRqwNl0os+XknDEppqSZKNhyTDjcG16pM"
                    "Ztw0E7HZlQ332Em0may6s5JRBi+sv1bZZZNJW4nLAHDrqfwi3HG9m/ua2VuUM0biPDw8gOAcWOG2mXcv0d+kLqjoIhyaGwWkMu/xoSStfrs50T6prPb1acWG"
                    "IfIa5nYiu36gPp1ct+n1luK1WW5DW8mcfZzDmO4dytjuPWOLVlpaVK0qucYyHGZlnkpYlktdFehT4UA8syCPkVS89aVM9btXkFXRzW9CvXfG+sdPI15MeMHy"
                    "PDMgluY0WxYgtVZnPaLtB7o4pvaika2KwyQOcx8kbi4tkaGwv9c/QS96Q2sfcrZCTN8bzEs8NS9JWbXtULULI3/o75jeYZH2GGSWvLEyJhZFJG6PqYJJMM9F"
                    "/UvkfSl1M8R820l8xR1uMZVXQuQDn/XLqrniK8mw4XKeP2rECSwuQ1IxFMy4olyCl11PnNDhmXTK2xPGWIq8h9c+D1uf+mXJMPPDG69So2Mzg7Ja10lbJ42J"
                    "1qNkTyWmNl2OGSnMepoMc4c/rEYace9B+Z2+Aep3HMnA+UY/KXa+CzdZjulljH5SdlVs0jCHNfJQsyw2oXdJkEbLEEJa6wXL0iBSsrm15tvWe+ZdYfVkk1Hv"
                    "/qd6gEmWiMvt5by8iMy/giIjMjMiM/P/AGvD9ID/AOyv03/bgvFQT28fwSj89ta1871oHfYbpF9X4gfVb1JIb55xygnzvZzV0nX52NE/jx8d9u/RakEr1Kum"
                    "5BKVtb3LhGRb0eun/ls9nsvBaT4+C+C8b0NPfWIQPRmyCRs8n4+Gg+er/fz2+Ozer8HyO63J9G0Rb6wyEjeuH57vvsB+swg8Aj8tGu/ne/t7X1/gVOq1ZeVN"
                    "xk+R8acZkXdr/TvB06Pu3/8Aa9UWlGZJVvyeyPyZ7/kxfVxJwHFOMbcD/wCruC2d9j/1XUHfto9tjuAfnROgKKOfREc8512+7+2nKtft/wBfZDTfJ+APBIA1"
                    "+DuwL+nm/wB3r0y1zuIya6UuWD8mW1Grlbp4SXaR6PelKMyLf2kRn8CKX1uE/wDRzxga7Hl8YP43/CMg4dh/n3J8n4J0pSfRM0N9QOWbGiOKRnxoEfxSuPPg"
                    "nvv86HjQK379Zb1XEYNGybpA6ZMsSnOnjk0XPPKOOzpzMzj+KtpSJnF+DXNY/E+m5BnEtLOc5NBmuu8e1f1GLVzSOQLeTbcbaP8Apm+np/MLVXnvNKfTxOnJ"
                    "7mIxdmMg8ktMJDJZGO0RiK72lz3HTrkjWsiBg9yRbz+pT6gWcIqWOD8OtNk5lfh6b+Qge1zeNVJACXE6cDlrMZ1UhILa7C+zP3ZFBPUtVJjx2O5S40GDDj7N"
                    "S1MQoECDDaPfepZsRIcKIw2XcajZjRmGzUZoabPts2kkgrQue90deCtA573fbDBXrwR9XVvtHDDBGwn/AAsiY3W2hvasZlezcsBrWT3LlycNa1oltW7lqzL2"
                    "DWNElixZsTv03pD5ppn6HXI9Ww/R29JVOLf011e9VmJMuZYtNPknAHD+RQpKXcCcaf8A3Kv5Z5EprBllBcgrcRXzOOMTmx3v9NW2SzK5N3kewoKziOs76lvq"
                    "IfyuxPwTg2Qlj4zVMlfO5as90Z5DP2a+pA4ac3EwEOjkP819/UXEVf7uWzL6a/p5h4bUh5tzShDNyy2GTYjHzsZKOOVek9Ep31NdlrAPW+QdqkfRDGDM2Sd8"
                    "qnqp9Ttr0o9FHKnIGJZC3jXJWUftPF/Ftmn2DsomYZzLODKuaBqSS4z9/h2GR8tz2qblMyYf1GKk5OiyYTchleg/RLg8XqH6mcZ41aYZMbLafeyzW627GY2J"
                    "9uzEdkACx7bKpOx0+/1b7LfPrHzSX0/9N+T8nq9JyFSkK2MDyQwZLISx0qb3loJDYpp2ynQ27oDB3cF59JPIbbQ02lLMdhHttNIUZoaabIkpaStS1LUlBaQR"
                    "rUpfjbilLUpR3WV4IKteGrWiir1qsTYIIIQ1kMMMY6Y4ooxtjWsaA1unDQA1okbpVsPs3bNi5bmlt3Lk8tmzZnLnzWbE73STTzPI+6WV7jI9xHdziRodIUun"
                    "Q96N3UP1tYExy2edYlwLxFdFOZwrOsjo18n32ZSqa8s8fvHKfjSizDDnYtHXW9RZ1TtzleaYtYPT4ZvU2MX1BLi3y4n+sn1UYX07zNri/HsZ/aXkNBzI8pJN"
                    "PJTxWNnMbZDVe5sLpbthrHw9ZryNjhf7kckhli6TLL0b+lXKc/w9PlPKMnJx7j+QY6fG1akMM2WyVYuDIbgfKXQ0K7yyR7I5q88tiJ0MrHRMkPTvjM/TP5yz"
                    "G92B1yYtYzGm+9UCb0vXVTFnukaT9lFo11K3j1O25pafqXKi/UySkqOJINGlaQpfW/yZs8RyPCcLPWDwZmVMlkK8pZ92xE+ZtlrH9wGue2RrfuPQ7YA3ne+i"
                    "jhj60rcbyzkVa2WOEE1yDFW4GPIHS6aGvTpPlYD3cyOaEu7NEjO5MCXVB00cs9H/ADFc8F81QKavzenp6rJY8nHblV7jmS4hkMq3hUOX4xZSYVTZy8ftbDH8"
                    "gq2HrijoLRq1obqvm00V2vV3za9MPU7jnqrx4cg4+6SHomNbJYq2WG7i7fQHiCx7Y6JGSs06vYjAbO1rugAsO4SeqPpZyH0q5CMHnRDZisxGxisrUa9lPJ1m"
                    "ODHviZIC6GaBz2ixWc6R0PuREyu9xq5zo457ynpx6meFuYMXsXIC8Rz/AB5eSMe++zFuuPbSfHquSscmoQ4hh9F3gsq+h16p7ciDVZGdHkX0rk2ihLa6frTw"
                    "upzv045Vhpq0c1mLFW8piJXj76uUxsLrkMsTmuYeqRsJgc0uDJBNpwcGlp7folzC3wX1N4rloZ5YqdvJ1MNloW9PRZx2WsR0Htl6mu02vNPBbD2N9xrYHNjI"
                    "9w79JYUnK6NARARARARARARARARARARARARARARARARARARARARARARARARARARdI5M48xblzjjPuKc4hOWWGcl4Zk+A5ZXsyHIj0zG8vpZtBdxmJbJk9Efe"
                    "rbCShiWwpL8Z00PsqS62hRdzH3rGLv0slTeY7ePt17taQEjosVZmTwu20td9sjGns4H8EeV1btSHIU7dGy0Pr3a09WdhDXB0NiJ0Ujel7XsdtjyNPa5p8Oa4"
                    "bB8z3lrjjLOFeTeQeJc5jTY2XcaZtk2CX7k6Edeuzn4tczaVeQw43e4SabKW4beS468ha2JuPW1XYQ3H4kuPIdvN4JymvzPiHHOUVXsdHmcXVtvDHh/t2iwR"
                    "3Ii5oGnRW2TxFjmtkZ0gPawgtbSLz3iE/CeZ8j4tM3Qw+SsQVz0uaXUZCJ8e77thxdRmrOe5j3sMgeGyP0VJX6KXUSXBfXjxzV2lrYQsP52g2HBuQw0WEhui"
                    "O4yyTBsONLafTNE4zZ3rfI1Nj+E4/OdbS/SQORsocbkswbC1RI0R9W/Cm8o9L7OdrwOkyfDbceWY9jQScVN01sqC7XWGRRugsloJj/uHlzdtDmb1+krmL+M+"
                    "pbeP2JmR43mVSTH6kLg45em19vG+33Ef97GLkLgWh7nGBrH73HJfOFTqtPVJP1+eoz/VHrNicPVip6aDpnwmDiLqZaGWYsnPuT4ePcj51PqjjyXVT6tzGC4h"
                    "pPfsGo0qDkWK5NEjRkxVlLsbOfou4W7EcJzPM7DA2zyu/wDoqbu+/wCE4d72HYIAImyJnJcC4ahY0FrusKtn6zOXNy/LcDwqu8Or8ZpOyt4Atd/1rlW+3XHz"
                    "0Or45r+phALm3mvO2iMiFTDcZy3kPMsU4/wKlRkedZ5k1BheFY+7MZr495l+V20PH8YqZNlIW1Fq4lhd2MKJMtJjrUOsiuvTpjzUWO64iV/KuQ0uLcbzfI8i"
                    "9sdPC46zkZ3PLwC2tGS2MiNsj/7yTpib7bJJNv8Asje7TDFPinF7fLOSYTjVESm1mcjWoRmJrHyRtld/fzNY98URMEDZJtSSRRfYeuRrCXD0kOOenTGeLulz"
                    "GelvELi4pccxjhpriGDlcE4CMoSZ4qvHp2bqf+hRXvZdY2D8vKZ892u9ibkUuTNkRVE+42dHGV5JdzHKb3K7wZPev5uXNzsk6zE+aS4bfsEF7n+w3tC1pkc4"
                    "RNDeskbV2+NwVLFcep8bptMFCliYsRAIw1rmQRVRVDm9LWMD+kF/2sa3rJIY0faIPWf003Ts0w2wXVL1KoJpv20FHrOCmmUJSSUtpbac4lkGhDaEmhKPcUXa"
                    "ZFvZGZyZ/wDTO9UWkBlDi7I2NayONuMd0sa1vSAD73Ue2vJPcfKjm76QvSl73ySHkEkkji+R7svZ6nvcep7nacO7nbcT52exA7LaTph4x9PT0csd5dqMm6vc"
                    "T/q3OJ9Nc5c/y3n3HbHJq6fDayQ3SYzR8b4RCqLezRXTMjyC5bjUeHTb6dKyg4yykRItRHjYDzXknqv9QOSw96Til/IOxtOSlU/gWGujH7szmeazNYLXVIXz"
                    "uDA5zpooQI2hgaNNGe8O4z6Y+huNydCrn6WMhyVwXbP8azFZ1xz4YWwMijbK8WZGQsaR09Mjx9znku63nvkr1yPTFjLNCOoW8mESe43IXAXUe+x/Oikf6RoY"
                    "UryWyS4et7PWh51T6dfWm51e1wHLx9Otm3JQpjvrwbVuEO8jYaSRvRAK9C567ekVHXvc+4+/e9CraddPbe9tpsncB2OiQAfglR/epl6q3QF1P9E/MnDnGXJ9"
                    "xlvIGSOYBMxGqn8R8s42lVnjXJ+G5JLeZucowmmrq1xukqrY1PvzYyZEc3q/udVPRFk7o9D/AEH9XuIeqHFORZjjMmMxOPtWjkrRyuIlDKtjGXa7g6CrkJZp"
                    "gXyxj22RyHq6X9PSwubpf1s9avSjlnplyzjuK5NBkspkaMTcbWjoZEF9yC9VsRFss9FkMOhE4ufJJH9ge0OBdpVJ/qiMjPuSet+TIy/J7Mi2R+T87/jxrwZH"
                    "ZkHeCR4J7f5b+NeBsa87I3o91WoYPAO/jfcnf9N/gn5799a+Dfi9D9z3PTF6cldxKM7XnxR6IvCl9SfL7hkej8mZrNR+C/u1rxs6dPqcAHrnz0gEdVvEPO99"
                    "y7j2IJPf9yQf3BVvv03bHonwNpO+ijkWDfwG5vJgN7/8I0B+wA+FAF+osc7OvXBiMyLfSHxGadmRF3FzL1OkZb3vyRl8fx/OhL/6IXa9OeW+NHmx3vyP+osX"
                    "ojtrts732+CDtRO+tSPq5zw0+dcTsjX9MxMd61/p+6r051K7cDzk0kR9uG5SfasjIlGVDY6JRpUSjJXhJkRpPSj0aT8lLLkrmnj+c34OGyvbvoj9DPvz/wCW"
                    "j2HyQVFvh0IHMOKE7AHJcF3B8f8AWtM/H79vJ8dwRtX/AL1vuDrLmnoBz6woaGfkeR8I5LjvOFVXV0hlmSzW43GucW5AuSZe7SsG8e4qzXPb5ytYcROmprTZ"
                    "rG5lkcSvl1GfTZy6nw71e43eyMrIMfkxd4/bnk/lgGYrur1pXOJAa1l79KZHEH+66wBsgi1r6heJXOZ+lHJcXj43TZCo2nm6cDQS6xJh7Ud2SBrWgudJNXjn"
                    "ZCBoe8Yy49AcDQlOUnuQaVpMlEhaFkpKkLQZdxLJSTNKkGR9yFJM0qL4My8i4fqBJ0QRrYLe4II7dLh2IcPBBIdsEOO9qocQOaCHAgtJBa5pa4Oa4hzXNcAW"
                    "vaQQ5pALXDuN9lN10EetvzR0iYVjvCuc4Rj/ADXwvj0l9rHkzbmww/kTBKmwkMPyaynyiNAyWnyXGqx8rKwpMWu8Xh2seZbv1qOQa3F4VFR0URvV76UsNz/N"
                    "3+Vcbzf9m89lHGxkKdyu61hbt7pcZLXVB/vlOe28t/UOjbZiLgZmwGVzxJLT0m+qTJcCwdDivJsHJyDC4pjK+Ou4+eODMVaOwIqbobT2U7UdRhIgc6xWk9oN"
                    "hc53S14na4c9f3oK5HrXn8/k8r8A20eS3EOq5B49m5cxYq+mQ+9YU9rwlK5TYTTIcNcduVk7OK2Li2zW7UR21oUqI3I/pW9YsAHSV8DX5HWDnATcfvQW5ekE"
                    "hrnUpjXujrA6tMgkDB2kLHdlLDj31NekOfcyF/Ijg7Lo2OMWeqT46IPd/NE269jqbnsPY/34D/MRe0OI2SzDgz0zPUyuKjkmzlcU9S19jWLxsfi3XHnNuRIt"
                    "KPHE2Eq4j1F9XcX5/SyYLsG0vZryoeSwG7OrmWMmK6iK664wNf0uQ+rHpWy1iad3lnCY7FszWavtXMU2xbbH7Hu9UkUfvO9uHojkje5pawujdrbjn1zAemXq"
                    "O6rlL2N4nzJ8NURVbNiLH5h0FZzxMY2B/vGBpkkDpWFrHFxDZR2DR23hf0vOhbp65MxbmLiDg0sT5Hwly5dxfJHOS+X8jXUO5Bjtxidw4zV5XyBeUj652PX1"
                    "vWOKl1sg0MzVuMm3IbZeb6Od9VPUXk+Mnw/IeY53M4uyYTPSyN19qCQ1pRPAS2UOIMUzWyMIIIeAV28L6bcA45kIstgOHcdwuShZLHFexmLq0rLGTs9qZgkr"
                    "xxu6ZYz0PadhzexB0FVq9fXmqFyL19XXH1edk0x0+cacecaWzE4+yEeWZVVOc1WlpSpS8tKolliXKHH9ZOdNph92xxlxpz3GYkRYsE+i7j8eO9N8vnyGGzyH"
                    "kU7Q7Tg5tPEwRVYou+t/71JblJb1Ah7QQXM7QK+svNz3+e4HAA9NXB8ebb6QQ7rtZe3L1yOA3osgpxRsDtkAyOA1IVDphMLHbzNcNo8wyk8Iw+7y7F6bL82b"
                    "qrC/dwvFba8rq3I8uao6hp21uXMYp5U27TVVrLs2ecEosVtx51BHKHlVvK0eMZ+5g6k+QzFbD5GbFUa/SbFvIsrSfpIYS4Oa18tjobG5wc1riHuaQCFGDiOO"
                    "xN/lPHqWeswUsNazOPhylu072qteg+wz9VJZeSzUAhDxKA9pLPta5pLXK9ZRest6VOF0NBiGNc9tY/jmN0tVj+OUFJ0/dQ8OpoqKnhMVtPUV8Cv4dTErK+tg"
                    "Ro8OFCbaYYiRWWmmm22kJIqerPof60WZ5rNn0/5XPYsyyTzzSUpJpZZpnufJLLIXuc58j3Oe573FziS4k72rfIfV70jggjih9QuFwwQRsiijbnsZExkUbQ1j"
                    "Y4/fYBG1gDWhjegAaHjQ5j/62/pj7Mi6kZyjSRGokcE9SCzIj2WzJHECj7fB/d8f5+B8f+gn1iO9enPKiR5AxkpP+g+P38L6H1k9J26J9RuGAHwTyLGAH86J"
                    "sa7fP47fkKKv1e+ur08utHpWLHuKObmsi5u43znG8143hTuFubcdkWsWZKTimd44jKsp4vp6+pgTMRvZuTKiSLWNAt8gw7GESiNyJFkw9/fTZwT1a4B6m469"
                    "lOE8jx3H8pUt4rOWLVZ8NSGtMwS1rczC4CR1e5FB7ZAMkTZZZGnoEjXaI+o3mHpXzf0xytKjzbi2SzWOnq5TCVqGVo27s9yvIWyVYBDK97BZrPmZIOzXlrGO"
                    "HUWEVZUSlE4k0KNKkmXae1JWlSdaV9ujJST2ojSZGRkZ+D8FZWHEdR2QdbHkH58E99+Nb3+2idmtYwggAjqB8g6IIPbR18Eb38aPfYOlej9N1l/q29HCm4by"
                    "DIGpk6/4Y5w6XLCWb7KHKHHklmnHGC1ktUGKpUVVJxdY4Y20ZxZE04KIst8pcl5a3aevW3Gs4D6852xUrBletyHF8qpxdOo5W2zUzL+gdTepjrbrEZ05g62v"
                    "aC3Wxb56L5Z3OfRPj36md0lmfAXOO35HEmRs9IWcQ8u6g7uYo45WbDtxvYSDvSo9ZfQZPheU5NiOb0knGczxbIrrGMwxqa/FkS8dyyhsZVXkmPypEJ6TCkya"
                    "e4iTK92VAkyIMk2SkQX5EV5h5y27Achx3J8PjeQ4udk9DM0q+RrSRkkGOywSlj99xJE5z4pWPDXskjc0hulU5nuM5HiuayXGspE+K9g7c2PmY5hYZP07zFHY"
                    "jBJ3DZjayxC7ZBjkaepx2Fvd0E+pTzj0C5BkC+PodHnnHeayKyXm3FGZy7SHRWc+sUy2jIsVu6tbsjBc1kVSHKKTkaKXJ6uxqvoG8ixHInsfxdyj0/6z+gvG"
                    "/WGKpcs3rGD5Lja7q1LMQRNsxy1nOlkZSv1XPjM1aOaR87HRStsRuLxG8Mkka7b3o166Z/0hkt0W0Ys3xvIzizbxckrq1iC30xxvuULAbK1sz4WMhfDKxtd/"
                    "SxznMLOtWTuF/wBQx0bZ2+7W8uYtyrwPOjVzMty7nUCeTcFnzVrZadqKSw46O15CkzGVredORd8W47WuRY5uInHJdRDKDfKPpC9WMG6R+IixPLKzXHpfirzK"
                    "tp0YD3dZp5T9G7Ya1vVHHLLIXvDImygFwmzxn6sfSnONiZk7WT4takADoszRkkrxyfYOh17H/q64a5zndEriyMMYXzGEkMW3uUckemF6m+G4lxhk/K3DvNtd"
                    "/VUbNMV40e5RuOOOTIuVV1ffYrDtCwaLkWB8sQH/ANuyO8rmItjUsRJ7Fl76Ysn/AMG+jU8eJ9WvSbIz5OLF8u4Ze/SPqzZOKldqMNOV0c8sIyEcbqxjc6CN"
                    "7+iclpiGy3RC2w7JemPqdQjoSXuJ8xofqI7EdCWzj8lH+piDmRS/o5HvcXsErmsJiIPWQN7C6496J3pkSWXY8npoORGfbcafjP8AM3UE9HfadStt1t9lzlZT"
                    "bzbja1NrQ4lSVNqNBl2GaT+knrf6uysfFL6h8plikDmyRy5OaSORrmlrmvjf1Me1zXEOa4EHZ2CvjF6PelcEsU8Hp5w+CeB8ckM0GBx0M0MkTmvikiligbJG"
                    "+N7GuY9jmua4Ag7ClQQhLaENp32oSlCdqUo+1JEktqWalKPRFtSlGoz8mZmZmNWkkkk+SST2A7nuew0B/QdlscAAADwAAO5PYdh3Oyf6nuvNW61JJf8AWR1a"
                    "kRkR/wDVF1CF92zI+zl7LkfGzIiLtJXnX/IyL4F3/pE7/wBlnpwD2A4LxTxs+cHQ7nvsEg/02O2u4dSv6uRE+qXqN875ryYjsP8AFmb37E72T/8ADa279E99"
                    "TnqZdNaftIvd5dNRbMz1/wBPnL2jT/nu7e4z38GRa2RjTn1hOLvRuftvXKMAfH8o6ciPyfJP7edLcn0fwhnq3K7ZB/shnddxo7u4PY158dx8jXfetK/gKola"
                    "KvKL41f9vjbjhJ/CMAwsvuMkmZljlaRqMtmRGoyIzItl8FvZbO+fibyeK8YJADjx3CEt0dNP8MrbDe5cOxAH83Ya2RtUdc8g3zrmw/PMOT9h3PfNXjtw8E/n"
                    "89/+6tyelrqw5K6Scq5Iz7iZ6FW5xnPCGacMVGUSFKcmYL/WuX8aZBNzKjgm2qJPyKsrcGmwcfXZOHW1FvZxL6ZXXjFW5SWGI+pnplifVOvxrGZ+xJFh8HyA"
                    "Z27UhAbJlGxUbNaLHmU7EEMkthslh7Q6R0EckEYa+VsseUemfqTkvSybk2RwlVk+YzeDZhqFmwT7GMe64yxLfdGNGeWKOPprQkhpmk9yRzWxFsmuMme9Keff"
                    "kSHpUuS+/KkSpUmROmy5Ul5ciXMnTpjz8yfOmSXHZdhYTXpE+fNdemTJD8h5553YlWtVoVK1KjBFWp0oI61WrC0MihgiAZFHHGPDWNaADsHuXOJcXFa2tzXM"
                    "lcs5DI2J7l69YltXLdhxdNYsTOL5JXvOhsuPYN0xjNMYGsaA2xr6Gfpz8bc8Ld6veXbXH8zx/jTPP2LAOJYctyWqHyNi5Vl+WT8qw1xGWSbqGbHHrvBcUZlT"
                    "YFomZGybISdZRVVSYI/Vr608gw9qX0xwdaxiIL+Or2s1mHsdFNkqVwO6KeNf1bbTcWPjtzdPW+RjomFnSSJ1/Sn6Mcfv1IfU3Mz1stbr3rFfC4xjhJXxdio7"
                    "2327rC37733dULD9kLCHacXFz7ggrwU+1ov189CWGdf/ABjhXF2c8hZzxzU4TyPG5KiWmBMYu/azrSLh+YYYiBKTllHkFamCddmlm+akV/1KZTMVSHktk627"
                    "sT0y9Scz6Wchl5LgamPtZF+Os41n8RikmihitPidNJGyOSM+6Wxe2HFxAY941s7GCeofp9hfUvAf2bz816LGm7VvStoT/p5ZpKbjJAx8nS4+02UtkLRrqcxm"
                    "zoEGJi4/Tg9LlHWzbq36teo+jqqqK7NsLO2X0/wqeviRk+49LsJcviCO3GhstpNch5+Y0hCO5SnW0+S30z6yvVmxI2KHHcbkllcGtjhxUr5JHHQ6GM9+Rxc4"
                    "DQ00u7kt0da0e76Q/SOGNz5DnGNaCXSy5iZoZv8AxFxc1rQD376HfR8ncgdd19+mr0S8W8U8CvdVXG1jA4vwfGuNais48ds+ZckZjYRRV9GmXldRwrTZzMob"
                    "SwTEKdPfvYtac+xkSnkKedW6Y0w7059V+e5TJ8grcH5LbmzGQs5GaY4y3DC6S/NJZPsy3Gxe5A3rLWPa5zGMDAXAaW4f7d+mnCcfQwdzmfGsfFiqVehFHPlq"
                    "DHtZSiZXDZWxydLJz0bfEQ15cXEM8rqqvXV9McjX2885QskK7e//AEA6h2W1+ddyHJPFjCFJ/wA9xDIIvpn9bpgC3g1poOv+0ymDiI33AIfkmuaT+4C8Cf6h"
                    "/RqA6fznGvPV06gr5Gwd/wD7ilINa+d6/dQA+tt1p9KvWNkHTjkXTllDuXW+GUXKtRntvNwHMMMsG6+0suOp2C1zsnL8eoZVtCjvRc+kxYsQ5TVVInS3nERl"
                    "W6VSpk/Sn6Z8/wDTj+2sXM8S/E1csME/GRG/j7jJbFU5Vl2QMo2rIiIinqsc+QM90dAa5wicGw7+qb1D4J6is4bLw/LR5a1iZM3FkpG0r1V0Va2zGuqs67lW"
                    "BsoM0ExAjL/b3IXBvubME8mWpuJPWhfaaa2x7VJUfeRnBkF9qiUelF3JPuMtEZedkREctMqerGZNvlrsXkGH8EOpTgjW+29u7Dfj43sRUw0QbmcO/uCzM4p4"
                    "J/4mX6zge34LR2Gu/wA9tL1TBQWr2EBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBFTD/URdNTHGnUbhPUljlImHj3UPjjdTm1hHkS"
                    "3W3eXeMoMGj+usG5Tio8GVlHFysLg08GqJMaUzxllFtKiR7BybOtrGvor5x+sw3IeBW5i6bEzfx3FMc5jf8Aq+46KC/DH/jkMFsR2DsO6GTu+4NDQK//AKxu"
                    "EGLKcd55Vi1FehPHsq4Me4CzAJreOmeQCyJr4TagcXFpkkbC0B5JLa91ZcWVXZQLWlu7jG7urmxLKkyPH5sqryDHbmufbm1GQUVnFcak193SWLMW1qZ0Zxt6"
                    "HYxI8hpaHGkLKauYxlXOYjKYW83rpZbH28dbYPu3XuQSV5AAexIZIXN2D9wHkKGuIyNzj+XxmcoAC7h8hUyVXrLgx09KxHYYyQt+4RyOi9uXpIcWOcWkEr0h"
                    "eHesTCc96Gcb62MkadqcUj8EWnLvI9VVON3UzGLLAsdspnK+KQnfbq2LiyxLJsbyjGWnG24EeznVRLZTHZkI7KQM7w3J4bnOQ4L7bpMpV5C/AQB0b2GeZ9wV"
                    "acvts957W2GyQzAMMxDJB0GTsXXRYTleMzXD8dzOGZn8KvYKLOGUPaWR13VP1UwL3+20GECRjzIIulzCJBGQ4N86Pk/kvIeV+Rc/5Qyxx5zKOR83y3kDI0vS"
                    "lTExr7OMhscpu4ER/wBtpJ1dfaXEqFTttMsRotXGhRYsWNGjssNXXcR45V4jxfAcYptIgwWKqY4FwAL5a8QbYlOnO7y2fek11OIDwC4gbVOHMM5Z5hyvkPKL"
                    "XSJM5lLd5jW70ys6QR0WHbWnqjpR1mP+xpMjXuOydmZ/0BOmc+a+sKRzNfVtdPwjpgxv+r0/WLJ51fKedsXWJ8aMoq3kKjzYcCpi8k5X+4uKN/HsqxPDZkRk"
                    "pclibCix9ZHPBh+E4zhNOYNu8su/qrwa4+6zDYl7Xua4s+0Mt33QscHkdba0jWBw93pk39IPBv4jy7McztwvNTjdT+H0HOYwwyZbKNDpHgv+/wB6nQYHN9sE"
                    "N/XBz3MPttku15JkdBh2O3+XZVcV2PYvi1LaZHkl/by2YFTR0NJBfs7i4tJ0hSGIddWV8WRNmyn1oZjxmHXnFJQhRlWbBBNZnhrV43zWLEscEEMbS+SWaV4j"
                    "ijjaNlz5Hua1rR3LiAO5VjMsscMck0r2xxRMfLLI8hrGRxtLnvcT2DWtBc4nsACVR59RH1r+cOpvJMgwHp8ybKOFOnWFOOJWO43KscT5Q5LjwpSlIyDMclgS"
                    "IuQY1RW6G2XIPHtPIqNVb8iHyC5fvzl0NBZz6L/S3xni9Ghn+fU4OQcnsQw2P4TbjZNiMI9/RIKz65c+LIXIiAJpZgYmP6omQaDnPrl9YPqX5FyC7ewPArb8"
                    "Hx+tLNWObqSOblMy1vuQumrzFkb6FN7S72/Z6rEhEc8diPTCYO2ZLcRhuJEbjQorslDMaHEQ1FirmzHySzHiQmG0tOS5khftsMMNKfkvrJtptxxaSVLGe5Qw"
                    "9F8s8tLGYyvpzzI+ChShDj9u/wDsoI+px10jRc46Ac/e4qw47J5zItbDHfzOVsnTellnIZCf227JLv72xI1jO7i4kMYC5xDO6kAwT0sfUg5LifuGH9HXLBwy"
                    "ZbeWrOXsK4amF7yF+2SannDLeNrd9Zm0femNBe9ojbN4m0rR36QzP1N+juFlfC/lcWRc1xYXYapdyTA5oa4t96vAYXA9QaHxyOjJ2C8EOA3Thvpo9WcxE2U8"
                    "ebi2uDXtZl7tanKWFzhv2hJM9hGjuOVrJSCCGlrg4985i9Irre4A4KyvqD5fw3BsIxDCWqF3I6KZyLQ3mYst5FktLikBuBCwo8sx6dIK2vYPvoLJ47SI5Puo"
                    "fdW2ht3p8Y+p7045fyrC8SwTc9YvZq0alaxNjmVKkcghmmHuOlse9p3t+2AITpzgd9tnu8i+mfnvFeMZvlGanwkNXDVXXJa1W7LbszRNexri0trRRMLQ8uIM"
                    "jiQDoHwoxfqjPxoiPzr8fJkrReP58Ef3fjwZbMSJD9HfzogA/wCY8fkg/gD435Uev042B4/P4/177H7jSv6+hkv3PTB6dj+CK456TrZmZa6jOWT8kZEZGZma"
                    "tefBkez2KffqdPV6485d8GfDEHsQR/Z7EgEa7a0PHbR2D3GzbR9OTSz0Z4Sw9y2rkgfj/wCjWSPg9x2Pz/y8Kv3+o7e9rr3wMteP+kHiQ9/g/wDzm6ndkevJ"
                    "dpF3eNbLezLtEuPold0+nfLPyea/HYj/AKixutfne9a8/jyotfWVH1824j47cVs9v/4vN5P4/wA/jwdqu5nMzuwPOD35LDcqMzM/j/8AIbDZd38+NFrfcov7"
                    "S2RnLDkbx/Z/OnX8uGyp7d9EUZ+/xsAg/g9tjR2FGLiUBby7i2u5/tFg9dvP/WlQeCQNbOh89/8AJesmaUqSaVJJSVEaVJMiNJpMtGkyPwZGXgyMtGXgUOAk"
                    "HYOiO4I8g/lXX62NEAgjRHkfuP3CrL9dH6e7Gc+tck5M6J8lx3jK9t5jlrL4GzhdnE4rdsLC1KXbrwDMKuJfXXGda2xIsJVVgjmMZVhcR5Fbi2IlxjiEaPHr"
                    "5iek/wBWud4jSp8f5tRm5RhqjGV6mTisCLO0qzAGMie+Zr4slHFGOiMTuinDGRs9/W3KKXqh9LfH+X3bWe4ncj4vmbcj7Fyp+n93C3rLwXPnNeN8b6ViaTTp"
                    "pax9uSSSaeWGWRwBrw8w+m5148FSLouQOlzl86yjOU/JybDMYXyliTtLEKQ6eQO5Lxc/mVPT1P00Vcx9zJpVJProrrJ28KBIN2KxMvjX1D+kfJ2xtqcupY6x"
                    "KQ2OjnurD2i8np6AbINV7i4gNay08vB21pGyIi8i+nr1T44C6xxifJwMDjLbwMgykEYADgSyMR3dFpJLhU6GFpaXA9PVo21NTIadWw4xJRGd9t76Z5t8ozqV"
                    "dqmpBsqX7LiFEXe26RLIz0ZF+dxU8jTyMDbWOtVr9Z4Op6diK3CSDp395C98Y/B+4dJI7/jUd/D3sXZNXJ07uNtaB/T36s9OZzSNtLYrTYpCHNJILWOBHfv3"
                    "1z2PZhkOK5BU5TjF7eYzlWPqkOY/lGOXVrj+U46/LiuwpMvHclppMG8oJr0N16KqdTz4UwmHXWUvJbdWR9fMYrFchpT4zPY6nmMdZjMctPI14rcJY4AHpbM1"
                    "xic0hrmSwuZIx4D2Oa4Aj6YjI5fjtyPJYDKXcNeikZK2zjrElR7nxv6mCZsREdiPqOnw2WywStcWyse1xBtx+jj6wmc8553S9JXVHYqyXPr6FN/0e5eWzVwp"
                    "2SvY7QybaywTkFiG1Xx5eRHTVFhb4xmESH7+Roiz6bJ0qyVquustrj+pD6dqHB6ruccIbIzjbpWMzGGkMksmHnszlkVmnKeovx0skjYjC8g0iGNDpInBzbCf"
                    "p79fLfNZW8N5k+M8ljhfJi8owMijzNWvH98diPs1mShYzrkewltzqdLHHGQ+OODL1kXFx/VK600L2Sl5rw84fefnsc6TOnY2jT5MzR7Pt+ftMlEadeCNcmfp"
                    "Lljf6M4hrdkxZnOtcdkakNxrunQ3saIP52d634jt9V0Dv+lqd7x/Nx3ChnfZLBJkdFw1rfV1aHyNEnZ0I68bo8jzG9qsWw/Hb/L8pv58esocYxWjtclyO6s5"
                    "jyWoldTUVJDsLe1nSHFJQxCgQpMh0zIkNHpQkPk8vj8LQuZXK3K+OxlGL3rt65K2vVqwh7WGWaZ5DI4wXDbnHsSBvWiI9YzC5HM36mLxNOxkMldkMNOlVjdN"
                    "YsShj3mOKJu3Od0Mc4gd9NO9nxsCroq61t+ejLrAUZ+DL/pY5/MjMiMtmo+OdfBa2R63oy8aGvj62ekgI16i8OI1rZztLY7nYAD/AMu321sb0SQdZ9/0K+q+"
                    "v/rA5QAe4H8Nl/r3AcSD52D332I+V/H/AEXdaySIj6LesBXxok9K3P56393kj44T8fB+S8/Ktj9HrZ6RnQ/6ReGje/Obp/6E9fk/Pnfjv5Q+ivqsdn+wPKu/"
                    "b/52Tf08Bx8geda3+Pj/AFPRb1r9xmnov6wCVv8A5dK3P5F5Pf8Ad/p1o9GZ+fPxvZlsfn/Tb6R70PUbh+g7W/41U76+Bs+N6+7ejrsfCH0U9Vi3R4ByggAa"
                    "3jpfPz22O+iRr+vnsv8AP+i7rX3s+jHrCPyXz0rdQJn4PRGZlxxrReTPZ+dEeiI0mfE+tvpJ1Ef9IvDx47jOViBvW9EE7OgPkeSCeyD0T9VuwHp/ybZJ3/1c"
                    "9v7724gD4Hn48Dsrh/oFcecq8Y9Gmd4vy5xjyJxTfJ6icusKrH+TsHy3j7I51DL424kQ3ct4/mlJQ237Y/cMXMCNORDXDlP10omn1OtPttVx/VJyDj/J/VA5"
                    "fjWYx+bx03HcRE67jbEVmt+pgfcZLF7sTnD3GN9tz2vDXt6wCC3pc6wr6bMBnOM+mdfEchxdzEZGHN5iU070ToZ2wTzRyRPMbgAGPBJYWlzSO+wSWt776h/o"
                    "78J9dF09ypR37/CvUA7ErINryDV068ixzPa6kguV9TE5BwsriiZsreBAKHVV2b1VnV5VGpqympLWTkWOUFJQQfK9IPX/AJX6TO/QQxsznF5ZnzTYG3KYhDLM"
                    "5hnsY62I5ZKkrw0udH0PrySHrki6nPc70/Vb0O4t6pRi3ZLsPyOGEQ1s9TiY+Z8cYf7Ve9C4tZcrMc7sHuEsTC72ZGO6XNqtc3ej36iPBM1bVnwBb8o0iI7j"
                    "ycv4Amq5dpJMpHvrdhwscq4FTy2lxDDSHkyLfi+ngyFPsxIMqXNJ2KieXFPqp9JeSRQtuZWxxa89rRLVz8JirMe4uAa3J1xNWe3sSZHiFrBr3C0ubuD/ACv6"
                    "W/U7ASTPx1Grymkx59mbDzNZckjaAeuTH2nRvids69uOzZc4nbNta9zY3sjpsiwvJbLDMzoLvD8yp3lRrnDstp7LGcup5CUIdONaYxfR6++rJHsuMvmxPgRn"
                    "faebd7CbWlQ3th+TYHkMTZ8HmcVl4nsa9j8bkatwFrtdLy2B73hpDgQHNDtDei0jejsvxLkOAc6PO4PLYgseWON/H2a0ZkB09omlibC9zSNO6JHAO7edrhHJ"
                    "LciO5CmMtSYT56ehSkNSIr3b57XoshDkd4tq8e42ot72n5I/alLZY3V5WMkgcCHwSsbLDIN92yRPBjeCOxa9jvkeN68aKKSvPHcrTSQWYu8VqBz4LEbu33RT"
                    "xOZNGQR/Mx4OwTvelMn0A+sx1DdJ+V0mMcpZLk/OXTxPuYacrxXMruyyPM8DpZa4MK1yTijJbFm2yH38drYqbGDxRLmPYRf/AEkqnx9jBLu/mZWUUvWb6YuK"
                    "8xxt/N8MoVeN8tghlsxV6ETYMTmnxtfIalijF0w1bMziWR260bHF7me5HMxjY3ym9HvqS5NxXIUMJzbIWeQ8XsSxVP1t14lyuGD3Mijsi7IRJapxN06eO3JI"
                    "9jGukEzdvc29tjuQ0mXY/RZXjNpDu8byemrMhx+6r3SfgW9JdQmLKptIL6fteh2ECTHlxnU+HGXULLwoVbTwS1p5q1iN0U9eWSCaJ40+OWJ5jkjcPhzHtc1w"
                    "+CCrKopY54o54XtkimjZLFI07a+ORoex7T8tc0hwP4K81Drcf9vrP6u0fkuqTqDLSVdu/wDzdy89bIjUZEnZa8aIz19p7K7P0jcP+i3068HXBuK7B123g6P+"
                    "Q76JJ2dfjWlTh6rQk+p3qEfG+Z8lI/f/AK4uHffXnxsE+PGtLcX0QH/c9TfpwT+NcvmkiPuIjLgHlff5+35V/OyL+PI1B9X7v/Y3YA7/APrPgNn/ACv62dd/"
                    "8P4Pcfkhbd+kqLo9V3OOhriubHfyd2sQAB30T2+e+gvQDFUys2Xk4cfTP/L/AADev/sbEt/JEZ/0/X9xdpa8EZ7L+0u0u3xsyF7/ABKUP4txh3ku47gz20D9"
                    "2Mqu1sdgQTr+p2POhSlzis7+23NNgg/2t5J2Oz4zN3uDre9dyNb/AGXaPrCJBrUaUEkjUpRqMklojM97T40fj4VvuIyM1dpD2Jp44opJppGxwQsfLLJK7ojj"
                    "jY0ufJJIdBjGsDnucdBoHfsO2PRUpZ5YoIWOmmnkZDDFHGXySyyODWRxsGy97iWhrB3JIA2db71yLx9yDw9mNpx/yphWR8e5zTRqCbcYjl0H9qyCpj5RjdRl"
                    "tA5ZVi3XXoD1hjl9U2J18tTNpVuSnKi7gVV9BtKqF4PF+Xcf5njP41xnI18tjBbuUf1ddxLDPSmdDMACAQ1zmiSJ/SGzQyRzN+1w17nKOGZ/huRZhuSY+TGZ"
                    "J1StebXk7k1rcfXC4PA0XtcJIbDRsRTxSxlzulrnSLekx16I6H+paFaZhaJruBuWWqvCucVLZhqRU1le9aO4VyU6+8ymchvjG3u7eZcNxp7bS8FyHN3k1N5f"
                    "RMZis6P+pn0nHqPw5+WxcAfyvicFm5jmtJ67+NDRNksboP6XTFsQsU+po/vWOjDh7zWrdP02epr/AE/5eMJk7DhxnllirUstIJZjsu4+xQvtaGlzYp3SNqW+"
                    "hp7GtK7ojhlefQRjSY02NHmQ5DEuJLYakxZUZ1t+NJjPtpdYkR32lLaeYeaWlxp1tSm3G1JWhRpMjOpdzXMc5j2lrmktc1wLXNc06LXA6IIIIIIBBGirSAQ4"
                    "BzSHNcAWuBBBBGwQR2II7gjsQtQ+unrIwToX6d8q5zzOGeQWUd5nGuOMDZsUVMzkTki3izX8exRmzXFnnVV5s19he5ReN1tq/j+IUmQXcSmvJlfGprDOPTjg"
                    "WW9SuW4zimILYZLrzJcvSML4MbjodOuX5mtLS9sEZ1HEHM96Z8UPXGHmRuIc85rivT/i+S5Rly59ejGBDWjIE165KeirShLtta+xKQ3rd9sbep5B6ek0GerL"
                    "rm6jutPK5eR88Z5Mu6ZUyPJoeLqVywqOIMMKDIlyqteMYA5YWFWm2gnOeR/Wl87fcgT2FNRLLKpdbDrYEC2j029GOCemNCKLDYqvey5jAt8iydeGxlbT9guM"
                    "cj2uFGIuaOmvUELWkbcXO7qrH1G9ZOd+pN6V+RytnF4UPeauAxVmatQhjcDGBaMbo35CTocWvfa6ozs+3GzZLsCYHg3JPMGXQMD4ywjMuSs1tI7suFimD45f"
                    "Zhkb0CMskSrQ6aig2NgzUQlKR+4XDsZurrkaXOmxmkqWnPOTcx47xKn/ABDlGbx+HqBvU2bJ3I4nP07oAijcXTSu6tRsETHlzy2NoLnNBwPjPCORcrtGnxjB"
                    "3srY6mteKFUujjc8dWpp9Mgh+wOkPuysPQ1zwCGuIkS479F71Ls+XDde6a52BU9m2hce+5A5H4joozTK0rX7thjkPPbTkKqWk0pR9FY4TGnoWtJPRWUpWtOh"
                    "sh9Wno7R91keUy2Rkid9raGGtOjnBBO457X6aP7u3aQsI6iHaLekb0pfSf6r3BC+WrhKDH7Lhcy0ZmhcD/LNHWisgkHe3QyzN7HoMjSHrAnWP0DdR3QrJ4/a"
                    "5/g4PEa5RTmK8OVhuXOZQ48jBF4m1kDlo0qmrEViVLzaiKsSiRMcmp/cDfbhnFbRKzj0u9auLerU+Zi43Bk4XYOPHy2nZKvFX625E3RGI2xyzffH+icJPvc3"
                    "+8b0u7HeG+pnovyL0sq4WfO2sdZOYlvQwtx8ksjY30G1nuJdIyMuZI2yOkuZG8Oj7sPVsaQTJnt19kvxtNZYqLe/koMj8aL/ALGZa14MyPQ2pkiP4dkdu0Dj"
                    "7wPYdv8AdJvne+w6u2zvv8edYYqE/wAWxRA3rK44+Dv/AObYD4Hx89ye34+PVrFB6vHQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQEQ"
                    "EUbvqy9L6Oqzob5gw+rpI1zyJgFYrmTiZK4xyLBOdceQrCe9T0Z+40iPb59hcnMOL2JjxqYhM5xIluNqNlJp2l6Mc3k9P/UfjfIPdfHR/WNx2XY3epcTkiKt"
                    "xr9bd0xB7LQLAXh8DdNf3Y7XXqvw6HnfAeQ8eexrrMtN1vGvd/7nKUP96oyjuAf76MRua4ta5kjmlzN9TfOeOcyS0nHktyGTS24xJjuEtl9p1CVsvtLI1dzL"
                    "qFpW0syJK0KSoj0sjK6KKcOa18cjZWPAex7CXMlY4bbIwjs5rmnqYRrbO4JG9VDvpPBc2SF8T2veyWKYBssUrHFskMrdfbJE9rmPGgQ9rh8KTXDfUCl476X3"
                    "LHQv9flacgyXqBwrKaSexcy4UCs4aumJWc5jjFOplLr7UD/VzjKBIy+iQ/Dh5FWc1XaZHvsrt4llG3L+kEGQ+onCeoYhZ/B4ePyZS81kLWt/tLjenHUTLJ0t"
                    "je+zHPHbeJOqXVKQEdEkZjkVivVGej9P2Z4KJWDLyZn+DU3SS9b3YDKF967IyIuL2Gq1k9OKRoaxjrEDhqRj3PjERO9xXalXk9FpZmkjIz8GZmrRaJRGfwXy"
                    "Rb87kkJQS0dRGyNHZA2e3nfcbPfv48/KjwanSN9I+e2tk62SB+/bz47fjWvQF9Djpra6f+gfjrJLKHWJzXqKec56yCxgtyFuv41mMOGjievekzmI81sofFsP"
                    "FrOfUJbRW1mW3uWOVyX/AK+TZWNPn1Ec1fzb1T5FYisPnxeGsPwOKb1H2WwY57oZ5oo+pzWus2xNLI8Eul+xzuloZHHax6EcNbwv02wNSau2DKZOu3M5YkN9"
                    "428gBO2GSRoBe2pC9leFp7Rxt6R1OLnv/wA9d7McjxH01OZWcbu7TH3swyPi7CrmfTzpVdNk4tc8hUD+T485JiOtLXT5jTQJeH5PXu+7EusWvrqkmMOxrF1B"
                    "/v03YmnmfWbhla9GyWCvZu5JscjWvY6xi8bbv1SWvBa7psV43j5BaHNIe1pXP15ylvEek3MLVJxbNLRgx7nDr6mwZS9WxtlzSxzXNe2C1IWO30hwHWCzqB8/"
                    "Y55qP73DNRmZ/caj1vzsjUZ/jRn53ojPWiIyuBEmuokkb8n/AD0fwPzoaIG+2ge9VBqAfysAAGtAeAANeB4Hf4A/PlXN/wBOd00cPN9PeUdU03HKfIuZr7kz"
                    "KcGqsotYcWfZcf4hjlPjaDocYVIQ8qglZJNnzbvILGAcSfeVcrHa6wU5CpYSTrK+r3mvIbnPncMdctVuPYfHY6xHj2TyMrXbl2AWJb1iFpbHKWhzIYA5rxGI"
                    "nODup5DbE/pa4hg8fwWPlMdatNncxdvxWL5ijdYrVqVl9WGjHIQZI2t6HyyNBaHPlJDek9T7K4iCpRqHD13+XMO409Ork3HMink1kPMGVca4BgNUhuQ49b3t"
                    "fnlByFcGtcdtz6KHVYbhGSWj06X7UJUyNXU5vlPua5iRv36ZcHfzPrJxOxTjc6vgrE+byMw1016lWtLG0uJIG5rU1eu1oJcfdLg1wY4LSv1CZinivSflUNmR"
                    "rZ8zTbhqERI65rV2Vjfsb5PswtlndoEAR9yAdigAqeWj+5RaIz2ozJRJLZbMy8mX4P48aUeyLRW6iQfB/Hb+g7dvJ15Pny7Q2qtP0h2PsGu29fg//LsvQQ9B"
                    "903vS76elbMyK750JJmZH4Vz/wAnufBGfae3D2k/O9n+SM6ifqdH/tu5ofO34Y7/ADvAYvv++x/p4+FaX9PILfSHiTPHRHk2gfjWXv8AbXwN7/8AH5VfT9SZ"
                    "INnr74/LuVo+j7iMzLf2FvmnqdLZFstKPXkyMjMkkXn8Sx+il3T6e8t2SP8A1yGvx/8AOPHg7P8AQjt8n8AFRm+r6L3OacT7A64xYHfz/wDPaXQ8eO/+X7De"
                    "66WaTEnhWbJMzI14hlBERbMlEdFYpSWiNRH5UkyLyRfktmZlKvkcm8BnQD5wmWA2fJ/QzjuRoAH+g2Bsn4EbOL1yOU8Yd0j7eQ4QknROv4nTc7uex7A9zod/"
                    "J7L0yuvjr+4v9PLjnB+TuWcE5SzfH8+z8uNqlvi+vw2fMg5K7ieT5nFTb/1nm2ExGIUqlw++Wy5ClWEo3YSyOF7e3Cpm9OfTnN+p2dl49x+zjK+RjpTXw3J2"
                    "Ja0csFd0bZxE+KCfqkjbIHlrgwFoOndWgbY+c85xHp9hhns5Fefj/wBTDUe+jXFh8U1glsHuML2EMkeBGHDenEAjuoY8j/U78TMuKPCukzky5YJSiT/WHJOH"
                    "YhMc/wDQlMfH6rkOOTqy8dq5yG0r8G8afvEiqv0Xc6lbuzybjdd4aSYohdsvLgP+zjPtQxue4/a0ufGwnXU9gOxomx9WfCI3FsGD5FYG9CQxVYYwN/zyF85k"
                    "aweT7ccr9b0xxGlZD4h5Tw3nHivjrmTjyyO3wXlLCsaz3E7FTL0V6RQ5TUxLmu+rhyUNS4E9qPLQxYV0xlibXTm5EGawxKjvNIiLlcZbw2TyGJvxGG7jblij"
                    "ajII6J60roZANgbaXMJa7w5pDh2IUnqF2vkqNTIVHiSrdrQ2oJAQQ6KeNsjDtpI30uAIBOjsfC6bzR0vdOXUXAjV3OvB3FvK7deqQ5UTM3wqhvLnH5MplMeR"
                    "OxnIJcJd7jNk4whDJ2dBY1tgTSUoTJSkiId7Cco5Hxqw21x/O5bDTtcx3uY2/Zpkuj30F7YZGNk6Op3SHtcGknQ7ldPLYDB56u+rm8RjcrXe1zHQ5ClXtMLX"
                    "6626mjfoP0OoDXVob3pUa/Wb6HOIehfqLxPHOEba3bwPk3BP66i4PkF47kdvgU48oyOnfrIlzPWq9n4fLjwIiMXdyaVc5D9VW5G1ZZJapZjlGs8+mD1R5L6j"
                    "cVzLeUu/WX+OX6lNmXEXtOyFe3Wkmb+qc0+3JdhkgcZHRsYHRSxktHYGun6jvTXjvBOSYeXjMLKVPP0rlmXFMeHR0Z6c8ERfViLS6KrOyfQYXFolid0629y0"
                    "E6SsgsKXqg6cLuoYkTLWk594bu6uFFdQ1Im2lLyVi9nXQWHnVoQhyfOix4Jd6yQaZCiWtKVKUW4fU+nVyXpxzylccxleXiOedJLJo+17GPsWI5CdE/ZLCx/g"
                    "76dAHZJ1N6bzXMZ6icJu02vdYj5PhoxGze5Y7NyKrPENH/HXnmbskdO+53oqa/8AUkcO2WJdV3FPNULHkwsW5b4ci4/MvoLTXt2nIfF2SXMbIHbpxDqnWbNe"
                    "EZlxfXVRym2U2VbQy0wFSf2KyKLF76K+UMs8c5TxGaYus4rJxZmnET3Zj8lDHWsCMdI2xl6uHkAuLZLJ6+hr29ckfq84wYs1xrlkUYbDfoz4W7JrYfbpvNun"
                    "v7jpzqslhvdrB0wN6S/7+iAri3lHIuIeTOO+V8OVBRlvF2cYlyHi5WjTz9U9kmEX9fk9JHt2Iz8eRJpZFpVQ2baKy+wuVXLksJdbU4Sylxy/Bw8s4ryHjUzv"
                    "bjzuIyGLLyddBuVpIY3kgfb7crmP3vbentslRa4nlZeLcnwHJIme4/C5elkegBpdJFXsMfMxof8AaXSQ+7G0u7AuBJGhv04OnTn/AI36ouF8B514otTtMLz+"
                    "kZtIbb/sotaOwQpUa7xbIYrD0huvyTGLdmZR30BLzyI1lBkJYfkRjZkO0mcn43l+I57J8cztV9TJ4q1LVsxPa4NcWOIZPC5zWmSvOzplglAAfG9ruxJAuCwG"
                    "dxnJcPQzmHsx28dkq0dmvKwtP2yNDjHI1pPRNESWSxk9THtLT4WbB4K9hax9XPVvw/0U8OzObOap1q1jLWQUWK1VPjkWvn5Tk+R5A+79LTY3XWlpSwZ06LVQ"
                    "7nJLBt+0iJi47QXdj3r+i9lzMODcF5D6ichg4zxmvFPkp69m051mYVqletUiMks9qw5rmwxdRjha8tIM80Mfbr2MW5jzHBcFwc/IeQ2JIMfBNXr/ANxC6xZm"
                    "nsyCOKGvAzT5pSOuQsb93tRyOAJboxiM/qIPT8dbJZsc4tOGZEUdzAscN81H26TpnPHWtmau0j93tM0qV3dhEo92/wDoi+s22huP4+7qOh08kx3netd3DZ+d"
                    "N2dEfvrUX/pQekuj1ZHNt0N6dx3Kd+wOu0BG++iCRpwIP77x5L6h3DWN9DzfX4vD+W7HhqRFx+fCx+JjWPQOSZ1bk/JdbxbTWsehyDL6WiOrl3FrDvY81eUI"
                    "ZmYk6i7rzmG/FiP6lo+mfIb/AKhM9M4p8Qzkbr02OL33+rFtuV6sluaE3oIZg5zGxPh2yNzTYHtA9+pbOu+oGCo8Gk9QZYso7j8dCPJkMoPGQNKWVkTJv0Mj"
                    "45GtPW2VweWuZCTI5oAIERuS/qb+nqMgyw3pe56uH+37E5fkHFOIsKc2WyU9j2X8hvJbJJmfufSqV3ERe1ozUW/Kv0Zeosm/1eb4zVAJG47Fmzs/Gv7iAkE/"
                    "I8dj8jelZ/qw9Pma/T4/kFntsg02wOH5BEjunYPYAOIP5B7KaXoU6w8P66enDFOoLEaVWJLuLbKMbyjBJV/X5HbYPlGK3syqlUtvYVzENtTtjVt1OWVBvwK6"
                    "VKxjJKKycgx0zm0nHX1C4NlPTnleS4nlpI57WP8A07m2oWPZXtwWa8diKxXEn3GMiQxk7I643gE6W9+FcvxvOuN4/kuKEkdS+JdQTFhnryQzPhkhnDC5rZWl"
                    "m3N3sBw2tkM5474/5Px2yw/kvBcO5DxK4jKh2+L5zjNJluO2sRa0rXFsqS/g2FZOjLWhC1MSorrSloSo0maSMsVp372OnjtY+7boWYndUVinZmqzxu0W9Uc0"
                    "D2SMdokba4HRI8FZHZp1LsToLlWvbgeNPhswRTxPAIcA6OVr2OAcA7RB7gHyFVv9b70wulPgvp1ldU3AWJQOF8ox7N8MxrJMExuxmRePcyosxsjx5pvG8JkH"
                    "LrMSyTG7CVXXcdvDl45j0jFIGVs2tDaWiqOxp5pfTN618+yfN6HCuQ5e7yTC5KldbC/JSPt3sXJUrPsQSwXHuMv6UmMxTxS+6CJGlhaWNUSPqH9H+D1eG5Hl"
                    "2ExNHAZrHWa0r3Y+FlWrkmWJ2w2IrVWJrYpLD2vEkM7WtkbLGA5xZJKH1P2J5odT2uKSpv8AtPZ+FJIzI/J/JGR62ZGfjWzIyOxcSaOyd61rej3BHf56iT38"
                    "kd9b3pQAkqhzSC0aPY70Cf27DY8jt577Pfz6Fnol5lk2cemR00WeV2T9tOpW+WsBqpchaFuN4fxlzlyXx1glUk0JSkmMewzF6LHYqTI1pi1TJOKWvuWqmX12"
                    "q1aXq96gQU444oRyK7J7cQ6WNlmcJZ9D5c+Z75Hu/wAUj3E6J0rdPR2exZ9LeBzWnSPnfxfEFz5CS9w/RxdJ2e4aG6EY8NjDWgaCozdc0wm+tXrAbNRl29U/"
                    "UQRkZ/H/AJv5j5LWvGkkWiM/t0k9EnYtV9JHhvpf6dNJO/7EcW/J7HDUj3+QAN6/A1o67CtH1RrdfqRz5waO/L+RHwP/AL7Wx22e+z3/APwt+ey3N9DCST3q"
                    "gdNhGpR7b5i/5Ge9dP8Ay0Rb/Bl87I/hWj0RlotRfV04O9HbAHYDk+A7AO1si/4321/Q6+Nb0tq/SxB7XqlvpA3xjMgHtsD9Tij8f01+fj8lehCKrVZGvJB4"
                    "+na4+4+2pW04Lh/cSjUakmWP1/nyZa7TSZGo9/yrZ/33p8TkB4txkg6D+PYMg9/5f4bVI3+djsR21sDY8CnDmtXq5pzJwbsnlnIiT2+cvc34/J8/sVO/+n+4"
                    "1445Z9QiC3yNh1LmjPHPC2c8tYNGv2Ez6+g5KxLO+JKrF8xYrXTVDlXWNxspuZePvTmZTVLeLr8krGo2Q0lLaV8cfq9zuVxfppTr4zIWaUWYz8ONybaz3Rut"
                    "491C7NJTkkGn+xNIxnvMaQJY2mGUuikfG7e30rYPG3Of3rF+jXtTYzCuuY987A/9LbNqGNtmJrvt95kZeGSFpdG5wewMkAc2bn1/+gxHM3DLPV/xxTE9yjwD"
                    "QOQ+TI0Np9yZl/AcaRYW02YmHCqp0qxuuIrmxnZhBNc2shRcCueVFPlbW/8ATMNiMX0t+rH9h+Xf2XzNv2uL8smjgc6Un2sdm9COhe31gRxTn/c7JDHk+5C/"
                    "TPbLxIr6jvS4c44qM7iqwk5LxdktqsGFrX38WR1ZDHnqHS5/Q1tqv1PjAmrtaXdLi00nU2LkV1O1raU04lRdqlEptST2RkaT2lSfP4MiUX50aStKc/pJDhp2"
                    "iCHDvsA9nN1oEeC1w89j3CrVFVsrNhoc17f5u33Bw/OvGt6I8j9tq6R+n36/z5l4smdGfJFumTyLwTjzVnxHYSfdS/k3BMV2vqWcaekyrSa5PvOI7ebFoWCj"
                    "QqiCzxpdcc1sWPaWmPZVcv1d/VN6TjhvJxzDC12s43yyxNJLFDGWxYzNtDX3YD0tDI4brnm3Bs/9o+dgJLdCyX6b/U2Tl3Gf7NZqx18i4zFFB7s0oM2TxRBb"
                    "TtgOPU+WANNWyWgjqiZL0sZKwHUr9TfmeQq5B6VMEK3sWMVq8J5LzEqCPPls1Nrkl3e4vRJvLWsS8UGxtMbqaSRW47PcYOZTQ8wyuLGeQxkEptzYf0SY6k6z"
                    "z3MvY12Qr18NjoHkNLo6tuW3Yse24guYXyVIQ4tIDgGh+/tIwH6w7tv9DwfDtJFG3dy+QnAB+6zj69aCu1x30ke3fsPaxzSS6MOYR0O3VuKx7Ve4fuPdp9yk"
                    "oURurSjyTaFH3I7lkXYkzQadmXyW9WAGbpb1ac8NaX9DR9zunZ6QCPLgA0dtk+N6IMFxTBcGbbGHEN63DTGk9upzRo9LdhxA0dDydr0svT/6b+Gemvpb4ixr"
                    "h7G6OC3k3H+E5hmWawocM8h5Qy29xissLPOMsvWmky72dZOSDKsTIeXAoKFusxrHYtXjdRVVcOkT1C5ZyDmPLM1luRXrVu2cjciihnlldFQgisSRx0qsLz01"
                    "4YGNbGI2NZstLngvLibjeEcZwfE+NYjE4GnWq04qNUl8EcbHW5HwsfJankjG55Z3udI6VznlxdvqI7rdQYSstVOL9TDzNhmR80dNnCtDZqm5nxJgvJGVZ7Hj"
                    "pS5Ep2OYLPjxOG1Lsxt1TaL76LjC4u7WlkNty4FFeYhcOJ+kyGudcsG+ifA5CtQ5tyOeF8eNyU2IxdGR7ZGtsz44ZCa4+Mua1kjIP1kEXVG55ZKZGPDCB1Qa"
                    "+sDLULNjh2BjkZLkKIymStRNfG51eG4ynBWErQ4yROn9iZ7OtrQ9jNsL9P6axkqSa4c1s9q74M1s07I0n3RHUmR/cZdqjMy+4lJLzvu0Jw5B4/huRaTonH3h"
                    "3H5qzdvg9/B1ojeuwUN8bARksYdeMjRPbWxq3Ee37jz2+O/buvWQL4LXx+BQwrrUBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBEBE+Q"
                    "Rea96qXTCvpD64ubeLoESNEwq6uP9U+Ko9dQIxqpicYclSJ1xjePUlXGT+3IqMCtmMl4tr369w4s0sBdl/S1zjrlVCt3+nnng5t6YYOaxP7mVwbP7PZTrlEs"
                    "zpMexralmV2+vdqi6F7esdbnNk2Xke4+sT114S3iXqJlW1oBDjM8452h7bXNiabkjjkIWA7HVFdMj5A3+7aLEYYG76Gx2pmu67C2RedF2n47i1pRGWyI+1Jl"
                    "s+0vktmStbw90/b32AN9vxsEjfnZ/fvrt2HjT/6aPe/l2u++50Dr5+Opw7fn5GtbXdDHT5K6s+rTgrp7Ql1NdyLncCNlj6Gn1/R8fY6zMzLkyUiQ2xJZhTP9"
                    "PscyVihkz2/onMlepYEhWprZK1d6y82bwX055Jno5Q27+jONxQbol+SyH+7Vydn7WwNfJYcfgxAhriWhbH9KeGnmPPuO4h8fu0I7f8Ty2z0tGPx7TNIx2gev"
                    "35hBX6BolsrndTWscV6gLLLUdpqPHabYYYbQyyyyhLTTLTSSQ2002gkobbbQlKEIQkkoSRJSREREKanOc5xc4lznEuc5xJc5xOyST3JJ7knuT3KtTADQGtAD"
                    "QAAAAAABoAAdgAOwA7ALRX1L+lef1l9E3OXBNA6+zm9tRVeYcbKjyYEJUzknjLIqjkTCaCVPtDKvr6jMb7GImGZDYyDR9Dj2Q2spl1iQ0083nHpry13BudcZ"
                    "5R9xhxeUgkuNbvqdQld7F0N01xLv00kjmtaNvcAzYDisU51xtnL+I8g444ta/KY2zXge8dTY7RjLq0hHU3+Sdsbtk/aQHaJaAvNLu662x+3tKG6q7GluqK1t"
                    "KW8praK7Ct6W7pJz9Xc01rCcL3INvTWcWVXW1c72vwLCNIiPpQ8y4kroMblaWXoU8pjLMVzH368dqnagfHLFLXmZ1xuDo3uaXEdngnqZICxwDmnpqcvYq5i7"
                    "93F5Ku+pkKE761qtK0sfHNGSHAh4a4xu7PikLQ2WNzZG/a4LffoQ9TTqT9P27yFfEVpj1/guZqal5jxbyBWzbjCba7jNR4ULLa9qqtseu8fzCJVxm6dVpU3s"
                    "SFcVaIUPKajIW8fxf9j1N6q+iXD/AFZbVs5Z1vFZujCa9bNYv2nWHV3O62QW69hr4rcET+oxMDoZIw97Y5m9ZI2d6bervJ/TJtipjYq2Uw9qZtibE33SsjZM"
                    "Q5sktSxDt9V8oIdKeidjnsY4x7L+uVi9/U7dTcuuOPjfTpwPQW/tqT+63Nxn+XV6nTjoSlaaOHbYTIZSmT3uqbVfy9tKQwThqbVIc0RS+i/jzbLTkOa5iSqD"
                    "t0dTHUoJ3tDtkGaaWeOIuYWtLhDKGHbg14IaNyXPqyy5ru/QcPotsOBDX3L9l0ETiNAlkMTHzBrgSW+5AXAhpe3RcYQ+oLqR6nOtrkC15L5kzDLOV8joaG6t"
                    "ksMV6W8Z48wiC/Dk26cexmhhRsfwrE4JuVTVvZMRI7tpJaqZeV215fOtWL0luI8N4B6TY6vjMFDSwoy1qGobV62x2SzV8kiGJ9qy4SWZtuJiqwNbHGzxEAC4"
                    "6A5Ryrm3qjenyGYfayjcVVltClQryMx2IptafdsCtD1NjDg0l9mw6R73EtZIGOEQ1a98yToyLfyZ/cX/AL9uteT7d/8Ax/Otk9Xfxs/jsf37jfYEn/X5Hk4A"
                    "IgNHv4130Nj47nXb8/67CvT/AKbfkS0yvoTzLC7OWiQxxZz9l9HjbCEMpOFjWU4lg2eLYcU2ROuuO5nkmay++SXuJakNMtq9hppDdXH1dYpuP9Wn3WMDWZrj"
                    "+HvkgO++WFkuPled/bv/AHNrQGdulrSR1Ek2LfTLkv13pnFVLi6TE5rK0nA601kkzbsQbruWlloHqOyXF43sECG79S86pHX9x0STPz0ecTmej/t1zV1Nkkz8"
                    "HruPZEf2kej2Zkn7ZCfRc/p9PuWdwNcyB7gbP/UmPHYn8DZI8f56WkvqxjD+Y8XJH/2szgHtr/56zEj/AD7ed+BrXcquDlz6lYdl6T8krFMiI0mRKVo6WaXc"
                    "ZePuM/P2n3FsvO9Gcqs68uweZA+4/wAIyYADvk0pwNdI0O+tdtH4B0VHfjcLW8l467R7Z3FHyRvWQrEjwNDXcH40SANL0/fUT6P4XXH0m8kcDIn11Jl85FZl"
                    "nF+TWjCXYePck4jLTaY47NkphWE6upMhQmfhOXWFRFduW8KynJGa1Lkl9DTlNvplzm36c82wfLKrXytx9gsvVmP6P1mNstMF6qfg+5Xe4sDvt91sZ20gObaV"
                    "z3iNTnPE8xxm2WMGQrEVp3sEn6W9CRLStNaSPugsMY8EHwCCHNJafNb5EwXO+KM5yjjjkfFb3CM8wi4lY/luKX8Uo1tQXUBRFIiSybcfivtKSbcqvs62RMpb"
                    "usfiXVFZWVJPgWMq4/j3JcNyvD0s/gMhBfxmQiZLXnieC5hIHVBOwHqhs13O6JopNPY5pHcdLjVbneNZXjWTuYPN05KeQpSvjkikYQJIw8tZPA/RE1edoD4p"
                    "W7Dg7Tg2QOYN/wDoi9Wnq66Fa97EeNcnx/MeKJMiwnHxLypTy8lw+rt7SREk2V9isqpt8ZzDF7OUqM8t6uq8qRhkidZ2l5Z4fY5BPetFam9T/p/4B6m2n5e6"
                    "y1guQyRsZLmsN7fVbMXUIzkKM26tiQh7RJYHtWXsjY33/tcTs7099beZenlWPE1hWzWCie50WLyXutNVrw3bKV2LcsETS0vjgdHLC1z3aaAQGyL2n6m3q8kV"
                    "y2afg7puq7U2lITYTYPJd3ES8feROJqmeRKR3tSRoPsVauaNCjUa0q7UacqfRdxdk7Te5rnp4A4ExVsfQryPaO5b70jrLWE6I6/acBvsN63tS19WOddCW0uH"
                    "4yKwWkCSxkbc0LD3Af0R14XPA7ENL2g/4nNBUDXO3UHzD1Mcm33MPOWeW/IvIuRtV8WxyG1jVcBpmuqYxRKmlp6Shr6nHseo61j3TiU9FV18EpkqytZLMq6t"
                    "reznSr4XwvjXp/g4eP8AGKDaNCKV9iR0ksk9m1Zk7Ps2rEjnPlkcC1rerTWMAbGGtGlG3l/LOQ84y8ua5HbFq26JleJkUbYa9avF/JDWhYemNpJL3uPU+R5L"
                    "nu7NDJMvRR6Ic06r+rjAeQZNVcV/CXTnmmMcp8gZmhH0cCXluGWVdlnHfG1VMkRpLNjeZHlNfTWuS1sVoiruO6+/enWVFaXeGld6P+p71Sx3E+E5LiVW1HNy"
                    "PltM0BUieHOpYmdxF27aMZa6MzRAwVY3H+990ucx0Oy/bv07enN7kXMKPK7Nd0eA41MbTLEjS1t3KsbJHVr1iQRI2s8mexICPbfFE1rnPLgy5N6knQ3jvX30"
                    "zX3D8qXW0XIFFZx884bzG0TOOBi/I1PBsa+IVuVao5TmNZPSW11h+SEmLaOQKy9cyGsqpeQ0VIpivn0s9Q8h6Y8xx3J6TXTwRh9TK0QWgX8XZ02zAC8ENlbp"
                    "s0DwWETRMHuMY55U3fUXhFH1B4pkeN3HCF9hrZ6FzRLqWQrn3K1gAEEtDx0St+5r4nvDmPH2nzkeW+LOTODeQct4r5Zwy/wHkLC7RdVkuL5HBVDsa6SRmpl5"
                    "lwjdg29PZMEmfj+S0kqzxrKKV6HkGNXNtRToFjIuB4tyvBczwdHkHHslXyONvxRvZLDJqStMWh0tS3AS2Wtbrv3HNDI1pa9rgzraA9Va8j4rmOJZmxguQUJa"
                    "F+q97eh+jFZhDtR26cwd0T1Jmhro5GHY6gyRscocxmeuk7r96qeii4kWHAHKU/F6i4tau2yvBLavrMlwHMlVZpSTF/jd5DltxXZ0NsqqbkGKScZzM6tLEOJl"
                    "ENuHB+lxL1C9IuB+pjGScnxJdkoIHQVszQsOpZSCPR6IzK0yR2IonOL4orUMzGuGgC3bDlfBfU/mfp46SLj+QYcdNIZp8Vfh/V0HSkgvkjjL45K8sgHTI+Ga"
                    "IO6i9w9w9YlomfqYOtuRUqjROL+lqBbuNe2dr/RHKkhhhS0qJT0ate5t0TyCMlsHKlSWG3Ul9RFkt9zKtDR/RjwJtlr5eVctfVBLv04jxLZXgEEMNkUtEEbD"
                    "i2BpO9sI1s7mk+q/mJr9DOK8ebZc0gWDYyRjBLddba3WfDu4a6y4HXftvULfUd1b9RHVrl0DN+orljJOVckpoL9RRP27NHUU2O1cp5l6bExjEcSqMbw3G02j"
                    "kWvXdSqTH4NhkTlbUu5HNtnqiudjSL4L6c8L9N6MtTiWIhx7rZabt6d7rWSudDnOZHYuzF0phjLne3A0sjadEh7vuGiea8+5b6g24bXJ8m6xDW6zUx1dhq42"
                    "q9zWtdLFVa4gzvaO8srpZAC5sftMkc12dvTz6B+Vev8A5uqeP8RYtqHjamn10zmTleM0huFx5iBqTIl/t82ZBsK+RyDfQW3IWAUL8KcmVcPNXFtX/wBKU+Q2"
                    "MPFPWf1gxPpdxu5L+rgn5VfgfBgMRHLG6yJ5oyGZO1Ex4kgo1SWy+6/pM72tihMjuvpyP0k9K8h6icgqtkqzR8YozsmzWSkY5teWKJ4ccZXk0Wy2rI/uXsYT"
                    "7MXW+Uxn2+r0HuYulfjDljpRy7pBRVx8U4wu+I0cSYtFq4xSE4DXUlHFq+P7ahjSnFIXO4+nVVBeY8mU44grCigHINxJL7qoMNybK4Xk9DlsFh8mXo5ePM+/"
                    "I7qfPabZ/UTGVx2T+pJkbK7+bUjiCDoiyvKYLHZbA3uOWIGDGXcZLin12DpYyrJXNdrGNboARs6ehvgdIBBHZebB1HdPvMPSvy7mHCPNOLvY3neGTUty0xlS"
                    "JlFf1ExUoqLMsQuHYsT9+wvKo0J+dj1ucaLLUhqZUXcClymnyHHqi5TgXqBgPUbjtHkWAtRyx2YoxfoOfH+txV/p1Yp3Imuc+J7HteIZCPbsw9EsTiC7VUvN"
                    "+BZjgPILWBzcJb0Pkdj8gGPbUytJhZ7duo9wa132PY2zCC59WcmJ+wY3yZS6POvvqZ6Hcxl5NwRnSamuvJMB7NMByWubv+Os8aq2pLMBrKMdddiSmpEUpC22"
                    "r7E7fFMvbioRXN5K3V98Jfgeo/pFwf1RrxDklCZmQqxPipZrGzfpslVY/ZMfU9ktexC1590Q2IHjrA0QzbXe5wH1P5d6bPn/ALP2oZsfZlbNaw+QidPRmmBY"
                    "DKwRyRz1ppGgxulil6SxznmN0jWubMQj9Tj1VNxzjvdP/TxImmhBJsmX+Ro0Pv7jJa1U68umPdq067Wk3ijRsjN5ZkaSjk/6LMAZj7fOMy2v1dvcxNN0wYe4"
                    "6i2dkYcB9pPy4b6QXdA32z6tMr7IMvDKJn6B1NjytkRdevIc6m55a4/HTtoI+538xix64fU56pevewqI/MeUU9Nx/jMw7LG+IuO6yXjfHVbeHDVXLyedCsLf"
                    "IMgyXJlw3JEWPZ5Vkt0zQNTrhnEIONRL+8i2O+vTD0N4P6VPlu4VlvJZyzCa0udy7oJLjYHOLpIKUMMcUFOKQlrZRG1z5msiE0spYHLS/qL6w8v9SYoqWT/T"
                    "YzDxTCwzD40SthfMA1sclueZ7pbboSC+LqDGRySSlrB9rW6gcVcZcj818iYjxXxPh93nnImd3EekxTFqSLIkSrWe+SnFuyHmmHmamkrIzci0yTJrL2KLFceh"
                    "2eS5DOrqKqsJ7GxOU8swnDcHkeRZ69BSx2NhdJI6WT232JtH2KcDGAySWbUnTDFGxrndbxsdIcRr/jfFcty7M0sBhKslm/elawe2A5tWAuAnu2C5zGMr1ouu"
                    "WRz3tDg0xtLpHsafTV6Penal6TemLhXp4pFxJKOMsIgVV7aQSkJi5Dm9m9IyLkTLGW5X++ynLs9uMkyc460tpjqtlMNtMtNoaRS1y7kVnlvJ89yW2A2xm8pb"
                    "yD2AACNs8rnRRdiQTFF0Rl2yXFpcSSSTbTxvCV+N4DD4Cq5zq+Ix1THxOcS4ubWhbHvZA0CWktboBrdNAAAC84DrtfUnrd6xiM/CeqvqJQR6JJJIuYMyIt6L"
                    "4SRF58KM/ne/FwXpRL/7MvT0Eg9PCeKjQ89sJRHjsO/YfA2O+zvdXvqXCP8ApD5ufBdyzkJGj4Jy9s/Hck73rv2PYLdX0HnlOeqN04J0syJjmIzIyPtTrgDl"
                    "f7jPWi14Ts/O1ERf8RqT6tXB3o/Z7jY5NgD/AJ9OQHY9h4PYaOgHHz3GzfplY1vqe067/wBm8yNg7OjPi+7tfBPbvob1+dL0PRVsrEF5C2ByVFgmCEkzNKcJ"
                    "xNKFkgi8f0/XEnfaZaPt+SL/AL/GjK8LiMm+KcYHgjjuCb5P/wB66uvnXwe42NHZ7DSqQ5jA08u5Y49O3cnz7jojyctcB2R8k7340fHdWLv02LpL9Q7J0q3s"
                    "ulTlok9xGRmr/UrgczNPgtl2kflXnXj+DONH1kPDvTnBa/8AprgP76/h1/sRvudlu/gdvz23v9LEfTzjNEeP7Ou33BG/1sAA7dhseP6Hsr2z7DMll6NJaakR"
                    "5DTjD7DzaXWXmXUG26060slIcacQpSHG1pNK0qNKiMjMhWqCQQQSCCCCDogjuCCO4IPcEeFPcgEEEAgjRB8ELzqvV/6CHegzqfmVGGUioHT5yzEs824Lksvv"
                    "SIlHVxZUVrL+LXXJcmTP+u4ytbOsZrXJXuNTMCyHBXDsrW/Yyn6G1T6bvVg8/wCGsxGUsCTlHF2RVLxcGiS/jz1NoZAAEBzmtYa1npEYa9kRYw9biq5PX70y"
                    "HDeUuzOMi9vj3JZpLEEbNltDJgB92pt2z7dlxdagBc9/U6yCWMbC0x/cB878hdOHL+Ac4cYWqKXO+N8ii5Fj0uQyl6ItRMSa60qbFlTThvUuSUNhbYzkLEcm"
                    "5UnH7q1hxZEV99ElrcXNeJ4bnvG8nxbORNkp5GFzY5uwlp3GAuqXoHdyyStP0vB0OuIPjPZ5C1RxLkuS4VyPG8jw8jm2qEv99ED9l2lIWC1QnAcA6KxG0hof"
                    "9rZ2wz6JiarV3q2YRTeqB6evTn6i/TdX215P4ipcxmZlhEM5Uy7pMFzCXj9VzTQTayPQnNyHKuEeTONahE2ZHdgUf9FweRsrpHL2DNoVzoB+gufd6NerWe4R"
                    "y6RtCvlScHLdmPTXivQWGzYu4T7j4Y6mQjIa6Q9b4/ehBezoma6bfrTgR6remOJ5Nxdn66xj3Q56tVjAdPYqPqyxX6TOprX/AKqu2VxEQDfdfE5hZt0bmU/E"
                    "ynkGhSe4tJQ42ovg0mlC21kf9qkqIycSruJJkZGRER+bKY5ftY/qBB05pB7OB0QWlu+rbdaLXEeNfk1+OrNJe1w0Q5zHtJI0Qektc06IcDsOB7ghwOiFNH0U"
                    "+uX1Z9HvHFPw67XYRzbxbijcaFhdHyOi3hZPhWOxGEMMYbi+bY/PjOM4tDJtxVRBynH8wkULS26ijm12MV9VQwoweov0t8L5vmbOfx1+/wAWyWQkdPeipwwW"
                    "8basyDclr9LKY5a0ksn97L7U7mPe6Q+21zm9MiuD/UfyviOMrYTJ46rySlTibDUnnmmq5CCGMnoifOxs0dlrGERs64onNYxhdI49bnbA8zfqSesfP8duMe4s"
                    "484g4Het4pQ4+WVjdzydnFGrubN6Zj9jlianCWJzmlNNnfca5HHZjvOE0x9Ylicxi3G/o34fRtxWeScizGbiid1HH1IoMbVn/wCET2Ny2uj8shMZLgC6Ux9T"
                    "HZHn/qs5Laqug4/x3H4uWQa/XXppbksPbZMVVojgJOunrfK/QJ6Yw7TmwG5VaZ5l0y25LzKbmOVzcqyKxRd8g5XLu72dk+VlFrrS3RaZheOzJeQ5I1CtKida"
                    "rl2U2zYiWdO9NNuPOr/flVho+OYP9PxHBjG48YzHsmhwlJ0TH06BeYop312uMgjlkdoTygmZ4JMjnh2o15b+PZwTcrzbr945K86GTMWg58dm81jnvhjl6REH"
                    "RsY4/p4gxsTQehjW7101S1SULinokyULjaNRtlqQRsq25/ckjSv+4j8J8kryY92f+/ikh7f38b4T38iVjmH5DdgOPyO5J7Dz48Lf08kdgbDq8rJ29i77oXCV"
                    "um6IPdgIBBJ7difHqmdOvJzfNnT/AMHcyNRfoUcr8Q8b8jlB7/dOCrNsOpskXCN3Re4cRdkqObhERLNs1F4MhRhnMe/E5rL4uUObJjsneovDtBwdUtSwEOAJ"
                    "AO4+4BPdXF4q22/i8deY4OZco1LTXN30uE8EcoLdgHR6u2wDr4WZB5a76AiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiii9TD0peP"
                    "vUhf4nurzlDJeJMt4payqqjXFHj9blFdkeN5YummSKy7p5thSyDnUtpRR5WOWce4bYr49xlESXV2LlvDl1G4PSf1m5B6SS5g4ijQydbMsqfqKeQdYbFHYpOl"
                    "MFmJ1eSNwkDJ5Yng76mOABbo71l6j+lmD9S4sW3K2bdGfEyWHVrdFtczmK01gnruM8Ug9t7oopNAD+8jYSHaAUVa/wBLfx8euzrEzJPyau7h2lWajM9+Ncgo"
                    "0WiIjIyUoz893wRbo/8ATO5kf/tR4yQPAM2VP+u7Xfvr8dhr8rVn/or8Q/8Aph5H40B/1Zodtb/+YfP5+P287kW9NX0b8D9OnlLPeXa/mbIeW8qzDAlcc1yb"
                    "HDKzDoFBQWGRUuTZApTce8ySXZzbSwxbF0MLRMrIsGPWykuQ57s5p6Dqj1W9d+R+rGNxOJyuLxmKp4q7LkGx459x4sWXwGtG+b9VPKP7mF8zY9DY95+iA5wd"
                    "sb069H8B6b3slkMXeyOQs5KrFTfJkBU3DXilM5ZCa1eAgSyBjpOrYcY2f8LdTNjRy20gIokuvT0a+lbrpubHkexRccNc7Tosdmbyvx5HrVFlzkCExXVf+p2H"
                    "2DP7TmxQIMSFXt3MSVjOdftNdWUbWasUdfFrW9zemXrpzf0vBqYqxBk8FJIZJcFlRJLTa5xBkkpyMe2WjLIQOt8JLSOrcfU4uWruf+kPEPUPpsZWvLTy0cft"
                    "Q5jHlkNwRguLY5w5ro7cTC97mRztc1rj1fGlX+5U/THdXdRkjjHCnOvThyHiRxmlt2/J0/kzhnIkTHNnJiljWMcd89V30jB9qWp/9Xk9LIjNyuieEiTtD60s"
                    "K6vGcpwXKQWgNSMx2Yqz1idEdUclmpBM0HyGOjf0n7TI/XUo+W/pPyDZntx/Mqj6u9xm5iZm2Okg/bI2vdERLewD2EdX83Qz+U5D4R/TBcx2aDldR/Utxngx"
                    "sWDSCoOEsfyjlVVnUGlKpDiMyzuDw2nHrUld7LCXMCyyESTS+8ThEcQ/Ozn1pj2SzjPCCJ3scDYzmTDmQOI0CytRhHv62XbksQt2GgxOAO+9ifpQaJg/O8vc"
                    "+BjgRDiccIpJW7JLXz3J5hEPAIjhe4gu1IwgOU/fE3pOdInBHTbzZ068U4tY0z/P3Fub8Wcjc0378DJ+ZbyrzbGJONPyXsmkV8OFXwqg32ruow3GqrHMDhX7"
                    "LtzGxli0sbObMiznvVzm3KOV4nlufybr1zCZGlkcdRYDWxdR1K021HBXqQuDYmOc0sfJt9hzHEOmd21InD+m3E+P8byPGMNjmU6eVo2qV6y7Vi9ZFusa0k1i"
                    "zKC+aQN6XBrtR9bQQwd1EG5+lz49UaTR1h5ok+7a+/h+iWlRGR7JKU542SFGZlpX3aI1Fr7vG/h9aHNNgu4lxgkaOxLlG+Dv/wDOj22T2Gv28LTH/op8P7j+"
                    "0XI9H/8AZuz213IpDfk78b/qFKB6Z3pYMem5ecvzaHn+95WoeXqjCYNljdvgcHFmKqzwOblL9Pdw5kbKb81SXYmY3VfYMtRYaJzRVq5Ljv7VCQ1pj1a9Ysj6"
                    "uy4azl8FjcXcwsVqtDYx09twnq2ZGSmKaKw+RvUyVnU2Vpa7Ti0jWtbV9NvTDH+mcOUq4rLZC/Uys0NmWG/HV6orEMfs+5FLXiid0uiDWmNwc0EdQPU5xOP/"
                    "AFHPRkw71D+ccb5vvOeMn4uscf4nx3in9gqMJq8mgzIeNZfyDl0O4OZLyCmkMSXpHIU+HIie2+ybMCK4042tx8let6Vevmd9J8JksHisJicnDksr/FZZ78tx"
                    "kscn6SvU9pja0sbDH012v24F3U4gEDS871G9G8L6k5TH5TKZXJ0JcdQdQhiotqGN7HWJLBkkNiCV3X1SBoDS1oDe4JPaO+d+li43sa2wrJHWLm5M2cGZXy1t"
                    "cP0SXPp5sZ2K97ClZ6v23PaeUaFK91KVERqQotkex7X1i8utVrNV/E+OtjtV568jmWMoJGsnjdG5zHGwQHNDiQSCNgdtAg4NU+mHi9K5Tuwciz3u0rVe3EHx"
                    "45zHSVpmTRiRoqtLmlzNOGwdE6IOiLXZeCIj8mRFvRaL/wCPOv8AtsQ/UmVoX1pemv0n9eNbHPm/B5ETOquvbq8f5fwKZGxflKirmDslxKxN85X2dbklFBkX"
                    "FpOgYvnFJlWLQrOdItYtKzZrKYWwuB+qXNfTe0+fi+Xkr15nB1vF2W/q8TcIAG7FGQ+37mgAJ4jFO0AFkrS1pGFcx9PuKc7rMr8ixcVmSEEVr0J/T5GrsgkQ"
                    "W4wJGscQOuN3XE8ba5hBKruc1/pfeSK+FMm9OfVDhWXz12ZprMX5vxK849ai0qnXVpctORcCPksra2jtew0tEHiehhTXffkJ/bULRFRKnBfWjKIhFyfhLJXt"
                    "jAbbweTMLpJNdJMlO9BIxoPcj27LQ3XT7ZB6hHXM/SlWfKJOP8slrxucS+tlqAstY0+PanrTwPOhodMsTyT9xk32OscH9M/6ibs+M1YchdF0GtcktInz4fMH"
                    "OFnPjRVOET8iJVO9LVTHnSWmlKcZivXFa2+siYXNjIX9Q3kb/rS4xr7OEZ4kb6C7KY8DwQA4CuSNdvDiACdeBvxR9KGX8u5jjPPdrcTbG/GyHG84N2N/4OxA"
                    "7EEhSc8Gfpj+D8VvGLXqC6h835iqWEJUWH4Dh8bhmsnuLLtfh3187lnI2USq821LJp/ErHA71D6WZTVvHQTkNzWHK/rD5ll6k9PjeAxXFxO10ZvOsTZfJMjc"
                    "CC6CWWKpWrzb05krKrnxbcGuLiHjPeN/S9xLGTw2c9lsjyJ0ThJ+kdFDjqDpA4OAkihdNPNF2IfFLZMU2wXxgbYbHnF/FXG3CeD0fGnEmD4zx1gWNsvM0uKY"
                    "lUxKangnKkOzJskosRtsn59lOfkWFrZSjfsbWykybGxlSpsl+Q5E/J5TI5q9YyeWvWsjkLchls3Lkz57Ez3HZL5JCXaHhrRprRprQGgBSQoY+ji6kFDG1K9G"
                    "lWjbFXq1YmQwxRsAa1rGMAA0ANnyfJJPdd/HQXcWm3V50DdLnXDj0Cn5/wCOY11c0Ta28U5CoJTuNcj4m24p1xcSoyyuJMqTRuyHlTJeI37V3htlPbiz7PHp"
                    "syFCfj5rwn1D5d6e5D+I8VzE+PkeR+oquDZ6FxrS0htulKHQTDbG6eWiVoHS2QNJBxTlfCeMc2o/oOSYqC/G0H2ZjuK5VcWub11bcRbPA7T3fyP6fuJLSVXN"
                    "5v8A0v8Ak8eHc2XTh1SUd5NTMM8bwznLDJuMpVXrkumhOQ8qcfOZOy5PiwjYbW9U8Hw4lhKQ88iHVMuNx2ZZcf8ArSusijh5TwqtakDQZb+DyL6cksga0H/c"
                    "bkNiGNr3B7i5lnTAQ0RvA2o25n6UMe+V8vHeV2qkbnfZUy9GO8yFm3EdNmtNUmkLftAbLG4uPU4yt30rTxr9NJ6jzjqW3816K47XcXuvI5x5wlGaPH3NNn0l"
                    "RTUpKdl7a1NIUfy4RkZnlZ+tLihbv+xPIOvx0/xTGkAd/LzWb+d/ynuSVjw+k/OdRB5liizudjD3Ad9yBr+Iu2N6O9gjXYeAJIuB/wBMVxpQ2VNc9RnUlk/I"
                    "EViLHlWWDcVYfF48grtVIjKk1UvOshus1trrH21FNilNp8W4/wAhmNORZkeZRPsuR3Nbcr+sfleRinq8T47jeOMf1MjyFyd+YyAYS4CRsT4q1CGXpLS0GvYE"
                    "bgSHv32zvjf0t8Yx8sVnkeayGee0te+lBGzF48ub0u6SY3zXZGdQIduzGJGHpLG992QOD+BuHem3jum4o4M49x3jXAqNKDi0WPRFNnNnFEiwpF5kFrKck3WU"
                    "5RZMQYn71leS2Ntkl48wiVcWs6Tt44lZrN5bkWStZjOZC1lMnckMlm5cldNNI4knXU46axu9MjYGxxt+1jWtACkni8VjsJQrYzE0q+Px9SMRVqlWJsUMTB8N"
                    "Y0AbJ7ucduc4lziSSVlweWvQWq/VT0VdNPWhiMfEeoXjKozEqtuQnGMrjOSaHP8AC1y5MCZLXiGcUrsLI6OPPlVda5c1DE9VDkbcCNCySpuK5Coasr4jzflP"
                    "BciMpxbM28TaOmzCFwdWtRtO/auVZA+vaiP/AATRvA760e6x7knFOP8ALqBxvIcXWyVXfXGJm6lgk0WiWtOwtmrytDjqSJ7XDflVt+pT9NnxzgVFnPJOC9bt"
                    "VxvgWP1btpHR1K4tj0Cnqn1JZjx28y5tx7JsLx3HaORaPNREWrfEcp+AzKjsqiXEtsvq5b8T+sbk8r6uNznB6vIbcp6DNx+exSvT9Ic55hxgrXoJHiIFxZGY"
                    "x9rnuPT2bGjkf0s8ca2e5heVWsHCzRbDmIoL1GEOLWNElt01Wy3chA63yP31NY0Nd9x0QrP063Xbdx4djTczdAlxTzUIlRLKs6h+X50eZAd7jYlxHonSu7Fl"
                    "tutkhSDaeJhe1JbeUlCVrzWX6x8DA50cvAeTwytJZJFLbosfE9pG2OD67ST8kFkbgfPfzi0f0sX5W9cPOcNLG9vVHJHjJ3skDgfuBbkXfaTsDTpAe57A9I3X"
                    "4k/TAZfY1lbP5q6tcTop7ikFd4zxLx1a5rEbaS+Xe3S8h5ZkuDKN5UQjS09P4tU0xJdS47EmMsGzIxXOfWlOQ6Pj3BIonlrg21m8rJK9hcO5FOlXhael3cF9"
                    "pzT0/cw9RA97D/ShUD2yZzl887Gu718VjooWvaCCA+xblsPaHD+drIWuHboe0gFWIOjD05ulXoQpZMfgzBnF5pb1pVWU8vZtKj5JyplUJX7U5JgTchTBroND"
                    "QzZdJV2crDsHpsUwpy5houixxNst6Y5FDnfqZzL1HvC5ynLyW44pHyVMdA39Ni6BeXbFSjGfajOnFvuu9yYsPS6UhSR4dwHivBKRp8bxcVQyMYyzckJnyFzo"
                    "7tNu5JuaXR7hhIjaf5WNW9AwFZkqznNf6bvCOaubOYOZbHqwyyie5a5Tz/k17H4PE9PKbonM8y24yt6matJGcpXPbrXLZUJqcqFCVJQwl9URlTntIlbx36su"
                    "WccwGE4/U4zx6avg8PjMPBPPLkfemixlKClHNKI7DI/dlZA17w1gaHOd0j5UdM59N3Fs9mstm7WczsdjL5O/k5oof0AhjkyFmW1JFGH1Hu9uN0rmRlzi/pA6"
                    "nE7Kzf0O+g/hPRR1KYD1IVXUblfINngbOXNRsWsOPafH6+wPK8JyPClqeso+TW0mOUGNkkichLUdXvvx2mlG00pYxn1K+ovkXqbxiTi+TwGFx1SS9SvixQkv"
                    "OnbNRM3Q3VmeWMxvE8gcOkOB6SHaBDvd4D6G4H0+z7eQYzMZe5ZFK1QMF0U/ZdDadA95/uK8T+tr68RaS4tADh07ILZ9BHhbuVTao/Ss8eU9PU0zHWXmjzFR"
                    "W19VHde4apTdcj1sNiEwp008iJJTq246FuqIkpU4au1CUmSSl7jvrB5bjKFHHV+J8ddBj6dWlAZJ8kZTFUrxVozI5s7Wl5jiBcWsYOskta0aaI2ZL6ZuLZTI"
                    "ZDJWM/nWT5G7cvzNibjxE2W7YlsytY19V7wwPlcGBzy4MABc523GQf06/RSxP0+OfZ/PNL1AZHydYT+Msp40cxyzwGtxiEiLlF/hV+9alPi5Rdvm/Eewthhq"
                    "N7BIcbnOqU8n2SS7r71S9fs76q4KpgsrgsTjIqeRjyMc9Ca6+R0jIJoPbe2zLIzoLZerbekhzfkHQy/099GML6dZaxl8ZlspeltUnUpIrzanQGGVkoe11eGF"
                    "wcHN1pxcNEjQ7kzhDQa3ItNeu3om4y6+eArTgzkqdZ46aLyqy/Cc5oY8KVf4JmlKmXGh3tbFsUOQZ0edT2d1jF9WSSbKyxu/t4kWXWWS4FvX5nwLnWb9O+SV"
                    "OS4J7P1VdskM9afrNW9UmHTPUtNjcxzoZAGu+1wc17GPB2NHFuZcQxHOMDb4/mo3Oq2THJHNF0CxUsQvD4bNdz2ua2WN4BG2lrgS1wLSQoIv/wALjxgfbvrA"
                    "5C8JUStcVYwRmozPykzyo+1Ou0jSfeZmnfeRaJMkP/TK5rvf9leM/wD8+T8f/wBV+5HzsHuN7J0T/wCixw8dhn+Q67bBOP76351UH5I/5+dESx+mx6bk305q"
                    "rkrE6XqDyDlvAuR59RkTmI3+FRsbi45mVVDTSP5PRvw8quGWn8hxqNTUmTRnIBlZFi+MymJEA6+WxY6N9UfVC16p5Onm8pgcXistVrinJbxclsC7UjJdXZaj"
                    "tSz9UlbqLYpY3x/3bixzHfaW7c9PPT2t6dY6ziMdmclksXPYfbiq5JlRxp2JSDOa0taGuRDOdvfFK2QiQ9TJGjbTql1ffp7+kjqGvLXOOGru36WM3vJxTreB"
                    "hlBAyjiCwfXHajvyWuKpNhjqsWkrTHZcbice5fheMrlO2VlZYzaW9k/YFnnp19TfP+B1oMVbNflmDrhrYaOYkkbcrRt7CKrlWNksRxdJcBHMyzG1xD2sGiHY"
                    "hzr0B4VzOxNkoWz8dzMxe6W/ims9mxI8lxlt495bXnl69EytMUzmD2/dDenohazD9Mx13wMmtYmA8s9JOVYczLebpMgy7POYsAyWxgEeo8m1w6m4G5JqqSYt"
                    "O1OwYmdZAwyf2IsHy+899wfWngPbYZ+DZhkrmt91kOWoyRRu19zInyVWPlaHEhj3xxO1/hZvTdLy/Sblepwh5jjjGHFsZkxNnrczZ6XSBt0Na4t0SxpeGOJD"
                    "XO/mO63Tt+mPp4D2O3vVR1FvXZobRJyXjbhDH3ayvVNadNRVcblvMlLsrWhkoShue6xxNiF6ppx5FXZVMtDFkjDOWfWVnb1aapw/jFTBvkaWx5XJ2zlLcYc0"
                    "BzoqTIK9KORpJMb5XWtODJOgEFhynjX0q4OlYhtcnz9vNNjf1Ox1GsMbUl6XbaJbPvT3HMdpvWyJ0HU0vY5zw7YlT6svR66ceobpg4k6W+M/pem7EOF83ZzH"
                    "ELPDsYYyqxfS/jNxjuRV+RSr64jX2Sz8vdn1WQZTld3kc7KMgyHG6q0vbO1kfULd0Xwb1r5dwvl+V5pLIOSZXM05al8ZmxZMU5ksRWGSv/TvjcP05ZI2CGMx"
                    "wxe6elga0MO4uX+k/GOWcZx3FWxnB43FWorNL+EwwMfAI4JYDFGJo5IwJWyN9172Pe4RjZ6tObGGf6XXjwlqU31hZuhPak0JPiKjUpDyfJrNf9eJ7mzURGTZ"
                    "JQpJdxe6fdstzu+s/mrgf/VPi486+/Kkb/JH6sbP5/5aIBWrB9KvDwAP7Q8iP/8AbP8A/S/+H7KwN0YdOVt0ldN3HHTzacpXPMTfGcW5qKTNb+mbobNWNS8h"
                    "tbfHscOvat7xLVbh1TYRcUoUfuDvsUFPWRSShMdKSi5y/kR5byPK8jfjqmKmzFp12zTomQ1GWZQDPJCJS6RonkDpntc959x73dXdSE4zg28awWNwMd2zkIcX"
                    "Wjp17VwR/qn14h0xNmdE1kbjGzTGlrGAMa1oaAAtoxjS95ARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARARAR"
                    "ARARARARARARARARARARARARARARARARARVwP1K/PcnBOlPifgSsJBTOf+Tnbm/N5ztQvA+FWKrJJzMdCJDLqp5cj5HxVNacWzKhogwbJt9DUl6DIalR9I3G"
                    "v4t6ly56SMSV+KYmzbIc3bf1mSjlx1Q7PbrYJLErGEbPtukB1E4GO31M8gGK9Pf4Qx/RY5Lka1FnSR1mvTezIW/t1vodHA2J79gNMrGHZkGqPS6+lkqSl6lq"
                    "JKlLLy/VV0hRqUetEt2Mpf5Pxs1GpSj/AOR7srfQxk5ImxmOsb7ET4+lP1Anu0mWF2w4+T48kgqAMOQzEGv0+WylfQPT7GSvQNHT8hsczdkHWu+hoHtoa9Fn"
                    "0Q+nhHTx6dXC7UhENF9zW3O6hb79vjJhwyb5Rj10vCGm4hR4yo8qJxbW4FCuELS533sa0facNh5pKag/W3kVLkvqVyS1i69WtiaNs4fGw0q0FSsK2NLoHyxw"
                    "1mtiAsWv1FgPDQ57ZQ54a4lotB9J8Hd4/wAC4/Uyc9qzlLFRuRyM121PcsGzeAn9t89l75SIInRQNZ1FsYj6WbaATLWNULYyAiAiAiAiAiAiAiAiAiAiAiAi"
                    "AiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAiAijn"
                    "63fS16VuvyzrMm5wi8kV+dY9iiMNxXN8C5DuKSdi1IVxOvHyqcSuUZJxg9YzpthIRZWtzgVtYWERuuiS5DrNNTJr9k8B9WOa+mv6pvFb9atBdlE1yrax9S3F"
                    "ZlZEYo3Svkj/AFHTE0ksjZOyMPJcWEk7wXmfpxxPnra45JRmsyVGOjrTwXLNWWBj3iR4Z7Ugj29zQHOdG5+hprmhRFXX6XzhEsioJeK9UXJqMTi5JWycpx7P"
                    "MFxXJrjIsNRNZO8xyBkuG2PGUPH7myqvq4UPJSxe1j10l1iYePyiZUw5u+L6wOcPxt6lfwOCsWrVGxVhyNV1qjLVnmrvhbZEDXzQye25/utjHtO6+/u9mgaj"
                    "k+mDhrchRuUsrma0FW7XtS0JjXuQ2Y4J2TGB8ssYmaJA3oe4ukBaSAxu9i0PGjR4cdiJEYZixIrLUaLFjNIYjxo7CEtMsMMtJS2yyy2lLbTTaUobQlKEJJJE"
                    "RRIc5znOc5xc5xLnOcSXOcTsucTskkkkknZPcqSoAaA1oDWtADWgAAADQAA7AAdgB2AX7D8X6gIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIgIg"
                    "IgIgIgIgIgIv/9k=' alt='Vaisesika logo' style='max-width:360px;width:70%;height:auto'>"
                    "</div>"
                    "</div>"
                )
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
                return

            if path == "/logout":
                cookie_header = self.headers.get("Cookie")
                clear_session_from_header(cookie_header)
                self.send_response(303)
                self.send_header("Location", "/")
                self.send_header("Set-Cookie", "session_id=deleted; Path=/; Expires=Thu, 01 Jan 1970 00:00:00 GMT")
                self.end_headers()
                return

            if path == "/upload":
                self.send_response(303)
                self.send_header("Location", "/upload/PPM")
                self.end_headers()
                return

            protected_prefixes = ("/user-management", "/upload", "/email-settings", "/export", "/timesheet", "/help", "/help/edit", f"/{UPLOAD_DIR}/")
            if any(path == p or path.startswith(p + "/") for p in protected_prefixes):
                if not session_email:
                    self.send_response(303)
                    self.send_header("Location", "/")
                    self.end_headers()
                    return

            # HARD BLOCK: Manager no access to user management
            if role == "Manager" and path.startswith("/user-management"):
                return self._msg("Access denied: your role does not have permission to access User Management.", display_name)

            # Employee can access only allowed modules + uploads + timesheet
            if role == "Employee":
                allowed_exact = {"/", "/logout", "/help"}
                ce, cx, ch = get_feature_flags_by_email(session_email)
                allowed_prefixes = [f"/{UPLOAD_DIR}/", "/timesheet", "/help"]
                if ce: allowed_prefixes.append("/email-settings")
                if cx: allowed_prefixes.append("/export")
                if screen_access in ("PPM", "BOTH"):
                    allowed_prefixes.append("/upload/PPM")
                if screen_access in ("NTT", "BOTH"):
                    allowed_prefixes.append("/upload/NTT")

                if screen_access == "EMAIL":
                    allowed_prefixes.append("/upload/EMAIL")
                if not (path in allowed_exact or any(path.startswith(p) for p in allowed_prefixes)):
                    return self._msg("Access denied: your role does not have permission to access this page.", display_name)

            # Secure file serving
            if path.startswith(f"/{UPLOAD_DIR}/"):
                if role == "Employee":
                    stored_fn = os.path.basename(path[len(f"/{UPLOAD_DIR}/"):])
                    owner = get_upload_owner_by_stored_filename(stored_fn)
                    if owner != (session_email or "").strip().lower():
                        return self._forbidden("Access denied: you can only view your own uploads.", display_name, role, screen_access)

                self.path = urlparse(self.path).path
                return SimpleHTTPRequestHandler.do_GET(self)

            # User management create (Admin only)
            if path == "/user-management/create":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can create users.", display_name)

                custom_fields_html = build_custom_fields_form_html(list_custom_fields(active_only=True), {}, name_prefix='cf_')

                content = (
                    "<div class='card' style='max-width:720px;margin:0 auto'>"
                    "<h2>Create User</h2>"
                    "<form method='post' action='/user-management/create'>"
                    "<label>Username</label><input type='text' name='username' required>"
                    "<label>Email</label><input type='email' name='email' required>"
                    "<label>Password</label><input type='password' name='password' required>"
                    "<label>Role</label>"
                    "<select name='role' id='roleSel' required>"
                    "<option value=''>Select role</option>"
                    "<option>Admin</option><option>Manager</option><option>Employee</option>"
                    "</select>"
                    "<label>Status</label>"
                    "<select name='status' required>"
                    "<option value=''>Select status</option>"
                    "<option>Active</option><option>Inactive</option>"
                    "</select>"
                    "<label>Screen Access (Employee/Manager)</label>"
                    "<select name='screen_access' id='saSel'>"
                    "<option value='BOTH' selected>PPM &amp; NTT Only</option>"
                    "<option value='PPM'>PPM Only</option><option value='EMAIL'>EMAIL Only</option>"                    "</select>"
                    "<hr style='margin:16px 0;border:none;border-top:1px solid #eef2f7'>"
                    "<h3 style='margin:0 0 6px'>Module Access</h3>"
                    "<label>Email Settings Access</label>"
                    "<select name='can_email' id='ceSel'>"
                    "<option value='0' selected>Disabled</option>"
                    "<option value='1'>Enabled</option>"
                    "</select>"
                    "<label>Export Data Access</label>"
                    "<select name='can_export' id='cxSel'>"
                    "<option value='0' selected>Disabled</option>"
                    "<option value='1'>Enabled</option>"
                    "</select>"
                    "<label>Help Access</label>"
                    "<select name='can_help' id='chSel'>"
                    "<option value='1' selected>Enabled</option>"
                    "<option value='0'>Disabled</option>"
                    "</select>"
                    f"{custom_fields_html}"
                    "<button class='btn' type='submit'>Create</button>"
                    "</form>"
                    "<script>"
                    "(function(){"
                    "const role=document.getElementById(\'roleSel\');"
                    "const sa=document.getElementById(\'saSel\');"
                    "const ce=document.getElementById(\'ceSel\');"
                    "const cx=document.getElementById(\'cxSel\');"
                    "const ch=document.getElementById(\'chSel\');"
                    "function sync(){"
                    " const r=role.value;"
                    " const canSetScreen=(r===\'Employee\' || r===\'Manager\');"
                    " if(sa){ sa.disabled=!canSetScreen; if(!canSetScreen) sa.value=\'BOTH\'; }"
                    " const isAdmin=(r===\'Admin\');"
                    " const isEmp=(r===\'Employee\');"
                    " if(ce) ce.disabled=isAdmin; if(cx) cx.disabled=isAdmin; if(ch) ch.disabled=(isAdmin||isEmp);"
                    " if(isEmp){ if(ch) ch.value=\'0\'; }"
                    " if(isAdmin){ if(ce) ce.value=\'1\'; if(cx) cx.value=\'1\'; if(ch) ch.value=\'1\'; }"
                    "}"
                    "role.addEventListener(\'change\', sync); sync();"
                    "})();"
                    "</script>"
                    "</div>"
                )
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
                return

            

            # User management: Create Fields (Admin only)

            # User management: Edit Field (Admin only)
            if path == "/user-management/create-fields/edit":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can manage custom fields.", display_name)

                qs = parse_qs(urlparse(self.path).query)
                fid = (qs.get('field_id', [''])[0] or '').strip()
                try:
                    fid_i = int(fid)
                except Exception:
                    return self._msg("Invalid field id.", display_name)

                row = get_custom_field_by_id(fid_i)
                if not row:
                    return self._msg("Field not found.", display_name)

                _id, fkey, flabel, ftype, freq, opt_json, active = row
                ftype_u = (ftype or 'TEXT').strip().upper()
                opts = _parse_options_json(opt_json) if ftype_u == 'DROPDOWN' else []
                opts_str = ', '.join(opts)
                checked_req = 'checked' if int(freq or 0)==1 else ''
                checked_active = 'checked' if int(active or 0)==1 else ''

                content = (
                    "<div class='card' style='max-width:860px;margin:0 auto'>"
                    "<h2>Edit Field</h2>"
                    "<p class='muted'>Update label / required / options and active status.</p>"
                    "<form method='post' action='/user-management/create-fields/update'>"
                    f"<input type='hidden' name='field_id' value='{_id}'>"
                    "<label>Field Key</label>"
                    f"<input value='{html.escape(str(fkey or ''), quote=True)}' disabled>"
                    "<label>Label</label>"
                    f"<input name='label' value='{html.escape(str(flabel or ''), quote=True)}' required>"
                    "<label>Type</label>"
                    f"<input value='{html.escape(str(ftype_u), quote=True)}' disabled>"
                    "<label style='margin-top:12px'>Options (comma-separated, for DROPDOWN only)</label>"
                    f"<input name='options' value='{html.escape(opts_str, quote=True)}' placeholder='e.g. A, B, C' {'disabled' if ftype_u!='DROPDOWN' else ''}>"
                    "<label style='margin-top:12px;display:flex;gap:10px;align-items:center'>"
                    f"<input type='checkbox' name='required' value='1' style='width:auto' {checked_req}> Required"
                    "</label>"
                    "<label style='margin-top:8px;display:flex;gap:10px;align-items:center'>"
                    f"<input type='checkbox' name='active' value='1' style='width:auto' {checked_active}> Active"
                    "</label>"
                    "<div class='controls'>"
                    "<button class='btn' type='submit'><i class='fa fa-save'></i> Save Changes</button>"
                    "<a class='btn outline' href='/user-management/create-fields' style='text-decoration:none'><i class='fa fa-arrow-left'></i> Back</a>"
                    "</div>"
                    "</form>"
                    "</div>"
                )
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode('utf-8'))
                return

            if path == "/user-management/create-fields":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can manage custom fields.", display_name)

                qs = parse_qs(urlparse(self.path).query)
                msg = (qs.get("msg", [""])[0] or "").strip()
                fields = list_custom_fields(active_only=False)

                rows_html = ""
                for fid, fkey, flabel, ftype, freq, opt_json, active in fields:
                    badge = "<span class='badge'>Active</span>" if int(active or 0)==1 else "<span class='badge' style='opacity:.6'>Inactive</span>"
                    rows_html += (
                        "<tr>"
                        f"<td>{fid}</td>"
                        f"<td>{html.escape(str(fkey or ''))}</td>"
                        f"<td>{html.escape(str(flabel or ''))}</td>"
                        f"<td>{html.escape(str(ftype or ''))}</td>"
                        f"<td>{'Yes' if int(freq or 0)==1 else 'No'}</td>"
                        f"<td style='max-width:260px;word-wrap:break-word'>{html.escape(str(opt_json or ''))}</td>"
                        f"<td>{badge}</td>"
                        "<td>"
f"<a class='btn outline small' style='margin-right:6px' href='/user-management/create-fields/edit?field_id={fid}'>Edit</a>"
"<form method='post' action='/user-management/create-fields/deactivate' style='display:inline-block;margin:0 6px 0 0'>"
f"<input type='hidden' name='field_id' value='{fid}'>"
f"<button class='btn danger small' type='submit' {'disabled' if int(active or 0)==0 else ""}>Deactivate</button>"
"</form>"
"<form method='post' action='/user-management/create-fields/delete' style='display:inline-block;margin:0' onsubmit='return confirm(\"Delete this field permanently? This will remove all saved values too.\");'>"
f"<input type='hidden' name='field_id' value='{fid}'>"
"<button class='btn small' type='submit' style='background:#111827'>Delete</button>"
"</form>"
"</td>"
                        "</tr>"
                    )

                alert = f"<div class='muted' style='margin-bottom:10px;color:#065f46;font-weight:800'>{html.escape(msg)}</div>" if msg else ""

                content = (
                    "<div class='card' style='max-width:1000px;margin:0 auto'>"
                    "<h2>Create Fields</h2>"
                    "<p class='muted'>Define custom user fields (Admin only).</p>"
                    f"{alert}"
                    "<form method='post' action='/user-management/create-fields'>"
                    "<label>Field Key</label>"
                    "<input name='field_key' placeholder='e.g. employee_id' required>"
                    "<label>Label</label>"
                    "<input name='label' placeholder='e.g. Employee ID' required>"
                    "<label>Type</label>"
                    "<select name='field_type'>"
                    "<option value='TEXT'>Text</option>"
                    "<option value='NUMBER'>Number</option>"
                    ""
                    ""
                    "<option value='DROPDOWN'>Dropdown</option>"
                    "</select>"
                    "<label>Required</label>"
                    "<select name='required'><option value='0'>No</option><option value='1'>Yes</option></select>"
                    "<label>Dropdown Options (comma-separated; only for Dropdown)</label>"
                    "<input name='options' placeholder='Option1, Option2, Option3'>"
                    "<button class='btn' type='submit'><i class='fa fa-plus'></i> Create Field</button>"
                    "</form>"
                    "<hr style='margin:18px 0;border:none;border-top:1px solid #eef2f7'>"
                    "<h3 style='margin:0 0 10px'>Existing Fields</h3>"
                    "<table class='ts-grid um-table' style='width:100%'>"
                    "<thead><tr><th>ID</th><th>Key</th><th>Label</th><th>Type</th><th>Required</th><th>Options</th><th>Status</th><th>Actions</th></tr></thead>"
                    "<tbody>"
                    f"{rows_html if rows_html else '<tr><td colspan=8 class=\'muted\'>No fields yet.</td></tr>'}"
                    "</tbody></table>"
                    "</div>"
                )
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
                return
# User list (Admin only) with Active/Inactive filter and row count number
            if path == "/user-management/list":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can view user list.", display_name)

                qs = parse_qs(urlparse(self.path).query)
                status_filter = (qs.get("status", ["All"])[0] or "All").strip()
                if status_filter not in ("All", "Active", "Inactive"):
                    status_filter = "All"
                rows = list_users_db(None if status_filter == "All" else status_filter)

                filter_html = f"""
                <form method="get" action="/user-management/list" class="controls um-filter" style="margin-top:0">
                  <div style="min-width:160px">
                    <label style="margin:0">Status Filter</label>
                    <select name="status" style="margin-top:6px">
                      <option value="All" {"selected" if status_filter=="All" else ""}>All</option>
                      <option value="Active" {"selected" if status_filter=="Active" else ""}>Active</option>
                      <option value="Inactive" {"selected" if status_filter=="Inactive" else ""}>Inactive</option>
                    </select>
                  </div>
                  <button class="btn small" type="submit" style="margin-top:0"><i class="fa fa-filter"></i> Apply</button>
                  <a class="btn secondary small" href="/user-management/list" style="margin-top:0"><i class="fa fa-rotate-left"></i> Reset</a>
                </form>
                <div class="muted" style="margin-top:8px">Showing <b>{len(rows)}</b> user(s).</div>
                """

                rows_html = ""
                for idx, (uid, uname, email_addr, rrole, status, sa, can_email, can_export, can_help) in enumerate(rows, start=1):
                    rows_html += (
                        "<tr>"
                        f"<td style='padding:8px'>{idx}</td>"
                        f"<td style='padding:8px'>{html.escape(uname)}</td>"
                        f"<td style='padding:8px'>{html.escape(email_addr or '')}</td>"
                        f"<td style='padding:8px'>{html.escape(rrole)}</td>"
                        f"<td style='padding:8px'>{html.escape(status)}</td>"
                        "<td style='padding:8px'>"
                        "<form method='post' action='/user-management/delete' style='display:inline;margin-right:6px' "
                        "onsubmit=\"return confirm('Delete this user?');\">"
                        f"<input type='hidden' name='user_id' value='{uid}'>"
                        "<button class='btn danger small' type='submit'>Delete</button></form>"
                        "<form method='get' action='/user-management/edit' style='display:inline'>"
                        f"<input type='hidden' name='user_id' value='{uid}'>"
                        "<button class='btn secondary small' type='submit'>Edit</button></form>"
                        "</td></tr>"
                    )

                content = (
                    "<div class='card'>"
                    "<h2>User List</h2>"
                    f"{filter_html}"
                    "<table class='um-table' style='width:100%;border-collapse:collapse;margin-top:10px'>"
                    "<thead><tr style='text-align:left'>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Number</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Username</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Email</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Role</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Status</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Actions</th>"
                    "</tr></thead><tbody>"
                    + (rows_html if rows_html else "<tr><td colspan='6' style='padding:8px'>No users found.</td></tr>")
                    + "</tbody></table>"
                    "</div>"
                )
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
                return

            if path == "/user-management/edit":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can edit users.", display_name)

                qs = parse_qs(urlparse(self.path).query)
                uid = qs.get("user_id", [None])[0]
                try:
                    uid_int = int(uid)
                except Exception:
                    return self._msg("Invalid user ID.", display_name)

                row = get_user_db_by_id(uid_int)
                if not row:
                    return self._msg("User not found.", display_name)

                uid_val, uname, email_addr, rrole, status, sa, _ce, _cx, _ch = row
                sa = normalize_screen_access(sa)

                cf_values = get_custom_field_values(uid_int)
                custom_fields_html = build_custom_fields_form_html(list_custom_fields(active_only=True), cf_values, name_prefix='cf_')

                role_opts = "".join(
                    f"<option{' selected' if rrole==opt else ''}>{opt}</option>"
                    for opt in ("Admin", "Manager", "Employee")
                )
                status_opts = "".join(
                    f"<option{' selected' if status==opt else ''}>{opt}</option>"
                    for opt in ("Active", "Inactive")
                )
                access_opts = "".join(
                    f"<option value='{opt}'{' selected' if sa==opt else ''}>{label}</option>"
                    for opt, label in (("BOTH","PPM &amp; NTT Only"), ("PPM","PPM Only"), ("EMAIL","EMAIL Only"))
                )

                content = (
                    "<div class='card' style='max-width:720px;margin:0 auto'>"
                    "<h2>Edit User</h2>"
                    "<form method='post' action='/user-management/update'>"
                    f"<input type='hidden' name='user_id' value='{uid_val}'>"
                    "<label>Username</label>"
                    f"<input type='text' name='username' value='{html.escape(uname, quote=True)}' required>"
                    "<label>Email</label>"
                    f"<input type='email' name='email' value='{html.escape(email_addr, quote=True)}' required>"
                    "<label>Role</label>"
                    f"<select name='role' id='roleSel' required>{role_opts}</select>"
                    "<label>Status</label>"
                    f"<select name='status' required>{status_opts}</select>"
                    "<label>Screen Access (Employee/Manager)</label>"
                    f"<select name='screen_access' id='saSel'>{access_opts}</select>"
                    "<hr style='margin:16px 0;border:none;border-top:1px solid #eef2f7'>"
                    "<h3 style='margin:0 0 6px'>Module Access</h3>"
                    "<label>Email Settings Access</label>"
                    "<select name='can_email' id='ceSel'>"
                    f"<option value='0' {'selected' if int(_ce)==0 else ''}>Disabled</option>"
                    f"<option value='1' {'selected' if int(_ce)==1 else ''}>Enabled</option>"
                    "</select>"
                    "<label>Export Data Access</label>"
                    "<select name='can_export' id='cxSel'>"
                    f"<option value='0' {'selected' if int(_cx)==0 else ''}>Disabled</option>"
                    f"<option value='1' {'selected' if int(_cx)==1 else ''}>Enabled</option>"
                    "</select>"
                    "<label>Help Access</label>"
                    "<select name='can_help' id='chSel'>"
                    f"<option value='0' {'selected' if int(_ch)==0 else ''}>Disabled</option>"
                    f"<option value='1' {'selected' if int(_ch)==1 else ''}>Enabled</option>"
                    "</select>"
                    "<label>Password (leave blank to keep unchanged)</label>"
                    "<input type='password' name='password' placeholder='********'>"
                    f"{custom_fields_html}"
                    "<button class='btn' type='submit'>Save Changes</button>"
                    "</form>"
                    "<script>"
                    "(function(){"
                    "const role=document.getElementById('roleSel');"
                    "const sa=document.getElementById('saSel');"
                    "const ce=document.getElementById('ceSel');"
                    "const cx=document.getElementById('cxSel');"
                    "const ch=document.getElementById('chSel');"
                    "function sync(){"
                    "  const r=role.value;"
                    "  const canSet=(r==='Employee' || r==='Manager');"
                    "  if(sa){ sa.disabled=!canSet; if(!canSet) sa.value='BOTH'; }"
                    "  const isAdmin=(r==='Admin');"
                    " const isEmp=(r===\'Employee\');"
                    " if(ce) ce.disabled=isAdmin; if(cx) cx.disabled=isAdmin; if(ch) ch.disabled=(isAdmin||isEmp);"
                    " if(isEmp){ if(ch) ch.value=\'0\'; }"
                    "  if(isAdmin){ if(ce) ce.value='1'; if(cx) cx.value='1'; if(ch) ch.value='1'; }"
                    "}"
                    "role.addEventListener('change', sync); sync();"
                    "})();"
                    "</script>"
                    "</div>"
                )
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
                return

            # Capture pages: /upload/PPM or /upload/NTT or /upload/EMAIL
            if path in ("/upload/PPM", "/upload/NTT", "/upload/EMAIL"):
                module_ctx = path.rsplit('/', 1)[-1]

                # Employee/Manager module permission check
                if role in ("Employee", "Manager"):
                    if module_ctx == "PPM" and screen_access not in ("PPM", "BOTH"):
                        return self._msg("Access denied: you do not have permission for PPM module.", display_name)
                    if module_ctx == "NTT" and screen_access not in ("NTT", "BOTH"):
                        return self._msg("Access denied: you do not have permission for NTT module.", display_name)
                    if module_ctx == "EMAIL" and screen_access != "EMAIL":
                        return self._msg("Access denied: you do not have permission for EMAIL module.", display_name)

                # Admin filters
                qs = parse_qs(urlparse(self.path).query)
                month_filter = (qs.get("month", [""])[0] or "").strip()
                name_filter = (qs.get("name", [""])[0] or "").strip()
                upload_error_alert_js = ''
                err_msg = (qs.get("err", [""])[0] or '').strip()
                if err_msg:
                    safe = html.escape(err_msg, quote=True)
                    safe = safe.replace('\\', '\\\\').replace("'", "\\'")
                    upload_error_alert_js = f"<script>window.addEventListener('load', function(){{alert('{safe}');}});</script>"

                if not re.match(r"^\d{4}-\d{2}$", month_filter):
                    month_filter = ""
                name_filter = name_filter[:60]

                # UPDATED FILTER UI (Admin only)
                filter_ui = ""
                if role == "Admin":
                    filter_ui = f"""
                    <div class="filter-card">
                      <div class="filter-title">
                        <div class="muted"><i class="fa fa-filter"></i> Filters</div>
                        <div class="muted" style="font-size:12px;">Admin view only</div>
                      </div>

                      <form method="get" action="{html.escape(path, quote=True)}" class="filter-form">
                        <div class="filter-field">
                          <label>Month</label>
                          <input type="month" name="month" value="{html.escape(month_filter, quote=True)}">
                        </div>

                        <div class="filter-actions">
                          <button class="btn small" type="submit">
                            <i class="fa fa-filter"></i> Apply
                          </button>
                        </div>
                      </form>
                    </div>
                    """

                # list rows
                if role == "Employee":
                    rows = list_uploads_db(module_ctx, uploaded_by=session_email)
                elif role == "Admin":
                    rows = list_uploads_db(module_ctx, uploaded_by=None,
                                           month=(month_filter or None),
                                           name=(name_filter or None))
                else:
                    # Manager: no filters requested (can add later if needed)
                    rows = list_uploads_db(module_ctx, uploaded_by=None)

                uploads_html = ""
                for upload_id, stored_fn, original_fn, uploaded_at, module_val, up_email, up_uname in rows:
                    file_url = f"/{UPLOAD_DIR}/{stored_fn}"
                    badge = f"<span class='badge'>{html.escape(module_val or module_ctx)}</span>"

                    # Timesheet status -> used for Submitted badge in header
                    ts = get_timesheet_by_upload(upload_id) or {}
                    is_submitted = int(ts.get("submitted", 0)) == 1
                    submitted_at = ts.get("submitted_at", "") or ""

                    submitted_badge = ""
                    if is_submitted:
                        submitted_badge = (
                            f"<span class='badge lock-badge' title='{html.escape(submitted_at)}'>"
                            f"<i class='fa fa-lock'></i> Submitted"
                            f"</span>"
                        )

                    # Role-aware button (Delete for Admin, Recapture for Manager/Employee)
                    # Hidden if already submitted to avoid overlap and duplication
                    recapture_html = ""
                    if role == "Admin" or not is_submitted:
                        btn_text = "Delete" if role == "Admin" else "Recapture"
                        recapture_html = (
                            "<form method='post' action='/upload/delete' style='display:inline;margin:0' "
                            f"onsubmit=\"return confirm('{btn_text} this upload?');\">"
                            f"<input type='hidden' name='upload_id' value='{upload_id}'>"
                            f"<input type='hidden' name='return_to' value='{html.escape(path, quote=True)}'>"
                            f"<button class='btn danger small' type='submit'>{btn_text}</button>"
                            "</form>"
                        )

                    # Show button only in timesheet header (confirmation area) to avoid duplication
                    timesheet_html = build_timesheet_ui(upload_id, path, admin_delete_html=recapture_html)

                    # Main item header only shows Submitted badge to prevent overlap
                    right_header = f"<div class='upload-actions'>{submitted_badge}</div>"

                    # Optional uploader label (Admin only)
                    uploader_line = ""
                    if role == "Admin":
                        shown_name = up_uname or ""
                        shown_email = up_email or ""
                        if shown_name and shown_email:
                            uploader_line = f"<div class='muted'>Uploader: <b>{html.escape(shown_name)}</b> ({html.escape(shown_email)})</div>"
                        elif shown_email:
                            uploader_line = f"<div class='muted'>Uploader: <b>{html.escape(shown_email)}</b></div>"

                    # Build verification screenshot row for PPM uploads (shown as separate row below PPM)
                    verify_row_html = ""
                    if module_ctx == "PPM":
                        conn_v = get_conn()
                        cur_v = conn_v.cursor()
                        cur_v.execute("SELECT stored_filename, uploaded_at FROM uploads WHERE module='VERIFY' AND original_filename=?", (f"PPM_{upload_id}",))
                        v_row = cur_v.fetchone()
                        conn_v.close()
                        if v_row:
                            v_url = f"/{UPLOAD_DIR}/{v_row[0]}"
                            v_uploaded_at = v_row[1] if len(v_row) > 1 else ""
                            verify_row_html = (
                                "<div style='margin-left:28px; margin-top:6px; margin-bottom:16px; "
                                "padding:12px 16px; border-radius:10px; "
                                "background:linear-gradient(135deg, #f0fdf4 0%, #dcfce7 60%, #bbf7d0 100%); "
                                "border:1px solid #86efac; display:flex; align-items:center; gap:14px;'>"
                                
                                # Thumbnail
                                f"<a class='thumb-link' href='{html.escape(v_url, quote=True)}' "
                                f"data-url='{html.escape(v_url, quote=True)}' "
                                f"data-filename='Confirmation Screen Print' "
                                f"style='flex-shrink:0;'>"
                                f"<img src='{html.escape(v_url, quote=True)}' alt='confirmation' "
                                f"style='width:72px; height:48px; object-fit:cover; border-radius:6px; "
                                f"border:2px solid #4ade80; box-shadow:0 2px 8px rgba(0,0,0,0.08);'></a>"
                                
                                # Details
                                "<div style='flex:1; min-width:0;'>"
                                "<div style='display:flex; align-items:center; gap:8px; flex-wrap:wrap;'>"
                                "<span style='font-weight:700; font-size:13px; color:#14532d;'>Confirmation Screen Print</span>"
                                "<span style='display:inline-flex; align-items:center; gap:4px; "
                                "padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; "
                                "background:linear-gradient(135deg, #10b981, #059669); color:#fff; "
                                "letter-spacing:0.3px;'>"
                                "<i class='fa fa-shield-halved' style='font-size:10px;'></i> Confirmed</span>"
                                "</div>"
                                f"<div style='font-size:11px; color:#6b7280; margin-top:3px;'>"
                                f"Captured at (IST): {html.escape(format_dt_ist(v_uploaded_at))}</div>"
                                "</div>"
                                "</div>"
                            )

                    uploads_html += (
                        "<div class='upload-item'>"
                        f"<a class='thumb-link' href='{html.escape(file_url, quote=True)}' "
                        f"data-url='{html.escape(file_url, quote=True)}' "
                        f"data-filename='{html.escape(original_fn, quote=True)}'>"
                        f"<img src='{html.escape(file_url, quote=True)}' alt='screenshot'></a>"
                        "<div class='upload-meta'>"
                        "<div class='upload-head'>"
                        f"<div><strong>{html.escape(original_fn)}</strong> {badge}</div>"
                        f"{right_header}"
                        "</div>"
                        f"<div>Captured at (IST): {html.escape(format_dt_ist(uploaded_at))}</div>"
                        f"{uploader_line}"
                        f"{timesheet_html}"
                        "</div></div>"
                    )
                    uploads_html += verify_row_html

                uploads_section = uploads_html if uploads_html else "<div class='muted'>No captures yet.</div>"
                capture_ui = build_capture_ui(module_ctx)
                modal_ui = build_saved_preview_modal()
                ts_js = build_timesheet_js()

                content = (
                    "<div class='card' style='max-width:1000px;margin:0 auto'>"
                    f"{upload_error_alert_js}{capture_ui}"
                    f"{filter_ui}"
                    f"<h2 style='margin-top:20px'>{html.escape('Email' if module_ctx=='EMAIL' else module_ctx)} Captures</h2>"
                    f"{uploads_section}"
                    f"{modal_ui}"
                    f"{ts_js}"
                    "</div>"
                )
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
                return

            # Email settings (Admin/Manager)
            if path == "/email-settings":
                if role in ("Employee", "Manager"):
                    ce, _cx, _ch = get_feature_flags_by_email(session_email)
                    if not ce:
                        return self._msg("Access denied: you do not have permission to access Email Settings.", display_name)

                active_users = list_active_users_for_reminders()
                recip_opts = "<option value='' selected disabled>Select active user</option>"
                for uname, email_addr in active_users:
                    label = f"{uname} ({email_addr})" if uname else email_addr
                    recip_opts += f"<option value='{html.escape(email_addr, quote=True)}'>{html.escape(label)}</option>"

                form_html = (
                    "<div class='card' style='max-width:760px;margin:0 auto'>"
                    "<h2>Email Reminder Settings</h2>"
                    "<form method='post' action='/email-settings'>"
                    "<label>Recipient (Active Users)</label><select name='recipient' required>" + recip_opts + "</select>"
                    "<label>Subject</label><input type='text' name='subject' required>"
                    "<label>Message</label><textarea name='message' required></textarea>"
                    "<label>Scheduled date</label><input type='date' name='date' required>"
                    "<label>Scheduled time (IST)</label><input type='time' name='time' value='09:00' required>"
                    "<button class='btn' type='submit'>Save Reminder</button>"
                    "</form>"
                    "<div class='muted' style='margin-top:10px'>SMTP must be configured on server (SMTP_HOST/SMTP_PORT/SMTP_FROM and optionally SMTP_USER/SMTP_PASS).</div>"
                    "</div>"
                )

                rows = list_email_reminders()
                reminders_html = ""
                if rows:
                    reminders_html = (
                        "<div class='card' style='max-width:1100px;margin:18px auto'>"
                        "<h2>Saved Reminders</h2>"
                        "<table style='width:100%;border-collapse:collapse;margin-top:10px'>"
                        "<thead><tr style='text-align:left'>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>ID</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Recipient</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Subject</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Date</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Time</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Status</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Created at</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Actions</th>"
                        "</tr></thead><tbody>"
                    )
                    for rid, recip, subj, _msg_body, sched, sched_time, sent, sent_at, last_error, _created_by, created_at in rows:
                        status_txt = "Sent" if int(sent)==1 else "Pending"
                        if int(sent)==1 and sent_at:
                            status_txt += f" ({html.escape(sent_at)})"
                        if int(sent)!=1 and last_error:
                            status_txt += f" - {html.escape(last_error)}"
                        reminders_html += (
                            "<tr>"
                            f"<td style='padding:8px'>{rid}</td>"
                            f"<td style='padding:8px'>{html.escape(recip)}</td>"
                            f"<td style='padding:8px'>{html.escape(subj)}</td>"
                            f"<td style='padding:8px'>{html.escape(sched)}</td>"
                            f"<td style='padding:8px'>{html.escape(sched_time or '')}</td>"
                            f"<td style='padding:8px'>{status_txt}</td>"
                            f"<td style='padding:8px'>{html.escape(created_at)}</td>"
                            f"<td style='padding:8px'>"
                            f"<form method='get' action='/email-settings/edit' style='display:inline;margin-right:6px'>"
                            f"<input type='hidden' name='id' value='{rid}'>"
                            f"<button class='btn secondary small' type='submit'>Edit</button></form>"
                            f"<form method='post' action='/email-settings/delete' style='display:inline' onsubmit=\"return confirm('Delete this reminder?');\">"
                            f"<input type='hidden' name='id' value='{rid}'>"
                            f"<button class='btn danger small' type='submit'>Delete</button></form>"
                            f"</td>"
                            "</tr>"
                        )
                    reminders_html += "</tbody></table></div>"

                full = form_html + reminders_html
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(full, display_name, role, screen_access).encode("utf-8"))
                return

            if path == "/email-settings/edit":
                if role == "Employee":
                    return self._msg("Access denied: your role does not have permission to access Email Settings.", display_name)
                qs = parse_qs(urlparse(self.path).query)
                rid = (qs.get('id', [''])[0] or '').strip()
                try:
                    rid_int = int(rid)
                except Exception:
                    return self._msg("Invalid reminder ID.", display_name)
                row = get_email_reminder(rid_int)
                if not row:
                    return self._msg("Reminder not found.", display_name)
                _id, recipient, subject, message, sched_date, sched_time, sent, sent_at, last_error, _created_by, created_at = row

                active_users = list_active_users_for_reminders()
                recip_opts = "<option value='' disabled>Select active user</option>"
                for uname, email_addr in active_users:
                    label = f"{uname} ({email_addr})" if uname else email_addr
                    sel = " selected" if email_addr.lower() == (recipient or '').strip().lower() else ""
                    recip_opts += f"<option value='{html.escape(email_addr, quote=True)}'{sel}>{html.escape(label)}</option>"

                content = (
                    "<div class='card' style='max-width:760px;margin:0 auto'>"
                    "<h2>Edit Email Reminder</h2>"
                    "<form method='post' action='/email-settings/update'>"
                    f"<input type='hidden' name='id' value='{_id}'>"
                    "<label>Recipient (Active Users)</label>"
                    f"<select name='recipient' required>{recip_opts}</select>"
                    "<label>Subject</label>"
                    f"<input type='text' name='subject' value='{html.escape(subject or '', quote=True)}' required>"
                    "<label>Message</label>"
                    f"<textarea name='message' required>{html.escape(message or '')}</textarea>"
                    "<label>Scheduled date</label>"
                    f"<input type='date' name='date' value='{html.escape(sched_date or '', quote=True)}' required>"
                    "<label>Scheduled time (IST)</label>"
                    f"<input type='time' name='time' value='{html.escape((sched_time or '09:00'), quote=True)}' required>"
                    "<div class='controls'>"
                    "<button class='btn small' type='submit'><i class='fa fa-save'></i> Update</button>"
                    "<a class='btn secondary small' href='/email-settings'>Cancel</a>"
                    "</div>"
                    "</form></div>"
                )

                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode('utf-8'))
                return


            # Export Preview (Detailed) (Admin/Manager)

            if path == "/export/preview":

                if role in ("Employee", "Manager"):

                    _ce, cx, _ch = get_feature_flags_by_email(session_email)

                    if not cx:

                        return self._msg("Access denied: you do not have permission to access Export Data.", display_name)


                qs = parse_qs(urlparse(self.path).query)

                month = (qs.get('month', [''])[0] or '').strip()

                if not re.match(r'^\d{4}-\d{2}$', month):

                    month = utc_now().strftime('%Y-%m')

                resource = (qs.get('resource', [''])[0] or '').strip()

                try:

                    page = int((qs.get('page', ['1'])[0] or '1'))

                except Exception:

                    page = 1

                try:

                    ps = int((qs.get('ps', ['200'])[0] or '200'))

                except Exception:

                    ps = 200

                show_days = (qs.get('days', ['1'])[0] or '1').strip().lower() in ('1','true','yes','on')


                try:

                    rows = build_monthly_export_detail_rows(month)

                    if resource:

                        rlow = resource.lower()

                        rows = [r for r in rows if (rlow in (r.get('resource_name','').lower()) or rlow in (r.get('email','').lower()))]

                    preview_table = build_export_details_table_html(month, rows, page=page, page_size=ps, show_days=show_days,

                                                                  base_path='/export/preview', extra_qs={'resource': resource})

                    content = f"""<div class='card' style='max-width:1200px;margin:0 auto'>

                      <div style='display:flex;align-items:flex-end;justify-content:space-between;gap:12px;flex-wrap:wrap'>

                        <div>

                          <h2 style='margin:0'>Export Preview — Details</h2>

                          <div class='muted'>Month: <b>{html.escape(month)}</b></div>

                        </div>

                        <a class='btn outline small' href='/export?month={html.escape(month, quote=True)}&resource={html.escape(resource, quote=True)}'><i class='fa fa-arrow-left'></i> Back</a>

                      </div>

                      {preview_table}

                    </div>"""

                except Exception as ex:

                    content = (

                        "<div class='card' style='max-width:900px;margin:0 auto'>"

                        "<h2>Export Preview</h2>"

                        f"<p class='muted'>{html.escape(str(ex))}</p>"

                        "<a class='btn outline small' href='/export'><i class='fa fa-arrow-left'></i> Back</a>"

                        "</div>"

                    )


                self.send_response(200)

                self.send_header("Content-type", "text/html; charset=utf-8")

                self.end_headers()

                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))

                return



            # Export Data (Admin/Manager)

            if path == "/export":

                if role in ("Employee", "Manager"):
                    _ce, cx, _ch = get_feature_flags_by_email(session_email)
                    if not cx:
                        return self._msg("Access denied: you do not have permission to access Export Data.", display_name)

                qs = parse_qs(urlparse(self.path).query)
                filters, persist_qs = parse_export_filters_from_qs(qs)

                today = utc_now().date()

                default_from = today.replace(day=1)
                fd = (filters.get('from_date') or '').strip()
                td = (filters.get('to_date') or '').strip()
                try:
                    from_d = datetime.date.fromisoformat(fd) if fd else default_from
                except Exception:
                    from_d = default_from
                try:
                    to_d = datetime.date.fromisoformat(td) if td else today
                except Exception:
                    to_d = today
                if to_d < from_d:
                    from_d, to_d = to_d, from_d
                filters['from_date'] = from_d.isoformat()
                filters['to_date'] = to_d.isoformat()

                try:
                    page = int((_qs_first(qs, 'page', '1') or '1'))
                except Exception:
                    page = 1
                try:
                    ps = int((_qs_first(qs, 'ps', '200') or '200'))
                except Exception:
                    ps = 200
                show_days = (_qs_first(qs, 'days', '1') or '1').strip().lower() in ('1','true','yes','on')

                try:
                    _all_rows = build_monthly_export_detail_rows_filtered('', filters)
                    rows = [r for r in _all_rows if (r.get('module') or '').strip().upper() == 'NTT']
                except Exception:
                    rows = []

                preview_table = build_export_details_table_html(
                    filters.get('from_date',''), filters.get('to_date',''), rows,
                    page=page,
                    page_size=ps,
                    show_days=show_days,
                    base_path='/export',
                    extra_qs=persist_qs
                )

                filter_ui = build_export_filter_ui_html(filters, ps=ps, show_days=show_days)

                content = f"""
                <div class='card' style='max-width:100%;margin:0 auto'>
                  <div style='display:flex;align-items:flex-end;justify-content:space-between;gap:12px;flex-wrap:wrap'>
                    <div>
                      <h2 style='margin:0'>Export Data</h2>
                      <div class='muted'>Preview is shown by default. Filters include uploader profile fields and custom user fields.</div>
                    </div>
                  </div>

                  {filter_ui}

                  {preview_table}
                </div>
                """
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
                return
            # Help (all roles) - DB-backed
            if path == "/help":
                help_html = (get_help_html() or "").strip()
                if not help_html:
                    help_html = "<div class='card' style='max-width:960px;margin:0 auto'><h2>Help</h2><p class='muted'>No help content available.</p></div>"

                edit_btn = ""
                if role == "Admin":
                    edit_btn = (
                        "<div style='max-width:960px;margin:8px auto 0;display:flex;justify-content:flex-end'>"
                        "<a class='btn outline small' href='/help/edit'><i class='fa fa-pen-to-square'></i> Edit Help</a>"
                        "</div>"
                    )

                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(edit_btn + help_html, display_name, role, screen_access).encode("utf-8"))
                return

            # Help editor (Admin only) - WYSIWYG (Quill) with robust Save
            if path == "/help/edit":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can edit the Help page.", display_name)

                import json
                current = get_help_html() or ""
                current_js = json.dumps(current)

                form_html = """
                <div class='card' style='max-width:1100px;margin:0 auto'>
                  <h2><i class='fa fa-pen-to-square'></i> Edit Help</h2>
                  <p class='muted'>Edit like a Word document. Click <b>Save</b> to publish.</p>

                  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/quill@1.3.7/dist/quill.snow.css">

                  <form id='helpForm' method='post' action='/help/edit'>
                    <div id="helpToolbar">
                      <span class="ql-formats">
                        <select class="ql-header">
                          <option selected></option>
                          <option value="1"></option>
                          <option value="2"></option>
                          <option value="3"></option>
                        </select>
                      </span>
                      <span class="ql-formats">
                        <button type="button" class="ql-bold"></button>
                        <button type="button" class="ql-italic"></button>
                        <button type="button" class="ql-underline"></button>
                        <button type="button" class="ql-strike"></button>
                      </span>
                      <span class="ql-formats">
                        <button type="button" class="ql-list" value="ordered"></button>
                        <button type="button" class="ql-list" value="bullet"></button>
                      </span>
                      <span class="ql-formats">
                        <select class="ql-align"></select>
                      </span>
                      <span class="ql-formats">
                        <button type="button" class="ql-link"></button>
                        <button type="button" class="ql-clean"></button>
                      </span>
                    </div>

                    <div id="helpEditor" style="height:420px;background:#fff;border-radius:10px"></div>
                    <!-- NOT required, because browser validation can block submit before JS copies content -->
                    <textarea id="helpHtml" name="html" style="display:none"></textarea>

                    <div class="controls" style="margin-top:12px">
                      <button class="btn small" type="submit"><i class="fa fa-save"></i> Save</button>
                      <a class="btn secondary small" href="/help">Cancel</a>
                    </div>

                    <div id='helpWarn' class='muted' style='margin-top:10px;font-size:12px;display:none'>
                      Note: Quill editor resources did not load (network/CDN blocked). Fallback editor is active.
                    </div>
                  </form>

                  <script src="https://cdn.jsdelivr.net/npm/quill@1.3.7/dist/quill.min.js"></script>
                  <script>
                    (function(){
                      const existingHtml = """ + current_js + """ || '';
                      const editorDiv = document.getElementById('helpEditor');
                      const hidden = document.getElementById('helpHtml');
                      const form = document.getElementById('helpForm');
                      const warn = document.getElementById('helpWarn');

                      // Initialize hidden value so submit is never empty even if JS fails later
                      hidden.value = existingHtml;

                      // Always allow typing even if Quill fails
                      editorDiv.setAttribute('contenteditable', 'true');

                      let quill = null;
                      try {
                        if (window.Quill) {
                          quill = new Quill('#helpEditor', {
                            theme: 'snow',
                            modules: { toolbar: '#helpToolbar' }
                          });
                        }
                      } catch(e) {
                        quill = null;
                      }

                      if (quill) {
                        quill.root.innerHTML = existingHtml;
                      } else {
                        editorDiv.innerHTML = existingHtml;
                        if (warn) warn.style.display = 'block';
                      }

                      // Copy HTML into hidden textarea before submit
                      form.addEventListener('submit', function(){
                        try {
                          hidden.value = quill ? (quill.root.innerHTML || '') : (editorDiv.innerHTML || '');
                        } catch(e) {
                          hidden.value = editorDiv.innerHTML || '';
                        }
                      });
                    })();
                  </script>
                </div>
                """

                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(form_html, display_name, role, screen_access).encode("utf-8"))
                return
            # Not found
            self.send_response(404)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(render_page(
                "<div class='card' style='text-align:center'><h2>404 Not Found</h2>"
                "<p class='muted'>The page you requested does not exist.</p></div>",
                display_name, role, screen_access
            ).encode("utf-8"))

        except Exception as e:
            logging.exception("Unhandled exception in do_GET")
            try:
                if DEBUG:
                    content = self._render_traceback_page(e)
                    self.send_response(500)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.end_headers()
                    self.wfile.write(render_page(content, "Visitor").encode("utf-8"))
                    return
                self.send_response(500)
                self.end_headers()
                self.wfile.write(b"Server error")
            except Exception:
                logging.exception("Failed to send 500 response")

    # -------------------------
    # POST
    # -------------------------
    def do_POST(self):
        try:
            path = self._normalize_path(self.path)
            session_email = get_email_from_request(self)
            display_name = get_display_name_by_email(session_email) if session_email else "Visitor"
            role = get_role_by_email(session_email) if session_email else None
            screen_access = get_screen_access_by_email(session_email) if session_email else "BOTH"

            ctype = self.headers.get("Content-Type", "")
            length = int(self.headers.get("Content-Length", "0") or "0")
            body = self.rfile.read(length) if length else b""

            form = {}
            if "multipart/form-data" in ctype:
                m = re.search(r'boundary=(?P<b>[^;]+)', ctype)
                if not m:
                    return self._msg("Malformed multipart/form-data request.", "Visitor")

                boundary_raw = m.group("b").strip()
                if boundary_raw.startswith('"') and boundary_raw.endswith('"'):
                    boundary_raw = boundary_raw[1:-1]
                boundary = boundary_raw.encode("utf-8", errors="ignore")
                form = parse_multipart(body, boundary)
            else:
                try:
                    parsed = parse_qs(body.decode("utf-8", errors="ignore"), keep_blank_values=True)
                    form = {k: (v[0] if isinstance(v, list) and v else "") for k, v in parsed.items()}
                except Exception:
                    form = {}

            # LOGIN
            if path == "/login":
                email_in = (form.get("email") or "").strip().lower()
                password = (form.get("password") or "").strip()

                if not email_in or not password:
                    return self._msg("Please enter email and password.", "Visitor")

                row = get_user_db_by_email(email_in)
                if not row:
                    self._redirect_login_popup("Invalid email or password.")
                    return

                db_password = row[3]
                db_status = row[5]
                if db_status != "Active":
                    self._redirect_login_popup("Account is inactive. Please contact administrator.")
                    return

                if db_password == password:
                    sid = make_session(email_in)
                    self.send_response(303)
                    self.send_header("Location", "/")
                    self.send_header("Set-Cookie", f"session_id={sid}; Path=/; HttpOnly")
                    self.end_headers()
                    return

                self._redirect_login_popup("Invalid email or password.")
                return

            if not session_email:
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            # HARD BLOCK: Manager no access to user management POST
            if role == "Manager" and path.startswith("/user-management"):
                return self._msg("Access denied: your role does not have permission to access User Management.", display_name)

            


            # UPDATE CUSTOM FIELD (Admin only)
            if path == "/user-management/create-fields/update":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can manage custom fields.", display_name)
                try:
                    field_id = int((form.get('field_id') or '0') or 0)
                    label = (form.get('label') or '').strip()
                    required = (str(form.get('required') or '0').strip().lower() in ('1','true','yes','on'))
                    active = (str(form.get('active') or '0').strip().lower() in ('1','true','yes','on'))
                    options_raw = (form.get('options') or '').strip()
                    options_list = [x.strip() for x in options_raw.split(',') if x.strip()] if options_raw else None
                    update_custom_field(field_id, label, 1 if required else 0, options_list, 1 if active else 0)
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Field updated successfully.'))
                    self.end_headers()
                    return
                except Exception as ex:
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Error: ' + str(ex)))
                    self.end_headers()
                    return

            # DELETE CUSTOM FIELD (Admin only)
            if path == "/user-management/create-fields/delete":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can manage custom fields.", display_name)
                try:
                    field_id = int((form.get('field_id') or '0') or 0)
                    delete_custom_field(field_id)
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Field deleted successfully.'))
                    self.end_headers()
                    return
                except Exception as ex:
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Error: ' + str(ex)))
                    self.end_headers()
                    return

            # CREATE CUSTOM FIELD (Admin only)
            if path == "/user-management/create-fields":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can manage custom fields.", display_name)
                try:
                    field_key = (form.get('field_key') or '').strip()
                    label = (form.get('label') or '').strip()
                    field_type = (form.get('field_type') or 'TEXT').strip().upper()
                    required = (str(form.get('required') or '0').strip().lower() in ('1','true','yes','on'))
                    options_raw = (form.get('options') or '').strip()
                    options_list = [x.strip() for x in options_raw.split(',') if x.strip()] if options_raw else None
                    create_custom_field(field_key, label, field_type, 1 if required else 0, options_list)
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Field created successfully.'))
                    self.end_headers()
                    return
                except Exception as ex:
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Error: ' + str(ex)))
                    self.end_headers()
                    return

            # DEACTIVATE CUSTOM FIELD (Admin only)
            if path == "/user-management/create-fields/deactivate":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can manage custom fields.", display_name)
                try:
                    fid = int(form.get('field_id') or 0)
                    deactivate_custom_field(fid)
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Field deactivated.'))
                    self.end_headers()
                    return
                except Exception as ex:
                    self.send_response(303)
                    self.send_header('Location', '/user-management/create-fields?msg=' + quote('Error: ' + str(ex)))
                    self.end_headers()
                    return
# CREATE USER
            if path == "/user-management/create":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can create users.", display_name)

                required = ("username", "email", "password", "role", "status")
                if not all(form.get(k) for k in required):
                    return self._msg("Missing required fields for Create User.", display_name)

                new_uid = create_user_db_return_id(
                    form["username"], form["email"], form["password"],
                    form["role"], form["status"],
                    form.get("screen_access", "BOTH"),
                    form.get("can_email", 0), form.get("can_export", 0), form.get("can_help", 1)
                )

                if new_uid:
                    try:
                        validate_and_save_custom_fields_for_user(int(new_uid), form, active_only=True, name_prefix='cf_')
                    except Exception as ex:
                        return self._msg("User created, but custom fields error: " + str(ex), display_name)
                    return self._msg("User created successfully.", display_name)

                return self._msg("Username or email already exists.", display_name)


            # UPDATE USER
            if path == "/user-management/update":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can update users.", display_name)

                user_id = form.get("user_id")
                if not user_id:
                    return self._msg("User ID is required for update.", display_name)

                updates = {k: form.get(k) for k in ("username", "email", "role", "status", "screen_access", "can_email", "can_export", "can_help") if form.get(k) is not None}
                if form.get("password"):
                    updates["password"] = form.get("password")

                if not updates:
                    return self._msg("No changes provided.", display_name)
                ok = update_user_db(int(user_id), updates)

                if ok:
                    try:
                        validate_and_save_custom_fields_for_user(int(user_id), form, active_only=True, name_prefix='cf_')
                    except Exception as ex:
                        return self._msg('User updated, but custom fields error: ' + str(ex), display_name)
                    self.send_response(303)

                    self.send_header("Location", "/user-management/list")
                    self.end_headers()
                    return
                return self._msg("Update failed: username or email may already exist.", display_name)

            
            # UPDATE SCREEN ACCESS (Admin only)
            
            # UPDATE ACCESS (Admin only): Screen module + Email/Export/Help
            if path == "/user-management/update-access":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can update user access.", display_name)

                user_id = (form.get("user_id") or '').strip()
                screen_access_val = (form.get("screen_access") or '').strip()
                return_to = (form.get("return_to") or '/user-management/list').strip()

                try:
                    uid_int = int(user_id)
                except Exception:
                    return self._msg("Invalid user ID for access update.", display_name)

                row = get_user_db_by_id(uid_int)
                if not row:
                    return self._msg("User not found.", display_name)

                _uid, _uname, _email, rrole, _status, _sa, _ce, _cx, _ch = row

                sa = normalize_screen_access(screen_access_val)
                if (rrole or '').strip() == 'Admin':
                    sa = 'BOTH'

                updates = {
                    'screen_access': sa,
                    'can_email': 1 if (form.get('can_email') in ('on','1','true','True','yes')) else 0,
                    'can_export': 1 if (form.get('can_export') in ('on','1','true','True','yes')) else 0,
                    'can_help': (0 if (rrole or '').strip()=='Employee' else (1 if (form.get('can_help') in ('on','1','true','True','yes')) else 0)),
                }

                ok = update_user_db(uid_int, updates)
                if not ok:
                    return self._msg("Update failed.", display_name)

                self.send_response(303)
                self.send_header('Location', return_to)
                self.end_headers()
                return



# DELETE USER
            if path == "/user-management/delete":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can delete users.", display_name)
                user_id = form.get("user_id")
                if not user_id:
                    return self._msg("User ID is required to delete.", display_name)
                delete_user_db(int(user_id))
                self.send_response(303)
                self.send_header("Location", "/user-management/list")
                self.end_headers()
                return

            # RECEIVE captured image
            if path in ("/upload/PPM", "/upload/NTT", "/upload/EMAIL", "/upload/VERIFY"):
                f = form.get("file")
                if not f or not isinstance(f, dict) or not f.get("filename"):
                    return self._msg("No captured image received.", display_name)

                module_val = path.rsplit('/', 1)[-1]
                upload_err = ""

                if role in ("Employee", "Manager"):
                    if module_val == "PPM" and screen_access not in ("PPM", "BOTH"):
                        return self._msg("Access denied: you do not have permission to upload to PPM.", display_name)
                    if module_val == "NTT" and screen_access not in ("NTT", "BOTH"):
                        return self._msg("Access denied: you do not have permission to upload to NTT.", display_name)
                    if module_val == "EMAIL" and screen_access != "EMAIL":
                        return self._msg("Access denied: you do not have permission to upload to EMAIL.", display_name)
                    if module_val == "VERIFY" and screen_access not in ("PPM", "BOTH"):
                        return self._msg("Access denied: you do not have permission to upload verification screenprints.", display_name)

                user_disp = get_display_name_by_email(session_email) or (session_email or "user")
                if "@" in user_disp:
                    user_disp = user_disp.split("@", 1)[0]
                user_safe = _safe_name_part(user_disp, default="user")

                now = datetime.datetime.now()
                date_str = now.strftime("%Y%m%d")
                time_str = now.strftime("%H%M%S")

                if module_val == "VERIFY":
                    parent_id = form.get("parent_upload_id")
                    if isinstance(parent_id, list):
                        parent_id = parent_id[0] if parent_id else ""
                    if parent_id:
                        original_fn = f"PPM_{parent_id}"
                    else:
                        original_fn = f"{module_val}_{user_safe}_{date_str}_{time_str}.png"
                else:
                    original_fn = f"{module_val}_{user_safe}_{date_str}_{time_str}.png"
                base, ext = os.path.splitext(original_fn)
                ext = ext or ".png"

                timestamp = now.strftime("%Y%m%d%H%M%S")
                safe_base = re.sub(r"[^a-zA-Z0-9._-]+", "_", base)[:80] or "capture"
                stored_fn = f"{safe_base}_{timestamp}{ext}".replace(" ", "_")
                save_path = os.path.join(UPLOAD_DIR, stored_fn)

                counter = 0
                while os.path.exists(save_path):
                    counter += 1
                    stored_fn = f"{safe_base}_{timestamp}_{counter}{ext}"
                    save_path = os.path.join(UPLOAD_DIR, stored_fn)

                with open(save_path, "wb") as out:
                    out.write(f["content"])

                upload_id = add_upload_db(stored_fn, original_fn, module_val, session_email)
                print(f"DEBUG: UPLOAD RECEIVED: module={module_val} id={upload_id}", flush=True)
                logging.info("===== UPLOAD RECEIVED: module=%s upload_id=%s path=%s =====", module_val, upload_id, save_path)

                # PPM: OCR extract Total-row hours, save, and auto-lock (disable edits)
                if module_val == "PPM":
                    try:
                        parsed = extract_ppm_hours_from_screenshot_total_row(save_path)
                        if parsed:
                            ws = parsed.get("week_start") or _week_start_from_iso_datetime(utc_now_str(sep=" ", timespec="seconds"))
                            sun = parsed.get("sun", 0.0)
                            mon = parsed.get("mon", 0.0)
                            tue = parsed.get("tue", 0.0)
                            wed = parsed.get("wed", 0.0)
                            thu = parsed.get("thu", 0.0)
                            fri = parsed.get("fri", 0.0)
                            sat = parsed.get("sat", 0.0)
                            rem = parsed.get("remaining_planned")
                            upsert_timesheet(upload_id, ws, mon, tue, wed, thu, fri, sat, sun, remaining_planned=rem)
                            logging.info("PPM OCR autofill ok upload_id=%s week_start=%s", upload_id, ws)
                    except Exception:
                        logging.exception("PPM OCR extraction failed (non-fatal).")


                # NTT: OCR extract Entered-column hours (SUM per day), save, and auto-lock (disable edits)
                if module_val == "NTT":
                    try:
                        print(f"DEBUG: NTT OCR starting for id={upload_id}", flush=True)
                        logging.info("NTT OCR starting for upload_id=%s path=%s", upload_id, save_path)
                        parsed = extract_ntt_hours_from_screenshot_entered_column(save_path)
                        print(f"DEBUG: NTT OCR parsed result: {parsed}", flush=True)
                        logging.info("NTT OCR parsed result: %s", parsed)
                        if parsed:
                            ws = parsed.get("week_start") or _week_start_from_iso_datetime(utc_now_str(sep=" ", timespec="seconds"))
                            sun = parsed.get("sun", 0.0)
                            mon = parsed.get("mon", 0.0)
                            tue = parsed.get("tue", 0.0)
                            wed = parsed.get("wed", 0.0)
                            thu = parsed.get("thu", 0.0)
                            fri = parsed.get("fri", 0.0)
                            sat = parsed.get("sat", 0.0)
                            logging.info("NTT OCR upsert: upload_id=%s ws=%s mon=%s tue=%s wed=%s thu=%s fri=%s sat=%s sun=%s",
                                         upload_id, ws, mon, tue, wed, thu, fri, sat, sun)
                            upsert_timesheet(upload_id, ws, mon, tue, wed, thu, fri, sat, sun, has_draft=(1 if parsed.get('has_draft') else 0))
                            found_days = int(parsed.get('_found_days') or 0)
                            # Both PPM & NTT now auto-extract and lock hours to ensure data integrity.
                            try:
                                owner = (session_email or '').strip().lower()
                                ppm = get_ppm_hours_for_user_week(owner, ws)
                                if ppm and float(ppm[7] or 0) > 0:
                                    ppm_hours = (ppm[0], ppm[1], ppm[2], ppm[3], ppm[4], ppm[5], ppm[6])
                                    ntt_hours = (float(sun), float(mon), float(tue), float(wed), float(thu), float(fri), float(sat))
                                    if not timesheet_hours_equal(ppm_hours, ntt_hours):
                                        upload_err = f"NTT hours do not match PPM for week {ws}. Please re-capture the NTT screenshot."
                                
                                # Draft detection error
                                if parsed.get('has_draft'):
                                    upload_err = "NTT status is in Draft mode. Please submit it in NTT portal and re-capture the screenshot."

                            except Exception:
                                logging.exception('NTT OCR compare failed (non-fatal).')
                            logging.info("NTT OCR autofill ok upload_id=%s week_start=%s found_days=%s", upload_id, ws, found_days)
                        else:
                            logging.warning("NTT OCR returned None for upload_id=%s", upload_id)
                    except Exception:
                        logging.exception("NTT OCR extraction failed (non-fatal).")


                # VERIFY uploads: return 200 OK (JS fetch handles reload)
                if module_val == "VERIFY":
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(b'{"ok":true}')
                    return

                self.send_response(303)
                loc = f"/upload/{module_val}"
                if upload_err:
                    loc = loc + "?err=" + quote(str(upload_err))
                self.send_header("Location", loc)
                self.end_headers()
                return

            # RECAPTURE/DELETE CAPTURE (Admin or Owner)
            if path == "/upload/delete":
                upload_id = form.get("upload_id")
                return_to = normalize_return_to(form.get("return_to") or "/upload/PPM", role, session_email)

                if not upload_id:
                    term = "delete" if role == "Admin" else "recapture"
                    return self._msg(f"Upload ID is required to {term}.", display_name)

                uid_int = int(upload_id)
                owner = get_upload_owner(uid_int)
                if role != "Admin" and (owner or "").strip().lower() != (session_email or "").strip().lower():
                    return self._msg("Access denied: you can only recapture your own uploads.", display_name)

                delete_upload(uid_int)
                self.send_response(303)
                self.send_header("Location", return_to)
                self.end_headers()
                return

            # Helper: parse and normalize timesheet payload
            def parse_timesheet_payload():
                week_start = (form.get("week_start") or "").strip()
                try:
                    ws_date = datetime.date.fromisoformat(week_start)
                except Exception:
                    ws_date = datetime.date.today()
                ws_date = _sunday_of_date(ws_date)
                week_start_norm = ws_date.isoformat()

                def hours(name):
                    val = (form.get(name) or "").strip()
                    try:
                        x = float(val)
                    except Exception:
                        x = 0.0
                    x = max(0.0, min(24.0, x))
                    x = round(x * 4) / 4.0
                    return x

                sun = hours("sun")
                mon = hours("mon")
                tue = hours("tue")
                wed = hours("wed")
                thu = hours("thu")
                fri = hours("fri")
                sat = hours("sat")
                total = sun + mon + tue + wed + thu + fri + sat

                return week_start_norm, mon, tue, wed, thu, fri, sat, sun, total

            # TIMESHEET SAVE
            if path == "/timesheet/save":
                upload_id = form.get("upload_id")
                return_to = normalize_return_to(form.get("return_to") or "/upload/PPM", role, session_email)

                if not upload_id:
                    return self._msg("Upload ID is required to save timesheet.", display_name)

                week_start, mon, tue, wed, thu, fri, sat, sun, _total = parse_timesheet_payload()
                # Block manual edits for PPM timesheets (auto-filled from screenshot)
                if get_upload_module(int(upload_id)) == "PPM":
                    return self._msg("PPM timesheets are auto-filled and cannot be edited.", display_name)


                try:
                    upsert_timesheet(int(upload_id), week_start, mon, tue, wed, thu, fri, sat, sun)
                except ValueError as ve:
                    return self._msg(str(ve), display_name)

                self.send_response(303)
                self.send_header("Location", return_to)
                self.end_headers()
                return

            # TIMESHEET SUBMIT
            if path == "/timesheet/submit":
                upload_id = form.get("upload_id")
                return_to = normalize_return_to(form.get("return_to") or "/upload/PPM", role, session_email)

                if not upload_id:
                    return self._msg("Upload ID is required to submit timesheet.", display_name)

                uid = int(upload_id)

                if is_timesheet_submitted(uid):
                    return self._msg("Already submitted: this timesheet is locked.", display_name)

                # Block if Draft
                ts_info = get_timesheet_by_upload(uid)
                if ts_info and get_upload_module(uid) == "NTT" and int(ts_info.get("has_draft", 0)) == 1:
                    return self._msg("NTT status is in Draft. Please submit the timesheet in NTT portal and re-capture the screenshot.", display_name)

                week_start, mon, tue, wed, thu, fri, sat, sun, total = parse_timesheet_payload()

                # verification screenprint check for PPM
                if get_upload_module(uid) == "PPM":
                    conn = get_conn()
                    cur = conn.cursor()
                    cur.execute("SELECT id FROM uploads WHERE module='VERIFY' AND original_filename=?", (f"PPM_{uid}",))
                    v_row = cur.fetchone()
                    conn.close()
                    if not v_row:
                        return self._msg("Please attach verification screenprint and click on submit.", display_name)

                owner = get_upload_owner(uid) or (session_email or "").strip().lower()

                if get_upload_module(uid) == "NTT":
                    latest_ppm = get_ppm_hours_for_user_week(owner, week_start)
                    ntt_candidate = (sun, mon, tue, wed, thu, fri, sat)
                    sep = "&" if "?" in return_to else "?"
                    
                    if not latest_ppm:
                        loc = return_to + sep + "err=" + quote("PPM and NTT do not match with correct hours, date and month. Please re-submit NTT timesheet.")
                        self.send_response(303)
                        self.send_header("Location", loc)
                        self.end_headers()
                        return
                    
                    ppm_hours = latest_ppm[:7]
                    
                    if not timesheet_hours_equal(ppm_hours, ntt_candidate):
                        loc = return_to + sep + "err=" + quote("PPM and NTT do not match with correct hours, date and month. Please re-submit NTT timesheet.")
                        self.send_response(303)
                        self.send_header("Location", loc)
                        self.end_headers()
                        return

                try:
                    submit_week_group(uid, owner, week_start, mon, tue, wed, thu, fri, sat, sun)
                except ValueError as ve:
                    sep = "&" if "?" in return_to else "?"
                    if get_upload_module(uid) == "NTT" and ("Entered hours in PPM and NTT must match" in str(ve) or "Cannot submit NTT: Hours must match" in str(ve)):
                        loc = return_to + sep + "err=" + quote("PPM and NTT do not match with correct hours, date and month. Please re-submit NTT timesheet.")
                    else:
                        loc = return_to + sep + "err=" + quote(str(ve))
                    
                    self.send_response(303)
                    self.send_header("Location", loc)
                    self.end_headers()
                    return

                if get_upload_module(uid) == "NTT":
                    sep = "&" if "?" in return_to else "?"
                    loc = return_to + sep + "err=" + quote("Submitted successfully.")
                    self.send_response(303)
                    self.send_header("Location", loc)
                    self.end_headers()
                    return

                self.send_response(303)
                self.send_header("Location", return_to)
                self.end_headers()
                return

            # EMAIL SETTINGS
            if path == "/email-settings":
                if role in ("Employee", "Manager"):
                    ce, _cx, _ch = get_feature_flags_by_email(session_email)
                    if not ce:
                        return self._msg("Access denied: you do not have permission to access Email Settings.", display_name)
                recipient = (form.get("recipient") or "").strip().lower()
                subject = (form.get("subject") or "").strip()
                message = (form.get("message") or "").strip()
                date_field = (form.get("date") or "").strip()
                time_field = (form.get("time") or "").strip()
                if not (recipient and subject and message and date_field and time_field):
                    return self._msg("All fields are required for email reminder.", display_name)
                try:
                    datetime.date.fromisoformat(date_field)
                except Exception:
                    return self._msg("Invalid date format.", display_name)
                if not re.match(r'^\d{2}:\d{2}$', time_field):
                    return self._msg("Invalid time format. Use HH:MM.", display_name)
                if not is_active_user_email(recipient):
                    return self._msg("Recipient must be an Active user from the user list.", display_name)
                add_email_reminder(recipient, subject, message, date_field, time_field, created_by=session_email)
                self.send_response(303)
                self.send_header("Location", "/email-settings")
                self.end_headers()
                return

            if path == "/email-settings/update":
                if role in ("Employee", "Manager"):
                    ce, _cx, _ch = get_feature_flags_by_email(session_email)
                    if not ce:
                        return self._msg("Access denied: you do not have permission to access Email Settings.", display_name)
                rid = (form.get('id') or '').strip()
                recipient = (form.get('recipient') or '').strip().lower()
                subject = (form.get('subject') or '').strip()
                message = (form.get('message') or '').strip()
                date_field = (form.get('date') or '').strip()
                time_field = (form.get('time') or '').strip()
                try:
                    rid_int = int(rid)
                except Exception:
                    return self._msg("Invalid reminder ID.", display_name)
                if not (recipient and subject and message and date_field and time_field):
                    return self._msg("All fields are required for updating the reminder.", display_name)
                try:
                    datetime.date.fromisoformat(date_field)
                except Exception:
                    return self._msg("Invalid date format.", display_name)
                if not re.match(r'^\d{2}:\d{2}$', time_field):
                    return self._msg("Invalid time format. Use HH:MM.", display_name)
                if not is_active_user_email(recipient):
                    return self._msg("Recipient must be an Active user from the user list.", display_name)
                if not get_email_reminder(rid_int):
                    return self._msg("Reminder not found.", display_name)
                update_email_reminder(rid_int, recipient, subject, message, date_field, time_field)
                self.send_response(303)
                self.send_header('Location', '/email-settings')
                self.end_headers()
                return

            if path == "/email-settings/delete":
                if role in ("Employee", "Manager"):
                    ce, _cx, _ch = get_feature_flags_by_email(session_email)
                    if not ce:
                        return self._msg("Access denied: you do not have permission to access Email Settings.", display_name)
                rid = (form.get('id') or '').strip()
                try:
                    rid_int = int(rid)
                except Exception:
                    return self._msg("Invalid reminder ID.", display_name)
                delete_email_reminder(rid_int)
                self.send_response(303)
                self.send_header('Location', '/email-settings')
                self.end_headers()
                return

            # EXPORT (Admin/Manager): date-range Excel download (Uploaded Date)

            if path == "/export":

                # Employees/Managers require Export permission; Admin allowed by default
                if role in ("Employee", "Manager"):
                    _ce, cx, _ch = get_feature_flags_by_email(session_email)
                    if not cx:
                        return self._msg("Access denied: you do not have permission to access Export Data.", display_name)

                from_date = (form.get("from_date") or "").strip()
                to_date = (form.get("to_date") or "").strip()
                try:
                    fd = datetime.date.fromisoformat(from_date)
                    td = datetime.date.fromisoformat(to_date)
                except Exception:
                    return self._msg("Invalid date range. Please select From Date and To Date (YYYY-MM-DD).", display_name)
                if td < fd:
                    return self._msg("Invalid date range: To Date cannot be earlier than From Date.", display_name)

                # Rebuild a parse_qs-like dict so we can reuse the same filter parser
                qs_like = {k: [v] for k, v in (form or {}).items() if isinstance(v, str)}
                filters, _persist = parse_export_filters_from_qs(qs_like)
                filters["from_date"] = fd.isoformat()
                filters["to_date"] = td.isoformat()

                try:
                    xlsx_bytes = generate_monthly_export_xlsx_filtered("", filters)
                except Exception as ex:
                    return self._msg(f"Export failed: {html.escape(str(ex))}", display_name)
                filename = f"{session_email}_{filters['from_date']}_to_{filters['to_date']}.xlsx"
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(len(xlsx_bytes)))
                self.end_headers()
                self.wfile.write(xlsx_bytes)
                return

            # HELP EDIT SAVE (Admin only)
            if path == "/help/edit":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can edit the Help page.", display_name)
                new_html = basic_sanitize_html((form.get("html") or "").strip())
                set_help_html(new_html)
                self.send_response(303)
                self.send_header("Location", "/help")
                self.end_headers()
                return

            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"Not Found")

        except Exception as e:
            logging.exception("Unhandled exception in do_POST")
            if DEBUG:
                content = self._render_traceback_page(e)
                self.send_response(500)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, "Visitor").encode("utf-8"))
            else:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(b"Server error")


# -------------------------
# Run server
# -------------------------
if __name__ == "__main__":
    init_db()
    logging.info("Starting server...")
    logging.info(f"URL: http://{HOST}:{PORT}")
    try:
        HTTPServer((HOST, PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        logging.info("Server stopped by user")
    except Exception:
        logging.exception("Server crashed")
