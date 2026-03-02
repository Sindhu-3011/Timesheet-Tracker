#!/usr/bin/env python3
"""
server_ppm_ntt_only.py

Single-file HTTP server with:
- Login by email + session cookie
- User management (Admin only)
- Screen Capture module with separate PPM and NTT pages ONLY
- Saved captures show thumbnails; clicking opens same-page modal preview
- Timesheet per saved screenprint (upload):
  * Week-wise (Sun–Sat) hours + week_start (Sunday)
  * Total computed in UI, recalculated/validated server-side
  * Submit (Lock) per timesheet
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
- Managers/Admin always BOTH (Option A)
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
from urllib.parse import parse_qs, urlparse
import os
import sqlite3
import html
import datetime
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
    import easyocr
    _EASYOCR_READER = None
except Exception:
    easyocr = None

# -------------------------
# Configuration
# -------------------------
DEBUG = True  # Set False in production
HOST = "127.0.0.1"
PORT = 8000
UPLOAD_DIR = "uploads"
DB_FILE = "users.db"

os.makedirs(UPLOAD_DIR, exist_ok=True)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


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
    return val if val in ("PPM", "BOTH", "EMAIL_NTT") else "BOTH"


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
    ensure_timesheets_verification_column()
    ensure_users_screen_access_column()
    ensure_timesheets_remaining_planned_column()

def ensure_timesheets_remaining_planned_column():
    """Add remaining_planned column to timesheets if missing."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(timesheets)")
    cols = [r[1] for r in cur.fetchall()]
    if "remaining_planned" not in cols:
        logging.info("Adding 'remaining_planned' column to timesheets table...")
        cur.execute("ALTER TABLE timesheets ADD COLUMN remaining_planned REAL NOT NULL DEFAULT 0.0")
        conn.commit()
    conn.close()


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


def ensure_timesheets_verification_column():
    """Add verification_filename column to timesheets if missing."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(timesheets)")
    cols = [r[1] for r in cur.fetchall()]
    if "verification_filename" not in cols:
        logging.info("Adding 'verification_filename' column to timesheets table...")
        cur.execute("ALTER TABLE timesheets ADD COLUMN verification_filename TEXT")
        conn.commit()
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


def ensure_default_admin():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM users")
    count = cur.fetchone()[0]
    if count == 0:
        cur.execute(
            "INSERT INTO users (username, email, password, role, status, screen_access) VALUES (?, ?, ?, ?, ?, ?)",
            ("admin", "admin@example.com", "admin", "Admin", "Active", "BOTH"),
        )
        conn.commit()
    conn.close()


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


def create_user_db(username, email, password, role, status, screen_access="BOTH"):
    username = (username or "").strip()
    email = (email or "").strip().lower()
    role = (role or "Employee").strip()
    status = (status or "Active").strip()

    if not username or not email:
        return False
    if username_exists(username) or email_exists(email):
        return False

    screen_access = normalize_screen_access(screen_access)
    if role != "Employee":
        screen_access = "BOTH"

    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO users (username, email, password, role, status, screen_access) VALUES (?, ?, ?, ?, ?, ?)",
            (username, email, password or "", role, status, screen_access),
        )
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def list_users_db(status_filter=None):
    """
    status_filter: None, 'Active', or 'Inactive'
    """
    conn = get_conn()
    cur = conn.cursor()

    if status_filter in ("Active", "Inactive"):
        cur.execute(
            "SELECT id, username, email, role, status, COALESCE(screen_access,'BOTH') "
            "FROM users WHERE status = ? ORDER BY id",
            (status_filter,),
        )
    else:
        cur.execute(
            "SELECT id, username, email, role, status, COALESCE(screen_access,'BOTH') "
            "FROM users ORDER BY id"
        )

    rows = cur.fetchall()
    conn.close()
    return rows


def get_user_db_by_id(user_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, username, email, role, status, COALESCE(screen_access,'BOTH') FROM users WHERE id = ?",
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
        "SELECT id, username, email, password, role, status, COALESCE(screen_access,'BOTH') "
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
    if role in ("Admin", "Manager"):
        return "BOTH"
    return normalize_screen_access(row[6])


def update_user_db(user_id, updates: dict):
    if not updates:
        return False

    existing = get_user_db_by_id(user_id)
    existing_role = existing[3] if existing else "Employee"
    target_role = (updates.get("role") or existing_role or "Employee").strip()

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
        if target_role != "Employee":
            sa = "BOTH"
        updates["screen_access"] = sa
    else:
        if "role" in updates and target_role != "Employee":
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
        dt = datetime.datetime.utcnow()
    ws = _sunday_of_date(dt.date())
    return ws.isoformat()

# -------------------------
# PPM OCR (auto-fill timesheet from Total row)
# -------------------------
MONTHS_MAP = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12
}

def _preprocess_for_ocr(img: Image.Image) -> Image.Image:
    gray = ImageOps.grayscale(img)
    gray = ImageEnhance.Contrast(gray).enhance(2.0)
    gray = gray.resize((gray.size[0] * 2, gray.size[1] * 2))
    bw = gray.point(lambda x: 0 if x < 170 else 255, '1')
    return bw

def _parse_week_start_from_ppm_text(text: str):
    if not text:
        return None
    t = ' '.join(text.split())
    m = re.search(
        r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})\s+to\s+"
        r"(?:(January|February|March|April|May|June|July|August|September|October|November|December)\s+)?(\d{1,2}),\s*(\d{4})",
        t,
        re.IGNORECASE,
    )
    if not m:
        return None
    month1 = (m.group(1) or '').lower()
    day1 = int(m.group(2))
    year = int(m.group(5))
    try:
        start = datetime.date(year, MONTHS_MAP[month1], day1)
        ws = _sunday_of_date(start)
        return ws.isoformat()
    except Exception:
        return None

def _to_hour_value(token: str, max_val=24):
    if token is None:
        return None
    s = token.strip().lower()
    if s in ('oh', 'o h'):
        return 0.0
    s = s.replace(' ', '')
    if s.endswith('h'):
        s = s[:-1]
    try:
        v = float(s)
        if v < 0:
            v = 0.0
        if v > max_val:
            return None
        v = round(v * 4) / 4.0
        return v
    except Exception:
        return None

def extract_ppm_hours_from_screenshot_total_row(image_path: str):
    if easyocr is None:
        return None
    global _EASYOCR_READER
    try:
        if _EASYOCR_READER is None:
            _EASYOCR_READER = easyocr.Reader(['en'])
        results = _EASYOCR_READER.readtext(image_path)
    except Exception:
        return None
    
    words = []
    full_text_parts = []
    for (bbox, text, prob) in results:
        txt = text.strip()
        if not txt:
            continue
        xs = [pt[0] for pt in bbox]
        ys = [pt[1] for pt in bbox]
        cx = sum(xs) / 4.0
        cy = sum(ys) / 4.0
        words.append({'t': txt, 'tl': txt.lower(), 'cx': cx, 'cy': cy})
        full_text_parts.append(txt)
        
    full_text = ' '.join(full_text_parts)
    week_start = _parse_week_start_from_ppm_text(full_text)
    day_keys = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat']
    col_x = {}
    for w in words:
        tl = w['tl']
        for dk in day_keys:
            if tl.startswith(dk):
                if dk not in col_x:
                    col_x[dk] = w['cx']
    if len(col_x) < 5:
        return None
    totals = [w for w in words if w['tl'] in ('total', 'totel', 'tota1')]
    if not totals:
        totals = [w for w in words if w['tl'].startswith('tot')]
    if not totals:
        return None
    total_y = totals[-1]['cy']
    y_band = 30
    row_words = [w for w in words if abs(w['cy'] - total_y) <= y_band]
    result = {'week_start': week_start}
    for dk in day_keys:
        cx = col_x.get(dk)
        if cx is None:
            result[dk] = 0.0
            continue
        candidates = []
        for w in row_words:
            v = _to_hour_value(w['t'])
            if v is None:
                continue
            dx = abs(w['cx'] - cx)
            if dx < 45:  # Bound the horizontal search gap
                candidates.append((dx, v))
        if not candidates:
            result[dk] = 0.0
        else:
            candidates.sort(key=lambda t: t[0])
            result[dk] = float(candidates[0][1])
            
    # Also extract Remaining Planned broadly from the image
    result['remaining_planned'] = 0.0
    rem_words = [w for w in words if "remaining" in w['tl']]
    if rem_words:
        rem_cx = sum(w['cx'] for w in rem_words) / len(rem_words)
        rem_cy = sum(w['cy'] for w in rem_words) / len(rem_words)
        rem_cands = []
        for w in words:
            if w['cy'] > rem_cy + 10 and abs(w['cx'] - rem_cx) < 80:
                v = _to_hour_value(w['t'], max_val=999)
                if v is not None:
                    rem_cands.append(v)
        if rem_cands:
            result['remaining_planned'] = float(rem_cands[-1])

    return result

def extract_ntt_hours_from_screenshot(image_path: str):
    import numpy as np
    from PIL import Image
    if easyocr is None:
        return None
    global _EASYOCR_READER
    try:
        if _EASYOCR_READER is None:
            _EASYOCR_READER = easyocr.Reader(['en'])
        
        img = Image.open(image_path)
        width, height = img.size
        crop_x = int(width * 0.28)
        img_cropped = img.crop((crop_x, 0, width, height))

        results = _EASYOCR_READER.readtext(np.array(img_cropped))
    except Exception:
        return None

    words = []
    for (bbox, text, prob) in results:
        txt = text.strip()
        if not txt:
            continue
        xs = [pt[0] for pt in bbox]
        ys = [pt[1] for pt in bbox]
        cx = sum(xs) / 4.0 + crop_x
        cy = sum(ys) / 4.0
        words.append({'t': txt, 'tl': txt.lower(), 'cx': cx, 'cy': cy})

    # Group into lines
    lines = []
    used = set()
    for i, w1 in enumerate(words):
        if i in used:
            continue
        line_words = [w1]
        used.add(i)
        for j, w2 in enumerate(words):
            if j not in used and abs(w1['cy'] - w2['cy']) < 15:
                line_words.append(w2)
                used.add(j)
        line_words.sort(key=lambda x: x['cx'])
        line_t = " ".join([w['t'] for w in line_words])
        line_tl = " ".join([w['tl'] for w in line_words])
        lines.append({
            'cy': sum(w['cy'] for w in line_words) / len(line_words),
            't': line_t,
            'tl': line_tl,
            'words': line_words
        })
    lines.sort(key=lambda x: x['cy'])
    try:
        import json
        with open("ocr_debug.json", "w") as f:
            json.dump(lines, f, indent=2)
    except Exception:
        pass

    entered_cxs = [w['cx'] for w in words if w['tl'] == 'entered']
    entered_cx = entered_cxs[0] if entered_cxs else None

    def guess_day(text):
        if 'sun' in text: return 'sun'
        if 'mon' in text: return 'mon'
        if 'tue' in text: return 'tue'
        if 'wed' in text or 'wco' in text: return 'wed'
        if 'thu' in text: return 'thu'
        if 'fri' in text: return 'fri'
        if 'sat' in text: return 'sat'
        return None

    date_pattern = re.compile(r'(\d{1,2})[,\.\s]+(202\d)')
    date_nodes = []
    
    for line in lines:
        m = date_pattern.search(line['tl'])
        if m:
            d = guess_day(line['tl'])
            date_nodes.append({
                'day': d,
                'cy': line['cy'],
                'date_tuple': (int(m.group(1)), int(m.group(2)))
            })

    result = {'sun': 0.0, 'mon': 0.0, 'tue': 0.0, 'wed': 0.0, 'thu': 0.0, 'fri': 0.0, 'sat': 0.0}
    if not date_nodes:
        return None

    # Connect missing days sequentially
    DAYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat']
    start_day = None
    for dn in date_nodes:
        if dn['day'] is not None:
            start_day = dn['day']
            break
    
    if start_day is None:
        start_day = 'sun'
        
    start_idx = DAYS.index(start_day)
    for i, dn in enumerate(date_nodes):
        dn['day'] = DAYS[(start_idx + i) % 7]

    for i, dnode in enumerate(date_nodes):
        min_cy = dnode['cy']
        max_cy = date_nodes[i+1]['cy'] if i + 1 < len(date_nodes) else float('inf')
        
        cands = []
        for line in lines:
            if min_cy < line['cy'] < max_cy:
                if entered_cx:
                    for w in line.get('words', []):
                        if abs(w['cx'] - entered_cx) < 90:
                            clean_tl = w['tl'].replace(' ', '')
                            m = re.search(r'^([0-9]{1,2})[^0-9]*([0-9]{2})', clean_tl)
                            if m:
                                val = float(m.group(1)) + float(m.group(2))/100.0
                                if val <= 24:
                                    cands.append(round(val * 4) / 4.0)
                            else:
                                # fallback for exact integer matching without .00, rare but possible
                                m2 = re.search(r'^([0-9]{1,2})$', clean_tl)
                                if m2 and float(m2.group(1)) <= 24:
                                    cands.append(float(m2.group(1)))
        
        if cands:
            result[dnode['day']] = max(cands)
    first_date_str = ""
    for line in lines:
        match = re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)', line['tl'])
        if match:
            first_date_str = match.group(1)
            break
            
    months_keys = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    month_int = 1
    if first_date_str in months_keys:
        month_int = months_keys.index(first_date_str) + 1
        
    try:
        import datetime
        dt = datetime.date(year=int(date_nodes[0]['date_tuple'][1]), month=month_int, day=int(date_nodes[0]['date_tuple'][0]))
        days_to_sub = (dt.weekday() + 1) % 7
        ws = dt - datetime.timedelta(days=days_to_sub)
        result['week_start'] = ws.strftime('%d-%m-%Y')
    except Exception:
        pass
        
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
    updated_at = datetime.datetime.utcnow().isoformat(sep=" ", timespec="seconds")
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
    uploaded_at = datetime.datetime.utcnow().isoformat(sep=" ", timespec="seconds")
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
               COALESCE(submitted,0), COALESCE(submitted_at,''), COALESCE(remaining_planned, 0.0)
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
        "remaining_planned": float(row[12] or 0.0)
    }


def get_timesheet_values(upload_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT week_start, sun, mon, tue, wed, thu, fri, sat, total, COALESCE(submitted,0), updated_at
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
    submitted_at = datetime.datetime.utcnow().isoformat(sep=" ", timespec="seconds")
    cur.execute("""
        UPDATE timesheets
           SET submitted = 1,
               submitted_at = ?
         WHERE upload_id = ?
    """, (submitted_at, upload_id))
    conn.commit()
    conn.close()


def get_verification_filename(upload_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT verification_filename FROM timesheets WHERE upload_id = ?", (upload_id,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row and row[0] else None


def save_verification_filename(upload_id: int, filename: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE timesheets SET verification_filename = ? WHERE upload_id = ?", (filename, upload_id))
    conn.commit()
    conn.close()


def upsert_timesheet(upload_id: int, week_start: str, mon, tue, wed, thu, fri, sat, sun, remaining_planned=0.0):
    if is_timesheet_submitted(upload_id):
        raise ValueError("Timesheet is submitted and locked. Contact Admin if changes are required.")

    total = float(mon) + float(tue) + float(wed) + float(thu) + float(fri) + float(sat) + float(sun)
    updated_at = datetime.datetime.utcnow().isoformat(sep=" ", timespec="seconds")

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
                   updated_at = ?,
                   remaining_planned = ?
             WHERE upload_id = ?
        """, (week_start, mon, tue, wed, thu, fri, sat, sun, total, updated_at, remaining_planned, upload_id))
    else:
        cur.execute("""
            INSERT INTO timesheets (
                upload_id, week_start, mon, tue, wed, thu, fri, sat, sun, total,
                submitted, submitted_at, updated_at, remaining_planned
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL, ?, ?)
        """, (upload_id, week_start, mon, tue, wed, thu, fri, sat, sun, total, updated_at, remaining_planned))

    conn.commit()
    conn.close()
    return total


def delete_upload(upload_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT stored_filename FROM uploads WHERE id = ?", (upload_id,))
    row = cur.fetchone()
    stored_fn = row[0] if row else None

    cur.execute("DELETE FROM timesheets WHERE upload_id = ?", (upload_id,))
    cur.execute("DELETE FROM uploads WHERE id = ?", (upload_id,))
    conn.commit()
    conn.close()

    if stored_fn:
        path = os.path.join(UPLOAD_DIR, stored_fn)
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            logging.exception("Failed to remove uploaded file from disk: %s", path)


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

    modules_in_this_week = {get_upload_module(u) for u in group_ids}
    if "PPM" not in modules_in_this_week:
        raise ValueError(
            f"Cannot submit: No matching PPM report found for the week starting {week_start}. "
            "Please upload a PPM screenshot for the exact same week before submitting NTT."
        )
    if "NTT" not in modules_in_this_week:
        raise ValueError(
            f"Cannot submit: No matching NTT report found for the week starting {week_start}."
        )

    candidate = (sun, mon, tue, wed, thu, fri, sat)

    for uid in group_ids:
        row = get_timesheet_values(uid)
        if not row:
            continue
        _, sun2, mon2, tue2, wed2, thu2, fri2, sat2, total2, submitted2, _updated2 = row
        existing = (sun2, mon2, tue2, wed2, thu2, fri2, sat2)

        if int(submitted2) == 1 and not timesheet_hours_equal(existing, candidate):
            raise ValueError(
                "Cannot submit: PPM and NTT hours must match for the same week. "
                "A submitted (locked) timesheet already exists for this week with different hours."
            )

        if float(total2 or 0) > 0 and not timesheet_hours_equal(existing, candidate):
            raise ValueError(
                "Cannot submit: Entered hours in PPM and NTT must match for the same week. "
                "Please make the hours identical in both modules before submitting."
            )

    for uid in group_ids:
        if is_timesheet_submitted(uid):
            continue
        upsert_timesheet(uid, week_start, mon, tue, wed, thu, fri, sat, sun)
        mark_timesheet_submitted(uid)


# -------------------------
# Email reminders
# -------------------------
def add_email_reminder(recipient, subject, message, scheduled_date, created_by=None):
    conn = get_conn()
    cur = conn.cursor()
    created_at = datetime.datetime.utcnow().isoformat(sep=" ", timespec="seconds")
    cur.execute("""
        INSERT INTO email_reminders (recipient, subject, message, scheduled_date, created_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (recipient.strip(), subject.strip(), message.strip(), scheduled_date.strip(), created_by, created_at))
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid


def list_email_reminders():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, recipient, subject, message, scheduled_date, created_by, created_at
        FROM email_reminders
        ORDER BY id DESC
    """)
    rows = cur.fetchall()
    conn.close()
    return rows



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

def _is_late_submission(module: str, week_start_iso: str, submitted_at: str) -> bool:
    if not submitted_at or not week_start_iso:
        return False
    try:
        ws = datetime.date.fromisoformat(week_start_iso)
    except Exception:
        return False
    try:
        sub_dt = datetime.datetime.fromisoformat(str(submitted_at))
    except Exception:
        return False
    mod = (module or '').strip().upper()
    if mod == 'PPM':
        cutoff_date = ws + datetime.timedelta(days=3)
    elif mod == 'NTT':
        cutoff_date = ws + datetime.timedelta(days=5)
    else:
        return False
    cutoff_dt = datetime.datetime.combine(cutoff_date, datetime.time(15, 0))
    return sub_dt > cutoff_dt

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
            MAX(COALESCE(t.submitted_at,'')) AS submitted_at
        FROM uploads u
        JOIN timesheets t ON t.upload_id = u.id
        LEFT JOIN users usr ON LOWER(COALESCE(usr.email,'')) = LOWER(COALESCE(u.uploaded_by,''))
        WHERE COALESCE(u.module,'') IN ('PPM','NTT')
        GROUP BY LOWER(COALESCE(u.uploaded_by,'')), COALESCE(usr.username, LOWER(COALESCE(u.uploaded_by,''))),
                 COALESCE(u.module,''), COALESCE(t.week_start,'')
    """)
    rows = cur.fetchall()
    conn.close()

    import calendar
    import string
    
    y, m = map(int, month_str.split('-'))
    last_day_num = calendar.monthrange(y, m)[1]
    
    buckets = []
    for start_day in range(1, last_day_num + 1, 7):
        end_day = min(start_day + 6, last_day_num)
        start_date = datetime.date(y, m, start_day)
        end_date = datetime.date(y, m, end_day)
        buckets.append((start_date, end_date))

    def week_bucket(ws: str):
        try:
            ws_date = datetime.date.fromisoformat(ws)
            if ws_date.month == m:
                target_day = ws_date.day
            elif (ws_date + datetime.timedelta(days=6)).month == m:
                target_day = (ws_date + datetime.timedelta(days=6)).day
            else:
                return None
            for i, b in enumerate(buckets):
                if b[0].day <= target_day <= b[1].day:
                    return i + 1
        except Exception:
            pass
        return None

    data = {}
    for email, rname, module, ws, total, sub_at in rows:
        if not ws or not _week_overlaps_month(ws, first_day, last_day):
            continue
        wk = week_bucket(ws)
        if wk is None:
            # Fallback for boundary weeks
            wk = 1
        mod = (module or '').strip().upper()
        if mod not in ('PPM','NTT'):
            continue
        rkey = (rname or email or '').strip()
        if not rkey:
            continue
        data.setdefault(rkey, {})
        data[rkey].setdefault(wk, {'PPM': 0.0, 'NTT': 0.0})
        data[rkey][wk][mod] += float(total or 0.0)

    wb = Workbook()
    sh = wb.active
    sh.title = month_str

    sh.cell(row=1, column=1, value='Resource Name')
    sh.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    
    num_weeks = len(buckets)
    col = 2
    week_cols = {}
    for wk in range(1, num_weeks + 1):
        week_cols[wk] = col
        col += 2
    total_col = col

    def format_week_header(b) -> str:
        return f"{b[0].strftime('%b %d')} - {b[1].strftime('%b %d')}"

    for wk in range(1, num_weeks + 1):
        c0 = week_cols[wk]
        header_title = format_week_header(buckets[wk-1])
        sh.cell(row=1, column=c0, value=header_title)
        sh.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0+1)
        sh.cell(row=2, column=c0, value='PPM')
        sh.cell(row=2, column=c0+1, value='NTT')
        
    sh.cell(row=1, column=total_col, value='Total')
    sh.merge_cells(start_row=1, start_column=total_col, end_row=2, end_column=total_col)

    thin = Side(style='thin', color='000000') if Side is not None else None
    border = Border(left=thin, right=thin, top=thin, bottom=thin) if Border is not None and thin is not None else None

    for r in (1, 2):
        for c in range(1, total_col + 1):
            cell = sh.cell(row=r, column=c)
            if Font is not None:
                cell.font = Font(bold=True)
            if Alignment is not None:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if border is not None:
                cell.border = border

    sh.row_dimensions[1].height = 22
    sh.row_dimensions[2].height = 20

    r_out = 3
    for rname in sorted(data.keys(), key=lambda x: x.lower()):
        sh.cell(row=r_out, column=1, value=rname)
        
        total_ntt = 0.0
        for wk in range(1, num_weeks + 1):
            c0 = week_cols[wk]
            ppm = data.get(rname, {}).get(wk, {}).get('PPM', 0.0)
            ntt = data.get(rname, {}).get(wk, {}).get('NTT', 0.0)
            total_ntt += ntt
            
            sh.cell(row=r_out, column=c0, value=round(ppm, 2))
            sh.cell(row=r_out, column=c0+1, value=round(ntt, 2))

        sh.cell(row=r_out, column=total_col, value=round(total_ntt, 2))

        for c in range(1, total_col + 1):
            cell = sh.cell(row=r_out, column=c)
            if Alignment is not None:
                cell.alignment = Alignment(horizontal='left', vertical='center') if c == 1 else Alignment(horizontal='center', vertical='center')
            if border is not None:
                cell.border = border
        r_out += 1

    sh.freeze_panes = 'B3'
    sh.column_dimensions['A'].width = 22
    letters = list(string.ascii_uppercase)
    for i in range(26):
        letters.append("A" + string.ascii_uppercase[i])
        
    for c in range(2, total_col):
        sh.column_dimensions[letters[c-1]].width = 16
    sh.column_dimensions[letters[total_col-1]].width = 12

    from io import BytesIO
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()
# Initialize DB and default admin
init_db()
ensure_default_admin()


# -------------------------
# UI templates
# -------------------------
PAGE_STYLE = """
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
:root{--bg:#f4f6f8;--nav:#0f1724;--accent:#2563eb;--card:#ffffff;--muted:#6b7280}
*{box-sizing:border-box}
body{margin:0;font-family:Inter,"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:#111827}
header.app-header{position:fixed;top:0;left:0;right:0;height:64px;background:linear-gradient(90deg,var(--nav),#0b1220);color:#fff;display:flex;align-items:center;justify-content:space-between;padding:0 20px;z-index:1000;box-shadow:0 2px 8px rgba(2,6,23,0.15)}
.brand{display:flex;align-items:center;gap:12px;font-weight:700}
.brand .logo{width:36px;height:36px;border-radius:8px;background:linear-gradient(135deg,var(--accent),#7c3aed);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;box-shadow:0 2px 6px rgba(37,99,235,0.2)}
.header-right{display:flex;align-items:center;gap:12px;font-weight:600}
.layout{display:flex;margin-top:64px;min-height:calc(100vh - 64px)}
aside.sidebar{width:220px;background:#0b1220;color:#e6eef8;padding:18px 12px;border-right:1px solid rgba(255,255,255,0.03);position:sticky;top:64px;height:calc(100vh - 64px)}
.nav-item{display:flex;align-items:center;gap:12px;padding:10px 12px;border-radius:8px;color:inherit;text-decoration:none;margin-bottom:6px;font-weight:600;position:relative}
.nav-item i{width:20px;text-align:center;color:#9fb3d8}
.nav-item:hover{background:rgba(255,255,255,0.04);cursor:pointer}
.dropdown-content{display:none;position:absolute;top:100%;left:0;background:#fff;min-width:220px;box-shadow:0 8px 20px rgba(2,6,23,0.12);border-radius:8px;overflow:hidden;z-index:999}
.nav-item:hover .dropdown-content{display:block}
.dropdown-content a{display:block;padding:10px 12px;color:#111827;text-decoration:none}
.dropdown-content a:hover{background:#f4f6f8}
main.content{flex:1;padding:28px}
.card{background:var(--card);border-radius:12px;padding:20px;box-shadow:0 6px 18px rgba(15,23,42,0.06)}
.card h2{margin:0 0 12px;font-size:18px;color:#0f1724}
.muted{color:var(--muted);font-size:14px}
label{display:block;font-weight:600;margin-top:12px;color:#111827}
input,select,textarea{width:100%;padding:10px 12px;margin-top:8px;border:1px solid #e6e9ee;border-radius:8px;background:#fff;font-size:14px}
textarea{min-height:110px;resize:vertical}
.btn{display:inline-block;background:var(--accent);color:#fff;padding:10px 16px;border-radius:8px;border:none;font-weight:700;margin-top:16px;cursor:pointer;text-decoration:none}
.btn.secondary{background:#111827;color:#fff}
.btn.danger{background:#ef4444}
.btn:disabled{opacity:.6;cursor:not-allowed}
.small{padding:6px 10px;font-size:13px}
.preview{margin-top:12px;border:1px solid #e6e9ee;padding:10px;border-radius:8px;background:#fafafa}
.controls{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}

.upload-item{display:flex;gap:12px;align-items:flex-start;padding:10px 0;border-bottom:1px dashed #f1f5f9}
.upload-item img{max-width:140px;border-radius:8px;border:1px solid #e6e9ee}
.upload-meta{font-size:13px;color:#374151;flex:1}
.badge{display:inline-block;font-size:11px;padding:2px 8px;border-radius:999px;margin-left:8px;background:#eef2ff;color:#4338ca;border:1px solid #c7d2fe}
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

/* ===== Filter Bar (Admin PPM/NTT) ===== */
.filter-card{
  margin-top:14px;
  padding:14px 14px 12px;
  border:1px solid #e6e9ee;
  border-radius:14px;
  background:linear-gradient(180deg,#ffffff,#fbfdff);
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
  font-size:12px;
  text-transform:uppercase;
  letter-spacing:.04em;
  color:#64748b;
}
.filter-field input,.filter-field select{
  margin-top:6px;
  height:40px;
  padding:8px 12px;
  border-radius:10px;
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
.controls .btn,
.filter-form .btn{margin-top:0}

/* Outline button variant */
.btn.outline{
  background:#fff;
  color:var(--accent);
  border:1px solid #bfdbfe;
}
.btn.outline:hover{background:#eff6ff}

/* Optional: make small buttons more pill-like */
.btn.small{border-radius:10px}

@media(max-width:880px){aside.sidebar{display:none}main.content{padding:18px}}
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
            '<div class="header-right">Welcome</div>'
            '</header>'
        )
    return (
        '<header class="app-header">'
        f'{brand}'
        f'<div class="header-right">Signed in as <strong>{html.escape(display_name)}</strong>'
        f'<a class="btn secondary" style="padding:8px 12px;font-size:13px;margin-left:12px" href="/logout">Logout</a>'
        '</div></header>'
    )


def build_sidebar_for_role(role, screen_access="BOTH"):
    if role in ("Admin", "Manager"):
        screen_access = "BOTH"
    screen_access = normalize_screen_access(screen_access)

    module_links = []
    if role in ("Admin", "Manager") or screen_access == "BOTH":
        module_links = [
            '<a href="/upload/PPM" title="PPM screen prints">PPM</a>',
            '<a href="/upload/NTT" title="NTT screen prints">NTT</a>',
        ]
    elif screen_access == "PPM":
        module_links = ['<a href="/upload/PPM" title="PPM screen prints">PPM</a>']
    elif screen_access == "NTT":
        module_links = ['<a href="/upload/NTT" title="NTT screen prints">NTT</a>']
    elif screen_access == "EMAIL_NTT":
        module_links = ['<a href="/upload/NTT" title="NTT screen prints">NTT</a>']

    if role == "Admin" or screen_access == "EMAIL_NTT":
        module_links.append('<a href="/upload/Email" title="Email confirmation from manager">Email</a>')

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
            '<a href="/user-management/list">User List</a>'
            '</div></div>'
        )
        parts.append(screen_module_dropdown)
        parts.append('<a class="nav-item" href="/email-settings"><i class="fas fa-envelope"></i> Email Settings</a>')
        parts.append('<a class="nav-item" href="/export"><i class="fas fa-file-export"></i> Export Data</a>')
    elif role == "Manager":
        parts.append(screen_module_dropdown)
        parts.append('<a class="nav-item" href="/email-settings"><i class="fas fa-envelope"></i> Email Settings</a>')
        parts.append('<a class="nav-item" href="/export"><i class="fas fa-file-export"></i> Export Data</a>')
    elif role == "Employee":
        parts.append(screen_module_dropdown)

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
    sidebar_html = build_sidebar_for_role(role or "", screen_access or "BOTH")
    return (
        f"<html><head>{PAGE_STYLE}</head>"
        f"<body>{NAV(display_name)}"
        f"<div class='layout'>{sidebar_html}<main class='content'>{content}</main></div>"
        f"</body></html>"
    )


# -------------------------
# Multipart parser
# -------------------------
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
        ensure_timesheet_for_upload(upload_id, datetime.datetime.utcnow().isoformat(sep=" ", timespec="seconds"))
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
    disabled_attr = "disabled" if is_submitted else ""

    if module_for_upload in ("PPM", "NTT"):
        disabled_attr = "readonly style='pointer-events:none; background-color:#f4f5f7; color:#7d8597; border-color:#e6e9ee;'" if not is_submitted else "disabled"

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

    if not is_submitted:
        if module_for_upload == "NTT":
            controls_html = """
            <div class="controls">
              <button class="btn secondary small ts-submit-btn" type="submit" formaction="/timesheet/submit">
                <i class="fa fa-check"></i> Submit
              </button>
            </div>
            <div class="muted" style="margin-top:6px">
              <b>Submit-only rule:</b> At Submit, PPM and NTT hours must match for the same week (Sun–Sat).
            </div>
            """
        elif module_for_upload == "PPM":
            controls_html = """
            <div class="muted" style="margin-top:6px; color:#ef4444; font-weight: 500;">
              <b>Action required:</b> Attach verification screenprint below and submit to lock this timesheet.
            </div>
            """
        else:
            controls_html = """
            <div class="controls">
              <button class="btn small" type="submit">
                <i class="fa fa-save"></i> Save Hours
              </button>
    
              <button class="btn secondary small ts-submit-btn" type="submit" formaction="/timesheet/submit">
                <i class="fa fa-check"></i> Submit (Lock)
              </button>
            </div>
            <div class="muted" style="margin-top:6px">
              <b>Submit-only rule:</b> At Submit, PPM and NTT hours must match for the same week (Sun–Sat). Submit locks the week.
            </div>
            """
    else:
        controls_html = """
        <div class="muted" style="margin-top:8px">
          This timesheet is locked. Contact Admin if changes are required.
        </div>
        """

    ts_head_right_extra = admin_delete_html or ""

    v_file = get_verification_filename(upload_id)
    v_html = ""
    if module_for_upload == "PPM":
        if v_file:
            v_file_url = f"/{UPLOAD_DIR}/{v_file}"
            v_html = f"""
            <div style="margin-top:16px; padding:10px 12px; border-top:1px solid #e6e9ee; background:#fafafa;">
              <h4 style="margin:0 0 8px 0;font-size:13px;color:#10b981;"><i class="fa fa-check-circle"></i> Verification Screenprint Attached</h4>
              <a class='thumb-link' href='{html.escape(v_file_url, quote=True)}' data-url='{html.escape(v_file_url, quote=True)}' data-filename='{html.escape(v_file, quote=True)}' style='display:inline-block; max-width:140px; margin-top:4px;'>
                 <img src='{html.escape(v_file_url, quote=True)}' alt='verification screenshot' style='max-width:100%; border-radius:8px; border:1px solid #e6e9ee;'>
              </a>
            </div>
            """
        else:
            v_html = f"""
            <div class="verify-wrap" data-uid="{upload_id}" style="margin-top:16px; padding:10px 12px; border-top:1px solid #e6e9ee; background:#fafafa;">
              <h4 style="margin:0 0 8px 0;font-size:13px;">Attach Verification Screenprint</h4>
              <div style="display:flex;gap:8px;align-items:center;">
                 <input type="hidden" name="return_to" value="{html.escape(return_to, quote=True)}">
                 <button type="button" class="btn outline small btn-cap-verify" id="btnCapVerify_{upload_id}">Capture Screen</button>
                 <button type="button" class="btn secondary small btn-submit-verify" id="btnSubmitVerify_{upload_id}" style="display:none;">Submit</button>
                 <button type="button" class="btn danger small btn-close-verify" id="btnCloseVerify_{upload_id}" style="display:none;">Cancel</button>
              </div>
              <video id="vid_verify_{upload_id}" style="display:none; max-width:100%; margin-top:10px;" autoplay muted playsinline></video>
              <div id="shot_verify_{upload_id}" style="display:none; margin-top:10px; max-width:100%;"></div>
            </div>
            """

    rem_plan = float(ts.get("remaining_planned", 0.0))
    topup_warning = ""
    # "pop-up before the PPM timesheet tablet in red color just as notification thats it."
    if module_for_upload == "PPM" and not is_submitted and rem_plan > 0 and rem_plan < 45:
        topup_warning = f"""
        <div style="background-color:#fee2e2; color:#b91c1c; padding:12px 16px; margin: 12px 12px 0 12px; border:1px solid #fca5a5; border-radius:6px; font-weight:600; font-size:13px; display:flex; align-items:center; gap:8px;">
            <i class="fa fa-triangle-exclamation"></i>
            <span>Warning: Your Remaining Planned period is {rem_plan:g} hrs. Please top up your hours for the next week.</span>
        </div>
        """

    return f"""
<details class="ts-details">
  <summary>
    Timesheet (Week-wise) — <span id="ts_range_{upload_id}">{html.escape(week_range_text)}</span>
  </summary>

  <div class="ts-wrap ts-form" data-upload-id="{upload_id}">
    {topup_warning}
    <div class="ts-head">
      <div class="muted" style="font-weight:700;">
        Week Start (Sun):
        <input class="ts-week" type="date" value="{week_start_iso}"
               style="width:auto;display:inline-block;margin-left:8px;" {disabled_attr}>
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
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="sun" value="{sun}" {disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="mon" value="{mon}" {disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="tue" value="{tue}" {disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="wed" value="{wed}" {disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="thu" value="{thu}" {disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="fri" value="{fri}" {disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="sat" value="{sat}" {disabled_attr}></td>
            <td>
              <input type="text" readonly value="{total}" class="ts-total-input" id="ts_total_in_{upload_id}">
            </td>
          </tr>
        </tbody>
      </table>

      {controls_html}
    </form>
{v_html}
  </div>
</details>
"""


def build_timesheet_js():
    return """
<script>
(function(){
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

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
      submitBtn.disabled = false;
      submitBtn.title = 'Submit and lock this week (PPM and NTT must match).';
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

  let vStream=null, vTrack=null, vBlob=null, vUrl=null, vUid=null, isVUploading=false;

  function stopVStream(){
    try { if(vTrack) vTrack.stop(); if(vStream) vStream.getTracks().forEach(t=>t.stop()); } catch(e){}
    vStream=null; vTrack=null;
    if(vUid) {
      const vid = document.getElementById('vid_verify_'+vUid);
      if(vid) { vid.pause(); vid.srcObject=null; vid.style.display='none'; }
    }
  }

  function clearVShot(){
    if(vUrl){ URL.revokeObjectURL(vUrl); vUrl=null; }
    vBlob=null;
    if(vUid){
      const shot = document.getElementById('shot_verify_'+vUid);
      if(shot){ shot.innerHTML=''; shot.style.display='none'; }
      const btnCap = document.getElementById('btnCapVerify_'+vUid);
      const btnSub = document.getElementById('btnSubmitVerify_'+vUid);
      const btnClo = document.getElementById('btnCloseVerify_'+vUid);
      if(btnCap) { btnCap.style.display = 'inline-block'; btnCap.disabled = false; }
      if(btnSub) btnSub.style.display = 'none';
      if(btnClo) btnClo.style.display = 'none';
      vUid = null;
    }
  }

  function waitForVVideoReady(video, timeoutMs){
    return new Promise((resolve) => {
      if(!video) return resolve();
      if (video.readyState >= 2 && video.videoWidth > 0 && video.videoHeight > 0) return resolve();
      let done=false;
      const finish = () => { if(done) return; done=true; video.onloadedmetadata=null; resolve(); };
      const t = setTimeout(finish, timeoutMs || 2500);
      video.onloadedmetadata = finish;
    });
  }

  document.addEventListener('click', async function(e){
    if(e.target.classList && e.target.classList.contains('btn-cap-verify')){
      e.preventDefault();
      const wrap = e.target.closest('.verify-wrap');
      if(!wrap) return;
      const uid = wrap.getAttribute('data-uid');
      if(vUid && vUid !== uid) { stopVStream(); clearVShot(); }
      vUid = uid;

      const btnCap = document.getElementById('btnCapVerify_'+uid);
      const btnSub = document.getElementById('btnSubmitVerify_'+uid);
      const btnClo = document.getElementById('btnCloseVerify_'+uid);
      const vid = document.getElementById('vid_verify_'+uid);
      const shot = document.getElementById('shot_verify_'+uid);

      if(!btnCap || !vid || !shot) return;

      try {
        btnCap.disabled = true;
        vStream = await navigator.mediaDevices.getDisplayMedia({ video: { cursor: 'always' }, audio: false });
        const tracks = vStream.getVideoTracks();
        vTrack = (tracks && tracks[0]) ? tracks[0] : null;
        if(vTrack) vTrack.onended = () => { stopVStream(); clearVShot(); };

        vid.srcObject = vStream;
        vid.style.display = 'block';

        try { await vid.play(); } catch(err){}
        await waitForVVideoReady(vid, 2500);
        await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));

        const w = vid.videoWidth || 1280;
        const h = vid.videoHeight || 720;
        const c = document.createElement('canvas');
        c.width = w; c.height = h;
        const ctx = c.getContext('2d');
        ctx.drawImage(vid, 0, 0, w, h);

        c.toBlob((b) => {
          if(!b){ alert('Could not format image.'); stopVStream(); btnCap.disabled=false; return; }
          vBlob = b;
          if(vUrl) URL.revokeObjectURL(vUrl);
          vUrl = URL.createObjectURL(b);
          shot.innerHTML = '<img style="max-width:100%;border-radius:10px;box-shadow:0 4px 18px #0001" src="'+vUrl+'">';
          shot.style.display='block';
          
          btnCap.style.display = 'none';
          btnSub.style.display = 'inline-block';
          btnSub.disabled = false;
          btnClo.style.display = 'inline-block';

          stopVStream();
        }, 'image/png');
      } catch(err) {
        console.error(err);
        alert('Capture was cancelled or denied.');
        stopVStream();
        btnCap.disabled = false;
      }
    }

    if(e.target.classList && e.target.classList.contains('btn-close-verify')){
      e.preventDefault();
      stopVStream();
      clearVShot();
    }

    if(e.target.classList && e.target.classList.contains('btn-submit-verify')){
      e.preventDefault();
      if(isVUploading) return;
      if(!vBlob || !vUid) { alert('No image captured.'); return; }
      
      const btnSub = e.target;
      isVUploading = true;
      btnSub.disabled = true;

      const fd = new FormData();
      fd.append('upload_id', vUid);
      
      const form = btnSub.closest('.verify-wrap');
      const rtNode = form ? form.querySelector('input[name="return_to"]') : null;
      if(rtNode) fd.append('return_to', rtNode.value);
      
      fd.append('verification_file', vBlob, 'verification.png');

      try {
        const res = await fetch('/timesheet/verify', { method: 'POST', body: fd });
        if(res.ok){
          location.reload();
        } else {
          alert('Upload failed: ' + res.status);
          btnSub.disabled = false;
          isVUploading = false;
        }
      } catch(err) {
        console.error(err);
        alert('Upload error.');
        btnSub.disabled = false;
        isVUploading = false;
      }
    }
  });
})();
</script>
"""


# -------------------------
# Capture UI builder — PPM/NTT only (NO download for unsaved)
# -------------------------
def build_capture_ui(module_ctx: str):
    title = f"Screen Print Capture - {module_ctx}"
    upload_url = f"/upload/{module_ctx}"
    hint = f"This page captures and lists only {module_ctx} screen prints."
    capture_label = f"Capture {module_ctx} Screen Print"

    parts = []
    parts.append(f"<h2>{html.escape(title)}</h2>")
    parts.append(
        "<p class='muted' style='margin-bottom:8px'>"
        "Click <b>Capture</b> to select a window/tab/screen. A snapshot will be taken and shown in the preview area. "
        "Click <b>Save</b> to upload it. "
        + html.escape(hint) +
        "</p>"
    )

    parts.append(
        "<div style='display:flex;gap:12px;flex-wrap:wrap;align-items:center'>"
        f"<button type='button' class='btn' id='btnCapture'>{html.escape(capture_label)}</button>"
        "</div>"
    )

    parts.append("<video id='capPreview' style='display:none' autoplay muted playsinline></video>")

    parts.append(
        "<div id='shotContainer' class='preview' style='display:none;margin-top:12px'>"
        "  <div id='shotBox'></div>"
        "  <div class='controls'>"
        "    <button type='button' class='btn' id='btnSave'><i class='fa fa-cloud-upload-alt'></i> Save</button>"
        "    <button type='button' class='btn danger' id='btnCloseShot'><i class='fa fa-times'></i> Close</button>"
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

  const AUTO_UPLOAD_PPM = (moduleName === "PPM" || moduleName === "NTT");

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

  // PPM / NTT: auto-upload is enabled, but Save button remains available (optional manual upload)
  if(AUTO_UPLOAD_PPM){{
    try{{ btnSave.innerHTML = '<i class=\"fa fa-cloud-upload-alt\"></i> Save'; }}catch(e){{}}
  }}
  btnCloseShot.onclick = clearShot;
}})();
</script>
"""
    parts.append(js)
    return "".join(parts)


def normalize_return_to(rt: str, role: str, email: str) -> str:
    rt = rt or "/upload/PPM"
    if rt not in ("/upload/PPM", "/upload/NTT", "/upload/Email"):
        rt = "/upload/PPM"

    if role != "Employee":
        return rt

    access = get_screen_access_by_email(email)
    if access == "PPM":
        return "/upload/PPM"
    if access == "NTT":
        return "/upload/NTT"
    if access == "EMAIL_NTT":
        return "/upload/NTT" if rt == "/upload/PPM" else rt
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
                    content = (
                        "<div class='card' style='max-width:420px;margin:40px auto;text-align:left'>"
                        "<h2 style='text-align:center'>Sign in</h2>"
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
                    "<h2>Welcome</h2>"
                    "<p class='muted'>Use the left navigation to capture PPM/NTT screen prints. "
                    "Save hours anytime; Submit (Lock) enforces PPM/NTT match for the week and locks the week.</p>"
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

            protected_prefixes = ("/user-management", "/upload", "/email-settings", "/export", "/timesheet", f"/{UPLOAD_DIR}/")
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
                allowed_exact = {"/", "/logout"}
                allowed_prefixes = [f"/{UPLOAD_DIR}/", "/timesheet"]
                if screen_access in ("PPM", "BOTH"):
                    allowed_prefixes.append("/upload/PPM")
                if screen_access in ("NTT", "BOTH"):
                    allowed_prefixes.append("/upload/NTT")

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
                    "<label>Employee Screen Access (only for Employee role)</label>"
                    "<select name='screen_access' id='saSel'>"
                    "<option value='BOTH' selected>PPM &amp; NTT (Both)</option>"
                    "<option value='PPM'>PPM Only</option>"
                    "<option value='EMAIL_NTT'>Email &amp; NTT</option>"
                    "</select>"
                    "<div class='muted' style='margin-top:8px'>Managers/Admin automatically get Both modules (Option A).</div>"
                    "<button class='btn' type='submit'>Create</button>"
                    "</form>"
                    "<script>"
                    "(function(){"
                    "const role=document.getElementById('roleSel');"
                    "const sa=document.getElementById('saSel');"
                    "function sync(){"
                    "  const isEmp=(role.value==='Employee');"
                    "  sa.disabled=!isEmp;"
                    "  if(!isEmp) sa.value='BOTH';"
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
                <form method="get" action="/user-management/list" class="controls" style="margin-top:0">
                  <div style="min-width:220px">
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
                for idx, (uid, uname, email_addr, rrole, status, sa) in enumerate(rows, start=1):
                    rows_html += (
                        "<tr>"
                        f"<td style='padding:8px'>{idx}</td>"
                        f"<td style='padding:8px'>{html.escape(uname)}</td>"
                        f"<td style='padding:8px'>{html.escape(email_addr or '')}</td>"
                        f"<td style='padding:8px'>{html.escape(rrole)}</td>"
                        f"<td style='padding:8px'>{html.escape(status)}</td>"
                        f"<td style='padding:8px'>{html.escape(sa)}</td>"
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
                    "<table style='width:100%;border-collapse:collapse;margin-top:12px'>"
                    "<thead><tr style='text-align:left'>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Number</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Username</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Email</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Role</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Status</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Screen Access</th>"
                    "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Actions</th>"
                    "</tr></thead><tbody>"
                    + (rows_html if rows_html else "<tr><td colspan='7' style='padding:8px'>No users found.</td></tr>")
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

                uid_val, uname, email_addr, rrole, status, sa = row
                sa = normalize_screen_access(sa)

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
                    for opt, label in (("BOTH","PPM &amp; NTT (Both)"), ("PPM","PPM Only"), ("EMAIL_NTT","Email &amp; NTT"))
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
                    "<label>Employee Screen Access (only for Employee role)</label>"
                    f"<select name='screen_access' id='saSel'>{access_opts}</select>"
                    "<label>Password (leave blank to keep unchanged)</label>"
                    "<input type='password' name='password' placeholder='********'>"
                    "<button class='btn' type='submit'>Save Changes</button>"
                    "</form>"
                    "<script>"
                    "(function(){"
                    "const role=document.getElementById('roleSel');"
                    "const sa=document.getElementById('saSel');"
                    "function sync(){"
                    "  const isEmp=(role.value==='Employee');"
                    "  sa.disabled=!isEmp;"
                    "  if(!isEmp) sa.value='BOTH';"
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

            # Capture pages: /upload/PPM, /upload/NTT, /upload/Email
            if path in ("/upload/PPM", "/upload/NTT", "/upload/Email"):
                module_ctx = "Email" if path == "/upload/Email" else ("PPM" if path == "/upload/PPM" else "NTT")

                if role == "Manager" and module_ctx == "Email":
                    return self._msg("Access denied: you do not have permission for Email module.", display_name)

                # employee module permission check
                if role == "Employee":
                    if module_ctx == "PPM" and screen_access not in ("PPM", "BOTH"):
                        return self._msg("Access denied: you do not have permission for PPM module.", display_name)
                    if module_ctx == "NTT" and screen_access not in ("NTT", "BOTH", "EMAIL_NTT"):
                        return self._msg("Access denied: you do not have permission for NTT module.", display_name)
                    if module_ctx == "Email" and screen_access != "EMAIL_NTT":
                        return self._msg("Access denied: you do not have permission for Email module.", display_name)

                # Admin filters
                qs = parse_qs(urlparse(self.path).query)
                month_filter = (qs.get("month", [""])[0] or "").strip()
                name_filter = (qs.get("name", [""])[0] or "").strip()
                msg_filter = (qs.get("msg", [""])[0] or "").strip()
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

                        <div class="filter-field grow">
                          <label>Name / Email</label>
                          <div class="filter-search">
                            <i class="fa fa-magnifying-glass"></i>
                            <input type="text" name="name"
                                   placeholder="Search uploader username/email"
                                   value="{html.escape(name_filter, quote=True)}">
                          </div>
                        </div>

                        <div class="filter-actions">
                          <button class="btn small" type="submit">
                            <i class="fa fa-filter"></i> Apply
                          </button>

                          <a class="btn outline small" href="{html.escape(path, quote=True)}">
                            <i class="fa fa-rotate-left"></i> Reset
                          </a>
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
                            f"<i class='fa fa-lock'></i> Submitted (Locked)"
                            f"</span>"
                        )

                    # Delete button moved into timesheet header (swap)
                    admin_delete_html = ""
                    can_delete = False
                    if role == "Admin":
                        can_delete = True
                    else:
                        try:
                            up_date = datetime.datetime.strptime(uploaded_at, "%Y-%m-%d %H:%M:%S")
                            if (datetime.datetime.now() - up_date).total_seconds() <= 2 * 86400:
                                can_delete = True
                        except:
                            pass

                    if can_delete:
                        admin_delete_html = (
                            "<form method='post' action='/upload/delete' style='display:inline;margin:0' "
                            "onsubmit=\"return confirm('Delete this upload?');\">"
                            f"<input type='hidden' name='upload_id' value='{upload_id}'>"
                            f"<input type='hidden' name='return_to' value='{html.escape(path, quote=True)}'>"
                            "<button class='btn danger small' type='submit'>Delete</button>"
                            "</form>"
                        )

                    if module_ctx == "Email":
                        timesheet_html = f"<div style='margin-top:10px'>{admin_delete_html}</div>"
                    else:
                        timesheet_html = build_timesheet_ui(upload_id, path, admin_delete_html=admin_delete_html)

                    # Right side of upload header now shows Submitted badge (swap)
                    right_header = ""
                    if submitted_badge:
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

                    image_col_html = (
                        "<div style='display:flex; flex-direction:column; max-width:140px;'>"
                        f"<a class='thumb-link' href='{html.escape(file_url, quote=True)}' "
                        f"data-url='{html.escape(file_url, quote=True)}' "
                        f"data-filename='{html.escape(original_fn, quote=True)}' style='display:block;'>"
                        f"<img src='{html.escape(file_url, quote=True)}' alt='screenshot'></a>"
                        "</div>"
                    )

                    uploads_html += (
                        "<div class='upload-item'>"
                        f"{image_col_html}"
                        "<div class='upload-meta'>"
                        "<div class='upload-head'>"
                        f"<div><strong>{html.escape(original_fn)}</strong> {badge}</div>"
                        f"{right_header}"
                        "</div>"
                        f"<div>Uploaded at: {html.escape(uploaded_at)}</div>"
                        f"{uploader_line}"
                        f"{timesheet_html}"
                        "</div></div>"
                    )
                uploads_section = uploads_html if uploads_html else "<div class='muted'>No captures yet.</div>"
                capture_ui = build_capture_ui(module_ctx)
                modal_ui = build_saved_preview_modal()
                ts_js = build_timesheet_js()

                if msg_filter == "verify_req":
                    alert_script = "<script>alert('Attach verification screenprint below and submit to lock this timesheet.');</script>"
                else:
                    alert_script = ""

                content = (
                    f"{alert_script}"
                    "<div class='card' style='max-width:1000px;margin:0 auto'>"
                    f"{capture_ui}"
                    f"{filter_ui}"
                    f"<h2 style='margin-top:20px'>{html.escape(module_ctx)} Captures</h2>"
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
                if role == "Employee":
                    return self._msg("Access denied: your role does not have permission to access Email Settings.", display_name)

                form_html = (
                    "<div class='card' style='max-width:720px;margin:0 auto'>"
                    "<h2>Email Reminder Settings</h2>"
                    "<form method='post' action='/email-settings'>"
                    "<label>Recipient (email)</label><input type='email' name='recipient' required>"
                    "<label>Subject</label><input type='text' name='subject' required>"
                    "<label>Message</label><textarea name='message' required></textarea>"
                    "<label>Scheduled date</label><input type='date' name='date' required>"
                    "<button class='btn' type='submit'>Save Reminder</button>"
                    "</form></div>"
                )

                rows = list_email_reminders()
                reminders_html = ""
                if rows:
                    reminders_html = (
                        "<div class='card' style='max-width:900px;margin:18px auto'>"
                        "<h2>Saved Reminders</h2>"
                        "<table style='width:100%;border-collapse:collapse;margin-top:10px'>"
                        "<thead><tr style='text-align:left'>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>ID</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Recipient</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Subject</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Date</th>"
                        "<th style='padding:8px;border-bottom:1px solid #e6e9ee'>Created at</th>"
                        "</tr></thead><tbody>"
                    )
                    for rid, recip, subj, _msg_body, sched, _created_by, created_at in rows:
                        reminders_html += (
                            "<tr>"
                            f"<td style='padding:8px'>{rid}</td>"
                            f"<td style='padding:8px'>{html.escape(recip)}</td>"
                            f"<td style='padding:8px'>{html.escape(subj)}</td>"
                            f"<td style='padding:8px'>{html.escape(sched)}</td>"
                            f"<td style='padding:8px'>{html.escape(created_at)}</td>"
                            "</tr>"
                        )
                    reminders_html += "</tbody></table></div>"

                full = form_html + reminders_html
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(full, display_name, role, screen_access).encode("utf-8"))
                return
            # Export Data (Admin/Manager)
            if path == "/export":
                if role == "Employee":
                    return self._msg("Access denied: your role does not have permission to access Export Data.", display_name)
                qs = parse_qs(urlparse(self.path).query)
                month = (qs.get('month', [''])[0] or '').strip()
                if not re.match(r'^\d{4}-\d{2}$', month):
                    month = datetime.datetime.utcnow().strftime('%Y-%m')
                info = 'Download month-wise Excel with grouped week headers (Week1-Week4) and PPM/NTT totals. Late submissions are highlighted in red (PPM after Wed 3:00 PM, NTT after Fri 3:00 PM).'
                content = f"""
                <div class='card' style='max-width:900px;margin:0 auto'>
                  <h2>Export Data</h2>
                  <p class='muted'>{html.escape(info)}</p>
                  <form method='post' action='/export'>
                    <label>Select Month</label>
                    <input type='month' name='month' value='{html.escape(month, quote=True)}' required>
                    <button class='btn' type='submit'><i class='fa fa-file-excel'></i> Download Excel</button>
                  </form>
                  <div class='muted' style='margin-top:10px'>
                    Week buckets are the first 4 week-starts (Sunday) overlapping the selected month. If there are more than 4, extra weeks are merged into Week4.
                  </div>
                </div>
                """
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(render_page(content, display_name, role, screen_access).encode("utf-8"))
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
                    return self._msg("Invalid email or password.", "Visitor")

                db_password = row[3]
                db_status = row[5]
                if db_status != "Active":
                    return self._msg("Account is inactive. Please contact administrator.", "Visitor")

                if db_password == password:
                    sid = make_session(email_in)
                    self.send_response(303)
                    self.send_header("Location", "/")
                    self.send_header("Set-Cookie", f"session_id={sid}; Path=/; HttpOnly")
                    self.end_headers()
                    return

                return self._msg("Invalid email or password.", "Visitor")

            if not session_email:
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            # HARD BLOCK: Manager no access to user management POST
            if role == "Manager" and path.startswith("/user-management"):
                return self._msg("Access denied: your role does not have permission to access User Management.", display_name)

            # CREATE USER
            if path == "/user-management/create":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can create users.", display_name)

                required = ("username", "email", "password", "role", "status")
                if not all(form.get(k) for k in required):
                    return self._msg("Missing required fields for Create User.", display_name)

                ok = create_user_db(
                    form["username"], form["email"], form["password"],
                    form["role"], form["status"],
                    form.get("screen_access", "BOTH")
                )
                return self._msg("User created successfully." if ok else "Username or email already exists.", display_name)

            # UPDATE USER
            if path == "/user-management/update":
                if role != "Admin":
                    return self._msg("Access denied: only Admin can update users.", display_name)

                user_id = form.get("user_id")
                if not user_id:
                    return self._msg("User ID is required for update.", display_name)

                updates = {k: form.get(k) for k in ("username", "email", "role", "status", "screen_access") if form.get(k) is not None}
                if form.get("password"):
                    updates["password"] = form.get("password")

                if not updates:
                    return self._msg("No changes provided.", display_name)

                ok = update_user_db(int(user_id), updates)
                if ok:
                    self.send_response(303)
                    self.send_header("Location", "/user-management/list")
                    self.end_headers()
                    return
                return self._msg("Update failed: username or email may already exist.", display_name)

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
            if path in ("/upload/PPM", "/upload/NTT", "/upload/Email"):
                f = form.get("file")
                if not f or not isinstance(f, dict) or not f.get("filename"):
                    return self._msg("No captured image received.", display_name)

                module_val = "Email" if path == "/upload/Email" else ("PPM" if path == "/upload/PPM" else "NTT")

                if role == "Manager" and module_val == "Email":
                    return self._msg("Access denied: you do not have upload permission to Email.", display_name)

                if role == "Employee":
                    if module_val == "PPM" and screen_access not in ("PPM", "BOTH"):
                        return self._msg("Access denied: you do not have permission to upload to PPM.", display_name)
                    if module_val == "NTT" and screen_access not in ("NTT", "BOTH", "EMAIL_NTT"):
                        return self._msg("Access denied: you do not have permission to upload to NTT.", display_name)
                    if module_val == "Email" and screen_access != "EMAIL_NTT":
                        return self._msg("Access denied: you do not have permission to upload to Email.", display_name)

                user_disp = get_display_name_by_email(session_email) or (session_email or "user")
                if "@" in user_disp:
                    user_disp = user_disp.split("@", 1)[0]
                user_safe = _safe_name_part(user_disp, default="user")

                now = datetime.datetime.now()
                date_str = now.strftime("%Y%m%d")
                time_str = now.strftime("%H%M%S")

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

                # PPM/NTT: OCR extract Total-row hours, save, and auto-lock (disable edits)
                if module_val in ("PPM", "NTT"):
                    try:
                        if module_val == "PPM":
                            parsed = extract_ppm_hours_from_screenshot_total_row(save_path)
                        else:
                            parsed = extract_ntt_hours_from_screenshot(save_path)
                            
                        if parsed:
                            ws = parsed.get("week_start") or _week_start_from_iso_datetime(datetime.datetime.utcnow().isoformat(sep=" ", timespec="seconds"))
                            sun = parsed.get("sun", 0.0)
                            mon = parsed.get("mon", 0.0)
                            tue = parsed.get("tue", 0.0)
                            wed = parsed.get("wed", 0.0)
                            thu = parsed.get("thu", 0.0)
                            fri = parsed.get("fri", 0.0)
                            sat = parsed.get("sat", 0.0)
                            rem_plan = parsed.get("remaining_planned", 0.0)
                            upsert_timesheet(upload_id, ws, mon, tue, wed, thu, fri, sat, sun, remaining_planned=rem_plan)
                            logging.info("%s OCR autofill ok upload_id=%s week_start=%s", module_val, upload_id, ws)
                    except Exception:
                        logging.exception("%s OCR extraction failed (non-fatal).", module_val)


                self.send_response(303)
                if module_val == "PPM":
                    self.send_header("Location", f"/upload/{module_val}?msg=verify_req")
                else:
                    self.send_header("Location", f"/upload/{module_val}")
                self.end_headers()
                return

            # DELETE CAPTURE
            if path == "/upload/delete":
                upload_id = form.get("upload_id")
                return_to = normalize_return_to(form.get("return_to") or "/upload/PPM", role, session_email)

                if not upload_id:
                    return self._msg("Upload ID is required to delete.", display_name)

                uid = int(upload_id)
                if role != "Admin":
                    owner = get_upload_owner(uid)
                    if not owner or owner.lower() != session_email.lower():
                        return self._msg("Access denied: you can only delete your own uploads.", display_name)
                    
                    with closing(sqlite3.connect(DB_FILE)) as conn:
                        cur = conn.cursor()
                        cur.execute("SELECT uploaded_at FROM uploads WHERE id = ?", (uid,))
                        upload_row = cur.fetchone()
                        
                    if upload_row:
                        try:
                            up_date = datetime.datetime.strptime(upload_row[0], "%Y-%m-%d %H:%M:%S")
                            if (datetime.datetime.now() - up_date).total_seconds() > 2 * 86400:
                                return self._msg("Access denied: you can only delete uploads within 2 days.", display_name)
                        except:
                            pass

                delete_upload(int(upload_id))
                self.send_response(303)
                self.send_header("Location", return_to)
                self.end_headers()
                return

            # VERIFICATION UPLOAD
            if path == "/timesheet/verify":
                upload_id = form.get("upload_id")
                return_to = normalize_return_to(form.get("return_to") or "/upload/PPM", role, session_email)

                if not upload_id:
                    return self._msg("Upload ID is required.", display_name)

                f = form.get("verification_file")
                if not f or not isinstance(f, dict) or not f.get("filename"):
                    return self._msg("No file selected for verification.", display_name)

                user_disp = get_display_name_by_email(session_email) or (session_email or "user")
                if "@" in user_disp:
                    user_disp = user_disp.split("@", 1)[0]
                user_safe = _safe_name_part(user_disp, default="user")

                now = datetime.datetime.now()
                base, ext = os.path.splitext(f["filename"])
                ext = ext or ".png"
                timestamp = now.strftime("%Y%m%d%H%M%S")
                safe_base = re.sub(r"[^a-zA-Z0-9._-]+", "_", base)[:80] or "verification"
                stored_fn = f"verify_{user_safe}_{safe_base}_{timestamp}{ext}".replace(" ", "_")
                save_path = os.path.join(UPLOAD_DIR, stored_fn)

                counter = 0
                while os.path.exists(save_path):
                    counter += 1
                    stored_fn = f"verify_{user_safe}_{safe_base}_{timestamp}_{counter}{ext}"
                    save_path = os.path.join(UPLOAD_DIR, stored_fn)

                with open(save_path, "wb") as out:
                    out.write(f["content"])

                save_verification_filename(int(upload_id), stored_fn)
                mark_timesheet_submitted(int(upload_id))
                
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
                # Block manual edits for auto-filled timesheets (from screenshot)
                if get_upload_module(int(upload_id)) in ("PPM", "NTT"):
                    return self._msg("PPM & NTT timesheets are auto-filled and cannot be edited.", display_name)


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

                week_start, mon, tue, wed, thu, fri, sat, sun, total = parse_timesheet_payload()

                # We now allow submitting even if total is 0, for verification purposes

                owner = get_upload_owner(uid) or (session_email or "").strip().lower()

                try:
                    submit_week_group(uid, owner, week_start, mon, tue, wed, thu, fri, sat, sun)
                except ValueError as ve:
                    return self._msg(str(ve), display_name)

                self.send_response(303)
                self.send_header("Location", return_to)
                self.end_headers()
                return

            # EMAIL SETTINGS
            if path == "/email-settings":
                if role == "Employee":
                    return self._msg("Access denied: your role does not have permission to access Email Settings.", display_name)

                recipient = (form.get("recipient") or "").strip()
                subject = (form.get("subject") or "").strip()
                message = (form.get("message") or "").strip()
                date_field = (form.get("date") or "").strip()

                if not (recipient and subject and message and date_field):
                    return self._msg("All fields are required for email reminder.", display_name)

                try:
                    datetime.date.fromisoformat(date_field)
                except Exception:
                    return self._msg("Invalid date format.", display_name)

                add_email_reminder(recipient, subject, message, date_field, created_by=session_email)
                self.send_response(303)
                self.send_header("Location", "/email-settings")
                self.end_headers()
                return
            # EXPORT (Admin/Manager): month-wise Excel download
            if path == "/export":
                if role == "Employee":
                    return self._msg("Access denied: your role does not have permission to access Export Data.", display_name)
                month = (form.get('month') or '').strip()
                if not re.match(r'^\d{4}-\d{2}$', month):
                    return self._msg("Invalid month. Please select a month (YYYY-MM).", display_name)
                try:
                    xlsx_bytes = generate_monthly_export_xlsx(month)
                except Exception as ex:
                    return self._msg(f"Export failed: {html.escape(str(ex))}", display_name)
                filename = f"export_{month}.xlsx"
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Content-Length', str(len(xlsx_bytes)))
                self.end_headers()
                self.wfile.write(xlsx_bytes)
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
    logging.info("Starting server...")
    logging.info(f"URL: http://{HOST}:{PORT}")
    try:
        HTTPServer((HOST, PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        logging.info("Server stopped by user")
    except Exception:
        logging.exception("Server crashed")
