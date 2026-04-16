"""app/database/uploads.py — Upload & timesheet CRUD."""
import sqlite3
import os
import re
import logging
import datetime
from .connection import get_conn
from ..config import UTC, utc_now_str

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


# ----- Upload & Timesheet CRUD -----
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


