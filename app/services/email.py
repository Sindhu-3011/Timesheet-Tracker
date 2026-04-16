"""app/services/email.py — SMTP email sending + reminder scheduler."""
import smtplib
import email as _email_lib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sqlite3
import logging
import threading
import time
import datetime

from ..config import (
    SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM, SMTP_SSL,
    UTC, utc_now_str,
)
from ..database.connection import get_conn
from ..database.user_crud import list_active_users_for_reminders

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

EMAIL_SCHEDULER_STARTED = start_email_scheduler()
