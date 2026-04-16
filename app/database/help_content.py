"""app/database/help_content.py — Help-page content CRUD."""
import sqlite3
import re
import logging
from .connection import get_conn

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




