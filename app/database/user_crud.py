"""app/database/user_crud.py — User query / update helpers."""
import sqlite3
import hashlib
import logging
from .connection import get_conn
from ..config import utc_now_str, normalize_screen_access

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
