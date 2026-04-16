"""app/auth/sessions.py — In-memory session store and helpers."""
import uuid
import http.cookies
import logging
from ..database.user_crud import (
    get_user_db_by_email, get_screen_access_by_email,
)

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
