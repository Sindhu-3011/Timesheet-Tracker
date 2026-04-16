"""app/database/custom_fields.py — Custom field definitions and form helpers."""
import sqlite3
import json
import html
import logging
from .connection import get_conn
from ..config import utc_now_str

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


