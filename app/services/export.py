"""app/services/export.py — Monthly Excel export and filter UI."""
import io
import html
import json
import datetime
import logging
import sqlite3

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except Exception:
    Workbook = None
    Font = Alignment = Border = Side = PatternFill = None

from ..config import IST, UTC, utc_now_str
from ..database.connection import get_conn
from ..database.custom_fields import list_custom_fields, _parse_options_json

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
