"""app/ocr/ppm.py — PPM screenshot OCR extraction."""
import re
import logging
import datetime
from .engine import pytesseract, Output, _preprocess_for_ocr, MONTHS_MAP
from ..config import UTC

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
