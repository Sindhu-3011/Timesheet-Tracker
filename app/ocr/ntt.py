"""app/ocr/ntt.py — NTT screenshot OCR extraction."""
import re
import logging
import datetime
from .engine import pytesseract, Output, _preprocess_for_ocr
from ..config import UTC

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
