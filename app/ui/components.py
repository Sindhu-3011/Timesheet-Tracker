"""app/ui/components.py — Reusable UI widgets (timesheet grid, capture UI, preview modal)."""
import html
import json
import datetime
import logging
import sqlite3

from ..config import UTC, utc_now_str, UPLOAD_DIR
from ..database.uploads import _sunday_of_date
from ..database.connection import get_conn
from ..database.uploads import (
    get_timesheet_by_upload, get_timesheet_values,
    ensure_timesheet_for_upload, get_upload_module,
    is_timesheet_submitted,
)

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
        ensure_timesheet_for_upload(upload_id, utc_now_str(sep=" ", timespec="seconds"))
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
    # Separate disable flags: hours may be locked while week_start stays editable until submission
    hours_disabled_attr = "disabled" if is_submitted else ""
    week_disabled_attr = "disabled" if is_submitted else ""
    # PPM & NTT: auto-extracted hours; prevent manual edits of hours, but allow week_start changes until submitted
    auto_locked_hours = (module_for_upload in ("PPM", "NTT"))
    if auto_locked_hours:
        hours_disabled_attr = "readonly"
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

    ppm_is_verify_attached = False
    verify_html = ""
    submit_disabled_attr = ""
    
    if module_for_upload == "PPM":
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT id, stored_filename FROM uploads WHERE module='VERIFY' AND original_filename=?", (f"PPM_{upload_id}",))
        v_row = cur.fetchone()
        conn.close()
        if v_row:
            ppm_is_verify_attached = True
            v_img = f"/uploads/{html.escape(v_row[1], quote=True)}"
            verify_html = f"""
            <div style="margin-top:12px; padding:10px 14px; border-radius:8px; background:linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%); border:1px solid #6ee7b7; display:flex; align-items:center; gap:10px;">
                <div style="width:32px;height:32px;border-radius:50%;background:#10b981;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                  <i class="fa fa-check" style="color:#fff;font-size:14px;"></i>
                </div>
                <span style="color:#065f46;font-weight:600;font-size:13px;">Confirmation Screen Print Attached</span>
            </div>
            """
        else:
            verify_html = f"""
            <div style="margin-top:12px; padding:10px 14px; border-radius:8px; background:linear-gradient(135deg, #fef2f2 0%, #fde8e8 100%); border:1px solid #fca5a5;">
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                <div style="width:32px;height:32px;border-radius:50%;background:#ef4444;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                  <i class="fa fa-exclamation" style="color:#fff;font-size:14px;"></i>
                </div>
                <span style="color:#991b1b; font-weight:600; font-size:13px;">Attach confirmation screen print and click on submit.</span>
              </div>
              <button type="button" class="btn small outline" onclick="captureVerify({upload_id}, this)" {'disabled' if is_submitted else ''}
                style="border-color:#ef4444; color:#dc2626; font-size:12px;">
                <i class="fa fa-camera"></i> Capture Confirmation Screen Print
              </button>
            </div>
            """
            
    if module_for_upload == "PPM" and not ppm_is_verify_attached:
        submit_disabled_attr = "disabled"

    has_draft_error = ""
    if module_for_upload == "NTT" and int(ts.get("has_draft", 0)) == 1:
        submit_disabled_attr = "disabled"
        has_draft_error = f"""
        <div style="margin-top:12px; padding:12px 16px; border-radius:10px; background:#fff1f2; border:1px solid #fecaca; color:#991b1b; display:flex; align-items:center; gap:12px; font-weight:600; font-size:14px; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
            <i class="fa fa-triangle-exclamation" style="font-size:18px; color:#ef4444;"></i>
            <span>NTT status is in Draft. Please submit the timesheet in NTT portal and re-capture the screenshot.</span>
        </div>
        """

    rem_planned = ts.get("remaining_planned")
    rem_warning = ""
    if module_for_upload == "PPM" and rem_planned is not None:
        try:
            if float(rem_planned) < 45.0:
                rem_warning = f"""
                <div style="margin-top:12px; padding:12px 16px; border-radius:10px; background:#fff1f2; border:1px solid #fecaca; color:#991b1b; display:flex; align-items:center; gap:12px; font-weight:600; font-size:14px; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
                    <i class="fa fa-triangle-exclamation" style="font-size:18px; color:#ef4444;"></i>
                    <span>Warning: Your Remaining Planned hours are less than 45 hrs. Please contact your manager to top up your hours for next week.</span>
                </div>
                """
        except (ValueError, TypeError):
            pass

    if not is_submitted:
        save_btn_html = ""
        if module_for_upload not in ("PPM", "NTT"):
            save_btn_html = """<button class="btn small" type="submit">
            <i class="fa fa-save"></i> Save Hours
          </button>"""
        controls_html = f"""
        {rem_warning}
        {has_draft_error}
        {verify_html}
        <div class="controls" style="margin-top:12px;">
          {save_btn_html}

          <button class="btn secondary small ts-submit-btn" type="submit" formaction="/timesheet/submit" {submit_disabled_attr}>
            <i class="fa fa-check"></i> Submit
          </button>
        </div>
        """
    else:
        controls_html = f"""
        {rem_warning}
        <div class="muted" style="margin-top:8px">
          This timesheet is locked. Contact Admin if changes are required.
        </div>
        """

    # these are defined above
    ts_head_right_extra = admin_delete_html or ""

    return f"""
<details class="ts-details">
  <summary>
    Timesheet (Week-wise) — <span id="ts_range_{upload_id}">{html.escape(week_range_text)}</span>
  </summary>

  <div class="ts-wrap ts-form" data-upload-id="{upload_id}">
    <div class="ts-head">
      <div class="muted" style="font-weight:700;">
        Week Start (Sun):
        <input class="ts-week" type="date" value="{week_start_iso}"
               style="width:auto;display:inline-block;margin-left:8px;" {week_disabled_attr}>
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
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="sun" value="{sun}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="mon" value="{mon}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="tue" value="{tue}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="wed" value="{wed}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="thu" value="{thu}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="fri" value="{fri}" {hours_disabled_attr}></td>
            <td><input class="ts-day" type="number" step="0.25" min="0" max="24" name="sat" value="{sat}" {hours_disabled_attr}></td>
            <td>
              <input type="text" readonly value="{total}" class="ts-total-input" id="ts_total_in_{upload_id}">
            </td>
          </tr>
        </tbody>
      </table>

      {controls_html}
    </form>
  </div>
</details>
"""


def build_timesheet_js():
    return """
<script>
(function(){
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  window.captureVerify = async function(uploadId, btn) {
    try {
      btn.disabled = true;
      const stream = await navigator.mediaDevices.getDisplayMedia({
        video: { cursor: 'always' },
        audio: false
      });
      const track = stream.getVideoTracks()[0];
      const v = document.createElement('video');
      v.srcObject = stream;
      v.autoplay = true;
      v.muted = true;
      v.playsInline = true;
      await v.play();
      
      // small delay to let user settle window
      await new Promise(r => setTimeout(r, 2000));
      
      const w = v.videoWidth || 1280;
      const h = v.videoHeight || 720;
      const c = document.createElement('canvas');
      c.width = w; c.height = h;
      const ctx = c.getContext('2d');
      ctx.drawImage(v, 0, 0, w, h);
      
      c.toBlob(async (b) => {
        stream.getTracks().forEach(t=>t.stop());
        if(!b) { alert('Could not create image.'); btn.disabled = false; return; }
        
        btn.innerHTML = "<i class='fa fa-spinner fa-spin'></i> Uploading...";
        const fd = new FormData();
        fd.append('file', b, 'Verify_capture_' + Date.now() + '.png');
        fd.append('module', 'VERIFY');
        fd.append('parent_upload_id', uploadId);
        
        const res = await fetch('/upload/VERIFY', { method:'POST', body:fd });
        if(res.ok) {
           const form = btn.closest('.ts-form');
           if(form) {
               const submitBtn = form.querySelector('.ts-submit-btn');
               if (submitBtn) submitBtn.disabled = false;
           }
           const container = btn.parentElement;
           container.style.background = 'linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%)';
           container.style.borderColor = '#6ee7b7';
           container.style.display = 'flex';
           container.style.alignItems = 'center';
           container.style.gap = '10px';
           container.innerHTML = `
                <div style="width:32px;height:32px;border-radius:50%;background:#10b981;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
                  <i class="fa fa-check" style="color:#fff;font-size:14px;"></i>
                </div>
                <span style="color:#065f46;font-weight:600;font-size:13px;">Confirmation Screen Print Attached</span>
           `;
        } else {
           alert('Upload failed: ' + res.statusText);
           btn.disabled = false;
           btn.innerHTML = "<i class='fa fa-camera'></i> Capture Confirmation Screen Print";
        }
      }, 'image/png');
    } catch(e) {
      console.error(e);
      alert('Capture cancelled or error.');
      btn.disabled = false;
    }
  };

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
      submitBtn.disabled = (total <= 0);
      submitBtn.title = submitBtn.disabled
        ? 'Enter hours (Total must be > 0) before submitting.'
        : 'Submit and lock this week (PPM and NTT must match).';
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
})();
</script>
"""


# -------------------------
# Capture UI builder — PPM/NTT only (NO download for unsaved)
# -------------------------
def build_capture_ui(module_ctx: str):
    title = f"Screen Print Capture - {"Email" if module_ctx=="EMAIL" else module_ctx}"
    upload_url = f"/upload/{module_ctx}"
    hint = f"This page captures and lists only {module_ctx} screen prints."
    capture_label = f"Capture {module_ctx} Screen Print"

    parts = []
    parts.append(f"<h2>{html.escape(title)}</h2>")
    parts.append(
        "<p class='muted' style='margin-bottom:8px'>"
        "Click <b>Capture</b> to select a window/tab/screen. A snapshot will be automatically saved without clicking any save button. "
        + html.escape(hint) +
        "</p>"
    )

    parts.append(
        "<div style='display:flex;gap:12px;flex-wrap:wrap;align-items:center'>"
        f"<button type='button' class='btn' id='btnCapture'>{html.escape(capture_label)}</button>"
        "</div>"
    )

    parts.append("<video id='capPreview' style='display:none' autoplay muted playsinline></video>")

    # For PPM & NTT, hide Save and Close buttons entirely (auto-upload handles it)
    save_btn_style = "display:none" if module_ctx in ("PPM", "NTT", "EMAIL") else ""
    close_btn_style = "display:none" if module_ctx in ("PPM", "NTT", "EMAIL") else ""
    parts.append(
        "<div id='shotContainer' class='preview' style='display:none;margin-top:12px'>"
        "  <div id='shotBox'></div>"
        "  <div class='controls'>"
        f"    <button type='button' class='btn' id='btnSave' style='{save_btn_style}'><i class='fa fa-cloud-upload-alt'></i> Save</button>"
        f"    <button type='button' class='btn danger' id='btnCloseShot' style='{close_btn_style}'><i class='fa fa-times'></i> Close</button>"
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

  const AUTO_UPLOAD = (moduleName === "PPM" || moduleName === "NTT" || moduleName === "EMAIL");

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
             // Auto-upload for PPM and NTT to avoid manual Save click
             if(AUTO_UPLOAD){{
               setTimeout(() => {{ try{{ saveShot(); }}catch(e){{}} }}, 50);
             }}
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

  // Auto-upload: enabled for PPM and NTT; Save remains for re-upload if needed
  if(AUTO_UPLOAD){{
    try{{ btnSave.style.display = 'none'; }}catch(e){{}}
  }}
  btnCloseShot.onclick = clearShot;
}})();
</script>
"""
    parts.append(js)
    return "".join(parts)


