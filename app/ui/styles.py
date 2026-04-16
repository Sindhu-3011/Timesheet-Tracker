"""app/ui/styles.py — Global CSS string (PAGE_STYLE)."""

PAGE_STYLE = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<style>
:root{--bg:#f4f6f8;--nav:#0f1724;--accent:#2563eb;--card:#ffffff;--muted:#64748b}
*{box-sizing:border-box}
body{margin:0;font-family:'Inter',"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:#1e293b;font-size:14px;line-height:1.5}
header.app-header{position:fixed;top:0;left:0;right:0;height:64px;background:linear-gradient(90deg,var(--nav),#0b1220);color:#fff;display:flex;align-items:center;justify-content:space-between;padding:0 20px;z-index:1000;box-shadow:0 2px 8px rgba(2,6,23,0.15)}
.brand{display:flex;align-items:center;gap:12px;font-weight:700}
.brand .logo{width:36px;height:36px;border-radius:8px;background:linear-gradient(135deg,var(--accent),#7c3aed);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;box-shadow:0 2px 6px rgba(37,99,235,0.2)}
.header-right{display:flex;align-items:center;gap:12px;font-weight:600}
.layout{display:flex;margin-top:64px;min-height:calc(100vh - 64px)}
aside.sidebar{width:200px;background:#0b1220;color:#e6eef8;padding:12px 8px;border-right:1px solid rgba(255,255,255,0.03);position:sticky;top:64px;height:calc(100vh - 64px)}
.nav-item{display:flex;align-items:center;gap:12px;padding:10px 12px;border-radius:8px;color:inherit;text-decoration:none;margin-bottom:6px;font-weight:600;position:relative}
.nav-item i{width:20px;text-align:center;color:#9fb3d8}
.nav-item:hover{background:rgba(255,255,255,0.04);cursor:pointer}
.dropdown-content{display:none;position:absolute;top:100%;left:0;background:#fff;min-width:220px;box-shadow:0 8px 20px rgba(2,6,23,0.12);border-radius:8px;overflow:hidden;z-index:999}
.nav-item:hover .dropdown-content{display:block}
.dropdown-content a{display:block;padding:10px 12px;color:#111827;text-decoration:none}
.dropdown-content a:hover{background:#f4f6f8}
@media(max-width:880px){aside.sidebar{display:none}main.content{padding:12px}}
main.content{flex:1;padding:16px}
.card{background:var(--card);border-radius:12px;padding:20px;box-shadow:0 6px 18px rgba(15,23,42,0.06)}
.card h2{margin:0 0 12px;font-size:18px;color:#0f1724}
.muted{color:var(--muted);font-size:14px}
label{display:block;font-weight:600;margin-top:12px;color:#111827}
input,select,textarea{width:100%;padding:10px 12px;margin-top:8px;border:1px solid #e2e8f0;border-radius:10px;background:#fff;font-size:14px;color:#0f1724 !important;outline:none !important;-webkit-appearance:none;appearance:none;font-weight:500;transition:border-color 0.2s, box-shadow 0.2s;text-align:left;}
input:focus,select:focus,textarea:focus,*:focus{outline:none !important;border-color:var(--accent) !important;box-shadow:0 0 0 3px rgba(37,99,235,0.1) !important;}
input:disabled,select:disabled,textarea:disabled,input:read-only,select:read-only,textarea:read-only{background:#f8fafc !important;color:#64748b !important;border-color:#e2e8f0;cursor:default;opacity:1 !important;-webkit-text-fill-color:#64748b !important;}
select option { color: #0f1724 !important; background: #fff !important; font-weight:500; }
textarea{min-height:110px;resize:vertical}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;background:var(--accent);color:#fff;padding:10px 18px;border-radius:10px;border:none;font-weight:600;font-size:14px;cursor:pointer;text-decoration:none;transition:all 0.2s;line-height:1;margin-top:16px;font-family:inherit;}
.btn:hover{filter:brightness(1.1);transform:translateY(-1px)}
.btn:active{transform:translateY(0)}
.btn.secondary{background:#1e293b;color:#fff}
.btn.danger{background:#ef4444}
.btn:disabled{opacity:.5;cursor:not-allowed;transform:none}
.small{padding:8px 12px;font-size:13px;border-radius:8px}
.preview{margin-top:12px;border:1px solid #e6e9ee;padding:10px;border-radius:8px;background:#fafafa}
.controls{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}


/* ===== User List Filter (compact) ===== */
.um-filter select{height:32px;padding:6px 10px;font-size:12px;border-radius:8px;}
.um-filter label{font-size:12px;}
.um-filter .btn.small,.um-filter a.btn.small{height:32px;padding:0 10px;font-size:12px;border-radius:8px;display:inline-flex;align-items:center;gap:6px;margin-top:0;}
.upload-item{display:flex;gap:12px;align-items:flex-start;padding:10px 0;border-bottom:1px dashed #f1f5f9}
.upload-item img{max-width:140px;border-radius:8px;border:1px solid #e6e9ee}
.upload-meta{font-size:13px;color:#374151;flex:1}
.badge{display:inline-block;font-size:11px;padding:2px 8px;border-radius:999px;margin-left:8px;background:#eef2ff;color:#4338ca;border:1px solid #c7d2fe;line-height:1.2}
.lock-badge{background:#ecfdf5;color:#065f46;border-color:#a7f3d0}

/* Header row within each upload item */
.upload-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}
.upload-actions{display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap}
.upload-actions form{margin:0}

/* ===== Same-page preview modal ===== */
.modal-overlay{
  position:fixed; inset:0;
  background:rgba(2,6,23,.75);
  display:none;
  align-items:center;
  justify-content:center;
  z-index:2000;
  padding:20px;
}
.modal{
  width:min(1100px, 96vw);
  max-height:90vh;
  background:#fff;
  border-radius:14px;
  overflow:hidden;
  box-shadow:0 20px 60px rgba(0,0,0,.35);
  display:flex;
  flex-direction:column;
}
.modal-header{
  display:flex;
  align-items:center;
  justify-content:space-between;
  padding:12px 14px;
  border-bottom:1px solid #e6e9ee;
  background:#f8fafc;
}
.modal-title{
  font-weight:800;
  font-size:14px;
  color:#0f1724;
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
  max-width:70%;
}
.modal-body{
  padding:14px;
  overflow:auto;
  background:#0b122012;
}
.modal-body img{
  width:100%;
  height:auto;
  border-radius:12px;
  border:1px solid #e6e9ee;
  background:#fff;
}
.modal-actions{display:flex;gap:10px;align-items:center;}

/* ===== Timesheet ===== */
.ts-wrap{margin-top:10px;border:1px solid #e6e9ee;border-radius:10px;background:#fff}
.ts-head{display:flex;gap:10px;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid #eef2f7;background:#f8fafc}
.ts-grid{width:100%;border-collapse:collapse}
.ts-grid th,.ts-grid td{padding:8px;border-bottom:1px solid #eef2f7;text-align:left;font-size:13px}
.ts-grid th{color:#0f1724;background:#fff;vertical-align:top}
.ts-grid td input{margin-top:0;padding:8px;border-radius:8px}
.ts-total{font-weight:900}
details.ts-details summary{cursor:pointer;user-select:none;font-weight:800;color:#111827;margin-top:8px}
details.ts-details{margin-top:8px}
.ts-grid th .ts-dd{
  display:block;
  font-size:12px;
  color:var(--muted);
  font-weight:800;
  margin-top:2px;
}

.filter-card{
  margin-top:14px;
  padding:14px 14px 12px;
  border:1px solid #e6e9ee;
  border-radius:14px;
  background:#fff;
  box-shadow:0 6px 18px rgba(15,23,42,0.05);
}
.filter-title{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:12px;
  margin-bottom:10px;
}
.filter-title .muted{margin:0;font-weight:700}
.filter-form{
  display:flex;
  gap:12px;
  align-items:flex-end;
  flex-wrap:wrap;
}
.filter-field{min-width:220px}
.filter-field.grow{flex:1;min-width:260px}
.filter-field label{
  margin:0;
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.05em;
  color:#64748b;
  font-weight:700;
}
.filter-field input,.filter-field select{
  margin-top:6px;
  height:40px;
  padding:8px 12px;
  border-radius:10px;
  background:#fff !important;
  color:#0f1724 !important;
  border:1px solid #cbd5e1;
}
.filter-search{position:relative}
.filter-search i{
  position:absolute;
  left:12px;
  top:50%;
  transform:translateY(-50%);
  color:#94a3b8;
  pointer-events:none;
}
.filter-search input{padding-left:38px}
.filter-actions{
  display:flex;
  gap:10px;
  align-items:center;
}

/* Buttons inside toolbars should not inherit the big top margin */
/* Filter action buttons: match input height and align icons */
.filter-actions .btn,
.filter-actions a.btn{
  height:40px;
  padding:0 16px;
  display:inline-flex;
  align-items:center;
  justify-content:center;
  gap:8px;
  border-radius:10px;
  white-space:nowrap;
  margin-top:0;
}
.filter-actions a.btn{ text-decoration:none; }
.filter-actions .btn.outline{ border-color:#cbd5e1; color:#475569; }
.filter-actions .btn.outline:hover{ background:#f1f5f9; color:#0f1724; border-color:#94a3b8; }

.controls .btn,
.filter-form .btn{margin-top:0}

/* Outline button variant */
.btn.outline{
  background:#fff;
  color:var(--accent);
  border:1px solid #cbd5e1;
}
.btn.outline:hover{background:#f8fafc}

/* Optional: make small buttons more pill-like */
.btn.small{border-radius:10px}

@media(max-width:880px){aside.sidebar{display:none}main.content{padding:12px}}

/* ===== User List Table (compact) ===== */
.um-table{font-size:12px;color:#334155;table-layout:auto;width:100%;border-spacing:0;}
.um-table th,.um-table td{padding:6px 4px !important;border-bottom:1px solid #f1f5f9;word-break:break-word;text-align:left;vertical-align:middle;}
.um-table th{font-size:11px;text-transform:uppercase;letter-spacing:0.02em;color:#64748b;font-weight:700;background:#f8fafc;}
.um-table select{height:28px;padding:2px 6px;font-size:12px;border-radius:6px;margin:0;}
.um-table .btn.small,.um-table a.btn.small{height:28px;padding:0 8px;font-size:11px;border-radius:6px;display:inline-flex;align-items:center;gap:4px;margin-top:0;}
.um-rowform{display:flex;gap:4px;align-items:center;flex-wrap:wrap;}

</style>
"""


