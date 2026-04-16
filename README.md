# Timesheet Tracker

## Overview
A self-contained Python web application for employee timesheet management. Employees upload PPM/NTT/EMAIL screen captures, hours are auto-extracted via OCR, and managers/admins oversee submissions, configure email reminders, and export data to Excel.

---

## Project Structure

```
Timesheet-Tracker/
├── tracker.py          # Main application (single-file HTTP server)
├── .env                # Local configuration — NOT committed to git
├── .env.example        # Configuration template (commit this)
├── .gitignore
├── requirements.txt    # Python dependencies
├── README.md
└── uploads/
    └── branding/       # Place logo/branding assets here
```

> **Database** (`users.db`) and **upload files** (`uploads/`) are created automatically on first run and are excluded from source control.

---

## Technical Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.9+ standard library (`http.server`) |
| Database | SQLite 3 (`users.db`) |
| OCR | `pytesseract` + `Pillow` |
| Export | `openpyxl` |
| Config | `python-dotenv` (`.env` file) |

---

## Roles & Access

| Role | Capabilities |
|---|---|
| **Admin** | Full access: user management, all screen modules, delete, email settings, export |
| **Manager** | Screen modules, email settings, export. No user management. |
| **Employee** | Own uploads only, assigned screen modules (PPM / NTT / EMAIL) |

---

## Setup

### 1. Prerequisites
- Python 3.9 or later
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed (Windows: note install path)

### 2. Create a virtual environment (recommended)
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# macOS / Linux
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure the environment
```bash
cp .env.example .env
```
Edit `.env` and fill in your values. The key settings are:

| Variable | Description | Default |
|---|---|---|
| `DEBUG` | Enable debug mode / tracebacks | `true` |
| `HOST` | Server bind address | `127.0.0.1` |
| `PORT` | Server port | `8000` |
| `UPLOAD_DIR` | Folder for uploaded screenshots | `uploads` |
| `DB_FILE` | SQLite database file path | `users.db` |
| `TESSERACT_DIR` | Path to Tesseract install (Windows) | *(blank = use PATH)* |
| `AUTO_LOCK_NTT_ON_OCR` | Auto-submit NTT after OCR fill | `true` |
| `DEFAULT_ADMIN_EMAIL` | Seed admin email (first run only) | `admin@example.com` |
| `DEFAULT_ADMIN_PASSWORD` | Seed admin password (first run only) | `admin` |
| `SMTP_HOST` | SMTP server for email reminders | *(blank = disabled)* |
| `SMTP_PORT` | SMTP port | `587` |
| `SMTP_USER` | SMTP username | |
| `SMTP_PASS` | SMTP password | |
| `SMTP_FROM` | Sender address | |
| `SMTP_SSL` | Use SSL (port 465) | `false` |

### 5. Run the server
```bash
python tracker.py
```

Open **http://127.0.0.1:8000** in your browser.

On the very first run the database is created automatically and seeded with the default admin account set in `.env`.

> **Important**: Change `DEFAULT_ADMIN_PASSWORD` in `.env` before deploying to any shared environment.

---

## Key Features

- **OCR autofill** — PPM and NTT screenshots are parsed automatically using Tesseract; hours are populated into the weekly timesheet.
- **Symmetry rule** — PPM and NTT hours for the same user/week must match before a final submit is allowed.
- **Auto-lock** — NTT timesheets are submitted and locked immediately after successful OCR extraction (`AUTO_LOCK_NTT_ON_OCR=true`).
- **Verification gate** — PPM requires a confirmation (VERIFY) screenshot before the Submit button is enabled.
- **Email reminders** — Scheduled via a background daemon thread; uses IST timezone for scheduling.
- **Excel export** — Week-bucketed `.xlsx` report with PPM/NTT columns and submission cutoff colour coding.
- **Custom user fields** — Admins can define extra profile fields (TEXT, NUMBER, DATE, BOOLEAN, DROPDOWN).
- **Help centre** — DB-backed HTML page; Admins can edit it via a WYSIWYG editor (Quill).

---

## Database Schema (auto-managed)

| Table | Purpose |
|---|---|
| `users` | Authentication, role, screen access, feature flags |
| `uploads` | Screen capture file records |
| `timesheets` | Weekly hours (Mon–Sun), submit/lock state |
| `email_reminders` | Scheduled SMTP reminders |
| `user_custom_fields` | Admin-defined extra profile field definitions |
| `user_custom_field_values` | Per-user values for custom fields |
| `help_content` | HTML body of the Help page |

All schema migrations run automatically at startup — the database is always up to date.

