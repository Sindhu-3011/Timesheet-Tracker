# Timesheet Tracker

## Project Overview & Scope
The **Timesheet Tracker** is a self-contained, lightweight Python web application built into a single server file. Its primary goal is to provide a central platform for employees to upload screenshots of their time entries directly from external systems, like PPM or NTT. 

By leveraging Optical Character Recognition (OCR), the system intelligently attempts to extract the hours worked across a given week. Employees can then submit their timesheets, and administrators/managers can oversee submissions, configure reminders, and export the logs to reports.

### Key Features
- **Single-File Server Application**: Runs on the built-in standard Python `http.server`, completely avoiding heavy web framework dependencies.
- **Image Uploads & OCR Extraction**: Automatically extracts inputted timesheet hours directly from PPM and NTT interface screenshots using `EasyOCR` and `Pillow` image manipulation.
- **Role-Based Access Control**:
  - **Admin**: Has full access, can manage users, oversee all timesheet submissions, delete entries (even after submission), configure email reminders, and export data.
  - **Manager**: Has managerial access over submissions, help viewing, and export tools but lacks admin user management and help editing rights.
  - **Employee**: Can view, track, and upload timesheets for their own account, and access the Help page content.
- **Help Center Access**: Integrated Help page where all users can view guidance, while editing capabilities are strictly restricted to Administrators.
- **Automated Capture Workflow**: Streamlined "One-Click" capture process for PPM, NTT, and EMAIL modules that automatically takes, shows, and saves screen prints without requiring manual save confirmation.
- **Strict Timesheet Verification Rules**: Requires that if an employee is working on both PPM and NTT, the hours correctly mirror each other before allowing a final locked submission. 
- **Export Capabilities**: Converts submitted tracking timelines directly to cleanly formatted `.xlsx` Excel spreadsheets utilizing `openpyxl`.
- **Automated Database Setup**: Automatically builds and queries a local `users.db` SQLite database using standardized schema handling. 

---

## Technical Stack & Logic
- **Backend Core**: Python 3 standard library
- **Database**: SQLite3 (`users.db`)
- **Image Processing Engine**: `easyocr` & `Pillow` (`PIL`)
- **Exporting**: `openpyxl`

## How to Run the Application

### 1. Prerequisites
Ensure you have **Python 3.7+** installed on your system.
It is highly recommended that you run this in a Python Virtual Environment.

### 2. Install Required Dependencies
All 3rd-party dependencies revolve around the image OCR and exporting capabilities. Install them using `pip`:

```bash
pip install Pillow easyocr openpyxl numpy
```

### 3. Start the Server
Navigate to the root directory where `test.py` is located, and execute:

```bash
python test.py
```

*Note: The script actively checks for the `uploads/` directory and `users.db`. It will transparently create them upon the first run if they are not already present.*

### 4. Access the Web Dashboard
Open your preferred browser and connect to the local server socket address prompted in your terminal:

```
http://127.0.0.1:8000
```

### 5. Initial Login Setup
Log in using the default built-in Administrative credentials set in the project script configuration parameters:

- **Email**: `[EMAIL_ADDRESS]`
- **Password**: `Password`

*Administrators should eventually rotate these credentials or onboard themselves properly to ensure secure deployments.*
