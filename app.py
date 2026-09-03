"""
LN Risk Map — Productivity Tracker
------------------------------------------------
A Flask web app for logging daily productivity / risk-map entries per
employee and saving everything into an Excel workbook (date-named sheets).

Two kinds of accounts:
  - ADMIN (fixed credentials below): can view all dates, edit/delete any
    entry, download the Excel workbook, and manage the employee master
    list, the process list, and user login accounts.
  - USERS (email + password, created by the admin): log in and fill the
    tracker form for any employee in the master list, picking a Process
    from the admin-maintained process list. Users can only VIEW entries
    (their own submissions) — no edit, delete, or download rights.
    Note: the login email identifies who is filling the form, NOT whose
    data is being entered — one login can be used to log data for many
    different employees picked from the master list.

Run it with:
    pip install flask openpyxl --break-system-packages
    python app.py

Then open http://127.0.0.1:5000
Admin login:  http://127.0.0.1:5000/admin/login
User login:   http://127.0.0.1:5000/login

Excel file created at: tracker_data/productivity_tracker.xlsx
  - "Users" sheet        -> login accounts (Email, Password, Name)
  - "Master" sheet       -> employee master list (Band, Emp_Id, Emp_Name)
  - "Process_List" sheet -> list of selectable process names
  - date-named sheets    -> one row per tracker entry submitted that day

SECURITY NOTE: Credentials are stored/checked in plain text, which is
fine for a small internal/local tool. If you ever deploy this somewhere
public, swap in hashed passwords and HTTPS, and move credentials to
environment variables (SECRET_KEY / ADMIN_USERNAME / ADMIN_PASSWORD are
already read from os.environ below, with local fallbacks).
"""

import os
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
from functools import wraps
from zoneinfo import ZoneInfo

from flask import (
    Flask, render_template_string, request, redirect, url_for,
    flash, session, send_file, abort
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DATA_FOLDER = "tracker_data"
EXCEL_FILE = os.path.join(DATA_FOLDER, "productivity_tracker.xlsx")

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "Mobius365")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "Mobius@123")

WORK_HOURS_PER_DAY = 8.0
LEAVE_HR = 8.0  # Hr auto-filled when a user checks "On Leave today"

# The server may run in a different timezone than the users (e.g. hosting
# providers default to UTC). All dates/times shown or logged in the app use
# this timezone instead of the server's local time, so "Login: 10:10:29"
# matches what the clock on the user's wall actually says.
APP_TIMEZONE = os.environ.get("APP_TIMEZONE", "Asia/Kolkata")


def now():
    """Current datetime in APP_TIMEZONE (IST by default), not server-local time."""
    return datetime.now(ZoneInfo(APP_TIMEZONE))

# SMTP settings for the 3:00 PM "you haven't filled 8 hrs today" reminder.
# If SMTP_SERVER isn't set, reminder emails are skipped (logged only) so the
# app still runs fine locally without mail configured.
SMTP_SERVER = os.environ.get("SMTP_SERVER")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USERNAME = os.environ.get("SMTP_USERNAME")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")
SMTP_FROM = os.environ.get("SMTP_FROM", SMTP_USERNAME or "noreply@example.com")
REMINDER_HOUR = int(os.environ.get("REMINDER_HOUR", "15"))   # 3:00 PM
REMINDER_MINUTE = int(os.environ.get("REMINDER_MINUTE", "0"))

USERS_SHEET = "Users"
MASTER_SHEET = "Master"
PROCESS_SHEET = "Process_List"
RESERVED_SHEETS = {USERS_SHEET, MASTER_SHEET, PROCESS_SHEET, "Info"}

USER_COLUMNS = ["Email", "Password", "Name"]
MASTER_COLUMNS = ["Band", "Emp_Id", "Emp_Name"]
PROCESS_COLUMNS = ["Process", "Target_Hr", "Target_Pct"]

# (internal key, label shown on the form / table header)
TRACKER_FIELDS = [
    ("Date", "Date"),
    ("Band", "Band"),
    ("Emp_Id", "Emp_Id"),
    ("Emp_Name", "Emp_Name"),
    ("Process", "Process"),
    ("Description", "Description"),
    ("Other", "Other"),
    ("Hr", "Hr"),
    ("Other_Description", "Description"),
    ("Logged_By", "Logged By"),
]
TRACKER_COLUMNS = [k for k, _ in TRACKER_FIELDS]

HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial")
CELL_FONT = Font(name="Arial")


# ---------------------------------------------------------------------------
# Excel helpers — generic
# ---------------------------------------------------------------------------
def style_header(ws, columns, widths=None):
    for col_idx, header in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    if not widths:
        widths = [16] * len(columns)
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def ensure_workbook():
    """Create the workbook + folder + fixed sheets if they don't exist."""
    os.makedirs(DATA_FOLDER, exist_ok=True)
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        wb.active.title = "Info"
        wb.active["A1"] = "LN Risk Map — Productivity Tracker. Data is stored in date-named sheets."
        wb.save(EXCEL_FILE)

    wb = load_workbook(EXCEL_FILE)
    changed = False
    if USERS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(USERS_SHEET)
        ws.append(USER_COLUMNS)
        style_header(ws, USER_COLUMNS, [30, 18, 20])
        changed = True
    if MASTER_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(MASTER_SHEET)
        ws.append(MASTER_COLUMNS)
        style_header(ws, MASTER_COLUMNS, [10, 12, 20])
        changed = True
    if PROCESS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(PROCESS_SHEET)
        ws.append(PROCESS_COLUMNS)
        style_header(ws, PROCESS_COLUMNS, [30, 14, 14])
        changed = True
    else:
        # Migrate older Process_List sheets that predate the Target_Hr /
        # Target_Pct columns by appending the missing headers.
        ws = wb[PROCESS_SHEET]
        existing_headers = [c.value for c in ws[1]]
        for col_idx, col_name in enumerate(PROCESS_COLUMNS, start=1):
            if col_idx > len(existing_headers) or existing_headers[col_idx - 1] != col_name:
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = Alignment(horizontal="center")
                changed = True
    if changed:
        wb.save(EXCEL_FILE)


def get_or_create_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)
    ws.append(TRACKER_COLUMNS)
    style_header(ws, TRACKER_COLUMNS, [12, 8, 10, 16, 16, 24, 12, 8, 24, 22])
    return ws


def read_sheet_rows(sheet_name, columns):
    """Generic reader: returns list of dicts (with _row) for a fixed sheet."""
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = load_workbook(EXCEL_FILE, data_only=True)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row and any(row):
            record = dict(zip(columns, row))
            record["_row"] = row_idx
            rows.append(record)
    return rows


# ---------------------------------------------------------------------------
# Users (login accounts) helpers
# ---------------------------------------------------------------------------
def get_users():
    return read_sheet_rows(USERS_SHEET, USER_COLUMNS)


def find_user_by_email(email):
    for u in get_users():
        if (u.get("Email") or "").strip().lower() == (email or "").strip().lower():
            return u
    return None


def add_user(email, password, name):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[USERS_SHEET]
    ws.append([email, password, name])
    row_idx = ws.max_row
    for col_idx in range(1, len(USER_COLUMNS) + 1):
        ws.cell(row=row_idx, column=col_idx).font = CELL_FONT
    wb.save(EXCEL_FILE)


def update_user(row_idx, email, password, name):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[USERS_SHEET]
    ws.cell(row=row_idx, column=1, value=email).font = CELL_FONT
    ws.cell(row=row_idx, column=2, value=password).font = CELL_FONT
    ws.cell(row=row_idx, column=3, value=name).font = CELL_FONT
    wb.save(EXCEL_FILE)


def delete_user(row_idx):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[USERS_SHEET]
    ws.delete_rows(row_idx, 1)
    wb.save(EXCEL_FILE)


# ---------------------------------------------------------------------------
# Master employee list helpers
# ---------------------------------------------------------------------------
def get_master_list():
    return read_sheet_rows(MASTER_SHEET, MASTER_COLUMNS)


def add_master(data):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[MASTER_SHEET]
    row = [data.get(c, "") for c in MASTER_COLUMNS]
    ws.append(row)
    row_idx = ws.max_row
    for col_idx in range(1, len(MASTER_COLUMNS) + 1):
        ws.cell(row=row_idx, column=col_idx).font = CELL_FONT
    wb.save(EXCEL_FILE)


def update_master(row_idx, data):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[MASTER_SHEET]
    for col_idx, col in enumerate(MASTER_COLUMNS, start=1):
        ws.cell(row=row_idx, column=col_idx, value=data.get(col, "")).font = CELL_FONT
    wb.save(EXCEL_FILE)


def delete_master(row_idx):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[MASTER_SHEET]
    ws.delete_rows(row_idx, 1)
    wb.save(EXCEL_FILE)


# ---------------------------------------------------------------------------
# Process list helpers
# ---------------------------------------------------------------------------
def get_process_list():
    return read_sheet_rows(PROCESS_SHEET, PROCESS_COLUMNS)


def add_process(name, target_hr="", target_pct=""):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[PROCESS_SHEET]
    ws.append([name, target_hr, target_pct])
    row_idx = ws.max_row
    for col_idx in range(1, len(PROCESS_COLUMNS) + 1):
        ws.cell(row=row_idx, column=col_idx).font = CELL_FONT
    wb.save(EXCEL_FILE)


def update_process(row_idx, name, target_hr="", target_pct=""):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[PROCESS_SHEET]
    ws.cell(row=row_idx, column=1, value=name).font = CELL_FONT
    ws.cell(row=row_idx, column=2, value=target_hr).font = CELL_FONT
    ws.cell(row=row_idx, column=3, value=target_pct).font = CELL_FONT
    wb.save(EXCEL_FILE)


def delete_process(row_idx):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb[PROCESS_SHEET]
    ws.delete_rows(row_idx, 1)
    wb.save(EXCEL_FILE)


# ---------------------------------------------------------------------------
# Tracker entries helpers
# ---------------------------------------------------------------------------
def save_entry(data: dict):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    today_str = now().strftime("%Y-%m-%d")
    ws = get_or_create_sheet(wb, today_str)

    row = [data.get(col, "") for col in TRACKER_COLUMNS]
    ws.append(row)

    new_row_idx = ws.max_row
    for col_idx in range(1, len(TRACKER_COLUMNS) + 1):
        ws.cell(row=new_row_idx, column=col_idx).font = CELL_FONT

    wb.save(EXCEL_FILE)


def get_records_for_date(date_str):
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = load_workbook(EXCEL_FILE, data_only=True)
    if date_str not in wb.sheetnames:
        return []
    ws = wb[date_str]
    records = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row and any(row):
            record = dict(zip(TRACKER_COLUMNS, row))
            record["_row"] = row_idx
            records.append(record)
    return records


def get_today_records():
    return get_records_for_date(now().strftime("%Y-%m-%d"))


def list_available_dates():
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = load_workbook(EXCEL_FILE, data_only=True)
    dates = [name for name in wb.sheetnames if name not in RESERVED_SHEETS]
    dates.sort(reverse=True)
    return dates


def update_record(date_str, row_idx, data: dict):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    if date_str not in wb.sheetnames:
        return False
    ws = wb[date_str]
    for col_idx, col in enumerate(TRACKER_COLUMNS, start=1):
        ws.cell(row=row_idx, column=col_idx, value=data.get(col, "")).font = CELL_FONT
    wb.save(EXCEL_FILE)
    return True


def delete_record(date_str, row_idx):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    if date_str not in wb.sheetnames:
        return False
    ws = wb[date_str]
    ws.delete_rows(row_idx, 1)
    wb.save(EXCEL_FILE)
    return True


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Hours worked / pending helpers
# ---------------------------------------------------------------------------
def parse_hr_total(hr_value):
    """Sum a possibly '; '-joined Hr string (from multi-process rows) into a float."""
    if not hr_value:
        return 0.0
    total = 0.0
    for part in str(hr_value).split(";"):
        part = part.strip()
        if not part:
            continue
        try:
            total += float(part)
        except ValueError:
            pass
    return total


def total_hr_for_records(records):
    return sum(parse_hr_total(r.get("Hr")) for r in records)


def pending_hr_for_records(records):
    return max(0.0, WORK_HOURS_PER_DAY - total_hr_for_records(records))


def process_breakdown(record, target_lookup=None):
    """Split a record's ';'-joined Process and Hr strings into a list of
    {process, hr, pct, target_hr, target_pct} dicts, where pct is that
    process's share of the entry's own total hours, and target_hr/target_pct
    come from the admin-maintained Process List (target_lookup keyed by
    process name). Admin-only display."""
    target_lookup = target_lookup or {}
    procs = [p.strip() for p in str(record.get("Process") or "").split(";")]
    hrs = [h.strip() for h in str(record.get("Hr") or "").split(";")]
    items = []
    parsed_hrs = []
    for h in hrs:
        try:
            parsed_hrs.append(float(h)) if h else parsed_hrs.append(0.0)
        except ValueError:
            parsed_hrs.append(0.0)
    total = sum(parsed_hrs)
    for i, proc in enumerate(procs):
        if not proc:
            continue
        hr = parsed_hrs[i] if i < len(parsed_hrs) else 0.0
        pct = round((hr / total) * 100) if total > 0 else 0
        target = target_lookup.get(proc, {})
        items.append({
            "process": proc, "hr": hr, "pct": pct,
            "target_hr": target.get("Target_Hr") or "",
            "target_pct": target.get("Target_Pct") or "",
        })
    return items


def entry_counts_by_user(records):
    """Per Logged_By user: number of entries submitted + total Hr, for a
    given date's records. Sorted by count, highest first."""
    counts = {}
    for r in records:
        who = r.get("Logged_By") or "(unknown)"
        if who not in counts:
            counts[who] = {"Logged_By": who, "count": 0, "hr": 0.0}
        counts[who]["count"] += 1
        counts[who]["hr"] += parse_hr_total(r.get("Hr"))
    return sorted(counts.values(), key=lambda x: x["count"], reverse=True)


# ---------------------------------------------------------------------------
# Email reminder (3:00 PM daily, for anyone under 8 hrs today)
# ---------------------------------------------------------------------------
def send_email(to_email, subject, body):
    if not SMTP_SERVER:
        print(f"[reminder] SMTP not configured — would have emailed {to_email}: {subject}")
        return
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = SMTP_FROM
    msg["To"] = to_email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            if SMTP_USERNAME and SMTP_PASSWORD:
                server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.sendmail(SMTP_FROM, [to_email], msg.as_string())
    except Exception as exc:  # pragma: no cover - best-effort reminder
        print(f"[reminder] Failed to email {to_email}: {exc}")


def send_daily_pending_hr_reminders():
    """For every login account, if today's logged Hr total is below the
    8 hr/day target (or nothing was filled at all), email that account's
    login address a reminder."""
    ensure_workbook()
    today = now().strftime("%Y-%m-%d")
    today_records = get_records_for_date(today)
    for user in get_users():
        email = user.get("Email")
        if not email:
            continue
        my_records = [r for r in today_records if r.get("Logged_By") == email]
        pending = pending_hr_for_records(my_records)
        if pending > 0:
            worked = total_hr_for_records(my_records)
            body = (
                f"Hi {user.get('Name') or ''},\n\n"
                f"As of 3:00 PM, you have logged {worked:g} of the {WORK_HOURS_PER_DAY:g} "
                f"working hours required for {today} in the Productivity Tracker.\n"
                f"Pending: {pending:g} hr.\n\n"
                f"Please log in and fill in the remaining entries.\n"
            )
            send_email(email, f"Productivity Tracker — {pending:g} hr pending for {today}", body)


def _start_reminder_scheduler():
    """Best-effort background scheduler for the 3:00 PM reminder. Uses
    APScheduler if installed; silently no-ops otherwise so the app still
    runs without the extra dependency."""
    try:
        from apscheduler.schedulers.background import BackgroundScheduler
    except ImportError:
        print("[reminder] apscheduler not installed — daily 3:00 PM reminder disabled. "
              "Run `pip install apscheduler` to enable it.")
        return
    scheduler = BackgroundScheduler(daemon=True, timezone=ZoneInfo(APP_TIMEZONE))
    scheduler.add_job(
        send_daily_pending_hr_reminders,
        "cron", hour=REMINDER_HOUR, minute=REMINDER_MINUTE,
    )
    scheduler.start()


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------
def admin_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login"))
        return view_func(*args, **kwargs)
    return wrapper


def user_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not session.get("user_email"):
            return redirect(url_for("user_login"))
        return view_func(*args, **kwargs)
    return wrapper


# ---------------------------------------------------------------------------
# HTML templates
# ---------------------------------------------------------------------------
BASE_STYLE = """
<style>
    body { font-family: Arial, sans-serif; max-width: 1000px; margin: 30px auto; background: #f5f6fa; color: #222; }
    h1 { color: #2c3e50; }
    h2 { color: #2c3e50; font-size: 18px; }
    .card { background: #fff; padding: 24px; border-radius: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.08); margin-bottom: 24px; }
    label { display: block; margin-top: 12px; font-weight: bold; font-size: 14px; }
    input[type=text], input[type=date], input[type=number], input[type=password], input[type=email], select {
        width: 100%; padding: 8px; margin-top: 4px; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box;
    }
    .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 0 16px; }
    .row3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0 16px; }
    button, .btn { margin-top: 18px; padding: 10px 22px; background: #305496; color: white; border: none;
        border-radius: 6px; font-size: 15px; cursor: pointer; text-decoration: none; display: inline-block; }
    button:hover, .btn:hover { background: #24406f; }
    .btn-danger { background: #b23b3b; }
    .btn-danger:hover { background: #8f2e2e; }
    .btn-small { padding: 5px 12px; font-size: 13px; margin: 0 4px 0 0; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th, td { border: 1px solid #ddd; padding: 8px; font-size: 13px; text-align: left; }
    th { background: #305496; color: white; }
    .flash { padding: 10px 14px; background: #d4edda; color: #155724; border-radius: 6px; margin-bottom: 16px; }
    .flash-error { background: #f8d7da; color: #721c24; }
    .note { font-size: 13px; color: #555; }
    .topbar { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }
    .topbar a { color: #305496; text-decoration: none; font-size: 14px; margin-left: 12px; }
    .tag { display: inline-block; background: #eef1f8; color: #305496; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
    body.login-page { max-width: 380px; margin: 0 auto; min-height: 100vh; display: flex; flex-direction: column;
        justify-content: center; padding: 16px; }
    body.login-page h1 { text-align: center; font-size: 20px; margin: 0 0 14px; }
    body.login-page .card { padding: 16px; margin-bottom: 14px; }
    body.login-page p { text-align: center; margin: 0; }
</style>
"""

FLASHES = """
{% with messages = get_flashed_messages(with_categories=true) %}
  {% if messages %}
    {% for cat, m in messages %}
      <div class="flash {% if cat == 'error' %}flash-error{% endif %}">{{ m }}</div>
    {% endfor %}
  {% endif %}
{% endwith %}
"""

# --- User: login -----------------------------------------------------------
USER_LOGIN_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>User Login</title>""" + BASE_STYLE + """</head>
<body class="login-page">
    <h1>👤 User Login</h1>
    """ + FLASHES + """
    <div class="card">
        <p class="note">Log in with the email &amp; password given to you by your admin.
        You may use this login to enter data for <b>any</b> employee in the list — it does not have to be your own record.</p>
        <form method="POST">
            <label>Email</label>
            <input type="email" name="email" required autofocus>
            <label>Password</label>
            <input type="password" name="password" required>
            <button type="submit">Login</button>
        </form>
    </div>
    <p><a href="{{ url_for('admin_login') }}">🔐 Admin login instead</a></p>
</body>
</html>
"""

# --- User: form + own submissions ------------------------------------------
USER_HOME_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Productivity Tracker</title>""" + BASE_STYLE + """</head>
<body>
    <div class="topbar">
        <h1>📊 LN Risk Map — Productivity Tracker</h1>
        <div>
            <span class="tag">Logged in as {{ user_name }} ({{ user_email }})</span>
            <a href="{{ url_for('user_logout') }}">Log out</a>
        </div>
    </div>
    """ + FLASHES + """

    <div class="card">
        <h2>New Entry</h2>
        <form method="POST" action="{{ url_for('submit') }}">
            <div class="row2">
                <div>
                    <label>Date</label>
                    <input type="date" name="Date" value="{{ today }}" required>
                </div>
                <div>
                    <label>Select Employee (from master list)</label>
                    <select id="empSelect" onchange="fillEmp()">
                        <option value="">-- choose employee --</option>
                        {% for m in master %}
                        <option value="{{ loop.index0 }}"
                            data-band="{{ m.Band or '' }}"
                            data-empid="{{ m.Emp_Id or '' }}"
                            data-empname="{{ m.Emp_Name or '' }}">
                            {{ m.Emp_Id }} — {{ m.Emp_Name }}
                        </option>
                        {% endfor %}
                    </select>
                </div>
            </div>

            <div class="row3">
                <div><label>Band</label><input type="text" name="Band" id="Band"></div>
                <div><label>Emp_Id</label><input type="text" name="Emp_Id" id="Emp_Id"></div>
                <div><label>Emp_Name</label><input type="text" name="Emp_Name" id="Emp_Name"></div>
            </div>

            <div class="card" style="background:#fff8e1;border:1px solid #f0d98c;padding:10px 14px;margin:10px 0;">
                <label style="display:flex;align-items:center;gap:8px;cursor:pointer;">
                    <input type="checkbox" id="leaveCheck" onchange="toggleLeave()">
                    On Leave today (auto-fills {{ '%g'|format(leave_hr) }} hr)
                </label>
            </div>

            <div id="processSection">
            <label>Process</label>
            <div id="processRows">
                <div class="row3 process-row">
                    <div>
                        <select class="proc-select">
                            <option value="">-- select process --</option>
                            {% for p in processes %}
                            <option value="{{ p.Process }}">{{ p.Process }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div><input type="text" class="proc-desc" placeholder="Description"></div>
                    <div><input type="number" step="0.25" class="proc-hr" placeholder="Hr"></div>
                </div>
            </div>
            <button type="button" class="btn btn-small" onclick="addProcessRow()">+ Add</button>
            </div>
            <input type="hidden" name="Process" id="Process_hidden">
            <input type="hidden" name="Description" id="Description_hidden">
            <input type="hidden" name="Hr" id="Hr_hidden">

            <div class="row2">
                <div><label>Other</label><input type="text" name="Other"></div>
                <div><label>Description</label><input type="text" name="Other_Description"></div>
            </div>

            <button type="submit" onclick="return collectProcessRows()">Save Entry</button>
        </form>
    </div>

    <div class="card" style="{% if pending_hr > 0 %}border:2px solid #c0392b;{% else %}border:2px solid #2e8b57;{% endif %}">
        <h2>Today's Hours — {{ today }}</h2>
        <div class="row3">
            <div><span class="tag">Target: {{ '%g'|format(target_hr) }} hr/day</span></div>
            <div><span class="tag">Logged: {{ '%g'|format(total_hr) }} hr</span></div>
            <div><span class="tag" style="{% if pending_hr > 0 %}background:#f8d7da;color:#721c24;{% else %}background:#d4edda;color:#155724;{% endif %}">
                Pending: {{ '%g'|format(pending_hr) }} hr
            </span></div>
        </div>
        {% if pending_hr > 0 %}
        <p style="color:#c0392b;font-weight:bold;margin-top:14px;margin-bottom:0;">
            ⚠ Productivity not complete — {{ '%g'|format(pending_hr) }} hr still pending out of {{ '%g'|format(target_hr) }} hr/day.
        </p>
        {% else %}
        <p style="color:#2e8b57;font-weight:bold;margin-top:14px;margin-bottom:0;">
            ✅ {{ '%g'|format(target_hr) }} hr/day completed.
        </p>
        {% endif %}
    </div>

    <div class="card">
        <h2>Your Submissions — {{ today }}</h2>
        <p class="note">View only — contact admin for edits or corrections.</p>
        {% if records %}
        <table>
            <tr>{% for _, label in fields %}<th>{{ label }}</th>{% endfor %}</tr>
            {% for r in records %}
            <tr>{% for key, _ in fields %}<td>{{ r[key] }}</td>{% endfor %}</tr>
            {% endfor %}
        </table>
        {% else %}
        <p>No entries submitted yet today.</p>
        {% endif %}
    </div>

<script>
function fillEmp() {
    var sel = document.getElementById('empSelect');
    var opt = sel.options[sel.selectedIndex];
    if (!opt || !opt.value) return;
    document.getElementById('Band').value = opt.getAttribute('data-band');
    document.getElementById('Emp_Id').value = opt.getAttribute('data-empid');
    document.getElementById('Emp_Name').value = opt.getAttribute('data-empname');
}

function toggleLeave() {
    var checked = document.getElementById('leaveCheck').checked;
    document.getElementById('processSection').style.display = checked ? 'none' : '';
}

function addProcessRow() {
    var container = document.getElementById('processRows');
    var row = container.children[0].cloneNode(true);
    row.querySelector('.proc-select').selectedIndex = 0;
    row.querySelector('.proc-desc').value = '';
    row.querySelector('.proc-hr').value = '';
    container.appendChild(row);
}

function collectProcessRows() {
    if (document.getElementById('leaveCheck').checked) {
        document.getElementById('Process_hidden').value = 'Leave';
        document.getElementById('Description_hidden').value = 'Leave';
        document.getElementById('Hr_hidden').value = '{{ leave_hr }}';
        return true;
    }
    var selects = document.querySelectorAll('#processRows .proc-select');
    var descs = document.querySelectorAll('#processRows .proc-desc');
    var hrs = document.querySelectorAll('#processRows .proc-hr');
    var procs = [], descVals = [], hrVals = [];
    for (var i = 0; i < selects.length; i++) {
        var p = selects[i].value.trim();
        var d = descs[i].value.trim();
        var h = hrs[i].value.trim();
        if (p || d || h) {
            procs.push(p);
            descVals.push(d);
            hrVals.push(h);
        }
    }
    document.getElementById('Process_hidden').value = procs.join('; ');
    document.getElementById('Description_hidden').value = descVals.join('; ');
    document.getElementById('Hr_hidden').value = hrVals.join('; ');
    if (procs.length === 0) {
        alert('Please select at least one process.');
        return false;
    }
    return true;
}
</script>
</body>
</html>
"""

# --- Admin: login ------------------------------------------------------------
ADMIN_LOGIN_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Admin Login</title>""" + BASE_STYLE + """</head>
<body class="login-page">
    <h1>🔐 Admin Login</h1>
    """ + FLASHES + """
    <div class="card">
        <form method="POST">
            <label>Username</label>
            <input type="text" name="username" required autofocus>
            <label>Password</label>
            <input type="password" name="password" required>
            <button type="submit">Login</button>
        </form>
    </div>
    <p><a href="{{ url_for('user_login') }}">👤 User login instead</a></p>
</body>
</html>
"""

# --- Admin: main panel (records) --------------------------------------------
ADMIN_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Admin Panel</title>""" + BASE_STYLE + """</head>
<body>
    <div class="topbar">
        <h1>🛠 Admin Panel</h1>
        <div>
            <a href="{{ url_for('admin_master') }}">Employee Master List</a>
            <a href="{{ url_for('admin_process') }}">Process List</a>
            <a href="{{ url_for('admin_users') }}">Login Accounts</a>
            <a href="{{ url_for('admin_logout') }}">Log out</a>
        </div>
    </div>
    """ + FLASHES + """

    <div class="card">
        <form method="GET" action="{{ url_for('admin_panel') }}">
            <label>Select date</label>
            <select name="date" onchange="this.form.submit()">
                {% for d in dates %}
                <option value="{{ d }}" {% if d == selected_date %}selected{% endif %}>{{ d }}</option>
                {% endfor %}
            </select>
        </form>
        <a class="btn" href="{{ url_for('download_excel') }}">⬇ Download Full Excel File</a>
    </div>

    <div class="card">
        <h2>Productivity Count — {{ selected_date }}</h2>
        <p class="note">Number of tracker entries each user submitted on this date.</p>
        {% if user_counts %}
        <table>
            <tr><th>Logged By</th><th>Entries</th><th>Total Hr</th></tr>
            {% for u in user_counts %}
            <tr>
                <td>{{ u.Logged_By }}</td>
                <td>{{ u.count }}</td>
                <td>{{ '%g'|format(u.hr) }}</td>
            </tr>
            {% endfor %}
        </table>
        {% else %}
        <p>No entries for this date.</p>
        {% endif %}
    </div>

    <div class="card">
        <h2>Records — {{ selected_date }}</h2>
        {% if records %}
        <table>
            <tr>
                {% for _, label in fields %}<th>{{ label }}</th>{% endfor %}
                <th>Process %</th>
                <th>Actions</th>
            </tr>
            {% for r in records %}
            <tr>
                {% for key, _ in fields %}<td>{{ r[key] }}</td>{% endfor %}
                <td>
                    {% for b in r._breakdown %}
                    {{ b.process }}: {{ b.hr|string }}hr ({{ b.pct }}%){% if b.target_hr or b.target_pct %} <span class="note">/ target {{ b.target_hr or '-' }}hr ({{ b.target_pct or '-' }}%)</span>{% endif %}{% if not loop.last %}<br>{% endif %}
                    {% endfor %}
                </td>
                <td>
                    <a class="btn btn-small" href="{{ url_for('admin_edit', date=selected_date, row=r['_row']) }}">Edit</a>
                    <form style="display:inline" method="POST" action="{{ url_for('admin_delete', date=selected_date, row=r['_row']) }}"
                          onsubmit="return confirm('Delete this record?');">
                        <button class="btn btn-small btn-danger" type="submit">Delete</button>
                    </form>
                </td>
            </tr>
            {% endfor %}
        </table>
        {% else %}
        <p>No records for this date.</p>
        {% endif %}
    </div>
</body>
</html>
"""

EDIT_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Edit Record</title>""" + BASE_STYLE + """</head>
<body>
    <h1>✏️ Edit Record — {{ date }}</h1>
    <div class="card">
        <form method="POST">
            {% for key, label in fields %}
            <label>{{ label }}</label>
            <input type="text" name="{{ key }}" value="{{ record[key] or '' }}">
            {% endfor %}
            <button type="submit">Save Changes</button>
        </form>
    </div>
    <p><a href="{{ url_for('admin_panel', date=date) }}">← Back to admin panel</a></p>
</body>
</html>
"""

# --- Admin: employee master list --------------------------------------------
MASTER_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Employee Master List</title>""" + BASE_STYLE + """</head>
<body>
    <div class="topbar">
        <h1>📋 Employee Master List</h1>
        <div><a href="{{ url_for('admin_panel') }}">← Admin panel</a></div>
    </div>
    """ + FLASHES + """

    <div class="card">
        <h2>Add Employee</h2>
        <form method="POST" action="{{ url_for('admin_master_add') }}">
            <div class="row3">
                <div><label>Band</label><input type="text" name="Band" required></div>
                <div><label>Emp_Id</label><input type="text" name="Emp_Id" required></div>
                <div><label>Emp_Name</label><input type="text" name="Emp_Name" required></div>
            </div>
            <button type="submit">Add to List</button>
        </form>
    </div>

    <div class="card">
        <h2>Current List</h2>
        {% if master %}
        <table>
            <tr>{% for c in columns %}<th>{{ c }}</th>{% endfor %}<th>Actions</th></tr>
            {% for m in master %}
            <tr>
                {% for c in columns %}<td>{{ m[c] }}</td>{% endfor %}
                <td>
                    <a class="btn btn-small" href="{{ url_for('admin_master_edit', row=m['_row']) }}">Edit</a>
                    <form style="display:inline" method="POST" action="{{ url_for('admin_master_delete', row=m['_row']) }}"
                          onsubmit="return confirm('Remove this employee from the list?');">
                        <button class="btn btn-small btn-danger" type="submit">Delete</button>
                    </form>
                </td>
            </tr>
            {% endfor %}
        </table>
        {% else %}
        <p>No employees added yet.</p>
        {% endif %}
    </div>
</body>
</html>
"""

MASTER_EDIT_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Edit Employee</title>""" + BASE_STYLE + """</head>
<body>
    <h1>✏️ Edit Employee</h1>
    <div class="card">
        <form method="POST">
            {% for c in columns %}
            <label>{{ c }}</label>
            <input type="text" name="{{ c }}" value="{{ record[c] or '' }}">
            {% endfor %}
            <button type="submit">Save Changes</button>
        </form>
    </div>
    <p><a href="{{ url_for('admin_master') }}">← Back to master list</a></p>
</body>
</html>
"""

# --- Admin: process list -----------------------------------------------------
PROCESS_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Process List</title>""" + BASE_STYLE + """</head>
<body>
    <div class="topbar">
        <h1>🧩 Process List</h1>
        <div><a href="{{ url_for('admin_panel') }}">← Admin panel</a></div>
    </div>
    """ + FLASHES + """

    <div class="card">
        <h2>Add Process</h2>
        <form method="POST" action="{{ url_for('admin_process_add') }}">
            <div class="row3">
                <div><label>Process name</label><input type="text" name="Process" required></div>
                <div><label>Target Hr</label><input type="number" step="0.25" name="Target_Hr"></div>
                <div><label>Target %</label><input type="number" step="1" name="Target_Pct"></div>
            </div>
            <button type="submit">Add Process</button>
        </form>
        <p class="note">This is the dropdown list users choose from when filling out an entry. Target Hr / Target % are used to compare against actual logged hours in the admin records view.</p>
    </div>

    <div class="card">
        <h2>Current Processes</h2>
        {% if processes %}
        <table>
            <tr><th>Process</th><th>Target Hr</th><th>Target %</th><th>Actions</th></tr>
            {% for p in processes %}
            <tr>
                <td>{{ p.Process }}</td>
                <td>{{ p.Target_Hr or '-' }}</td>
                <td>{{ p.Target_Pct or '-' }}</td>
                <td>
                    <a class="btn btn-small" href="{{ url_for('admin_process_edit', row=p['_row']) }}">Edit</a>
                    <form style="display:inline" method="POST" action="{{ url_for('admin_process_delete', row=p['_row']) }}"
                          onsubmit="return confirm('Remove this process from the list?');">
                        <button class="btn btn-small btn-danger" type="submit">Delete</button>
                    </form>
                </td>
            </tr>
            {% endfor %}
        </table>
        {% else %}
        <p>No processes added yet.</p>
        {% endif %}
    </div>
</body>
</html>
"""

PROCESS_EDIT_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Edit Process</title>""" + BASE_STYLE + """</head>
<body>
    <h1>✏️ Edit Process</h1>
    <div class="card">
        <form method="POST">
            <label>Process name</label>
            <input type="text" name="Process" value="{{ record.Process or '' }}" required>
            <div class="row2">
                <div><label>Target Hr</label><input type="number" step="0.25" name="Target_Hr" value="{{ record.Target_Hr or '' }}"></div>
                <div><label>Target %</label><input type="number" step="1" name="Target_Pct" value="{{ record.Target_Pct or '' }}"></div>
            </div>
            <button type="submit">Save Changes</button>
        </form>
    </div>
    <p><a href="{{ url_for('admin_process') }}">← Back to process list</a></p>
</body>
</html>
"""

# --- Admin: login accounts ---------------------------------------------------
USERS_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Login Accounts</title>""" + BASE_STYLE + """</head>
<body>
    <div class="topbar">
        <h1>👥 User Login Accounts</h1>
        <div><a href="{{ url_for('admin_panel') }}">← Admin panel</a></div>
    </div>
    """ + FLASHES + """

    <div class="card">
        <h2>Add Login Account</h2>
        <form method="POST" action="{{ url_for('admin_users_add') }}">
            <div class="row3">
                <div><label>Email</label><input type="email" name="Email" required></div>
                <div><label>Password</label><input type="text" name="Password" required></div>
                <div><label>Name</label><input type="text" name="Name" required></div>
            </div>
            <p class="note">"Name" identifies the person this login belongs to. It does not
            restrict which employee's data they can fill in — that's chosen from the master list.</p>
            <button type="submit">Add Account</button>
        </form>
    </div>

    <div class="card">
        <h2>Current Accounts</h2>
        {% if users %}
        <table>
            <tr><th>Email</th><th>Password</th><th>Name</th><th>Actions</th></tr>
            {% for u in users %}
            <tr>
                <td>{{ u.Email }}</td><td>{{ u.Password }}</td><td>{{ u.Name }}</td>
                <td>
                    <a class="btn btn-small" href="{{ url_for('admin_users_edit', row=u['_row']) }}">Edit</a>
                    <form style="display:inline" method="POST" action="{{ url_for('admin_users_delete', row=u['_row']) }}"
                          onsubmit="return confirm('Remove this login account?');">
                        <button class="btn btn-small btn-danger" type="submit">Delete</button>
                    </form>
                </td>
            </tr>
            {% endfor %}
        </table>
        {% else %}
        <p>No user accounts yet.</p>
        {% endif %}
    </div>
</body>
</html>
"""

USER_EDIT_PAGE = """
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Edit Account</title>""" + BASE_STYLE + """</head>
<body>
    <h1>✏️ Edit Login Account</h1>
    <div class="card">
        <form method="POST">
            <label>Email</label>
            <input type="email" name="Email" value="{{ record.Email or '' }}" required>
            <label>Password</label>
            <input type="text" name="Password" value="{{ record.Password or '' }}" required>
            <label>Name</label>
            <input type="text" name="Name" value="{{ record.Name or '' }}" required>
            <button type="submit">Save Changes</button>
        </form>
    </div>
    <p><a href="{{ url_for('admin_users') }}">← Back to accounts</a></p>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# User-facing routes
# ---------------------------------------------------------------------------
@app.route("/")
def home():
    if session.get("user_email"):
        return redirect(url_for("user_home"))
    return redirect(url_for("user_login"))


@app.route("/login", methods=["GET", "POST"])
def user_login():
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        user = find_user_by_email(email)
        if user and str(user.get("Password")) == password:
            session["user_email"] = user["Email"]
            session["user_name"] = user.get("Name") or user["Email"]
            flash("Logged in successfully.")
            return redirect(url_for("user_home"))
        flash("Invalid email or password.", "error")
    return render_template_string(USER_LOGIN_PAGE)


@app.route("/logout")
def user_logout():
    session.pop("user_email", None)
    session.pop("user_name", None)
    return redirect(url_for("user_login"))


@app.route("/tracker")
@user_required
def user_home():
    ensure_workbook()
    master = get_master_list()
    processes = get_process_list()
    records = [r for r in get_today_records() if r.get("Logged_By") == session["user_email"]]
    today = now().strftime("%Y-%m-%d")
    total_hr = total_hr_for_records(records)
    pending_hr = pending_hr_for_records(records)
    return render_template_string(
        USER_HOME_PAGE, master=master, processes=processes, records=records, fields=TRACKER_FIELDS,
        today=today, user_name=session.get("user_name"), user_email=session.get("user_email"),
        total_hr=total_hr, pending_hr=pending_hr, target_hr=WORK_HOURS_PER_DAY, leave_hr=LEAVE_HR
    )


@app.route("/submit", methods=["POST"])
@user_required
def submit():
    data = {key: request.form.get(key, "").strip() for key, _ in TRACKER_FIELDS if key != "Logged_By"}
    data["Logged_By"] = session["user_email"]
    if not data.get("Date"):
        data["Date"] = now().strftime("%Y-%m-%d")
    save_entry(data)
    flash(f"Saved entry for {data.get('Emp_Name') or 'employee'} ({data.get('Emp_Id')}).")
    return redirect(url_for("user_home"))


# ---------------------------------------------------------------------------
# Admin: auth
# ---------------------------------------------------------------------------
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["is_admin"] = True
            flash("Logged in successfully.")
            return redirect(url_for("admin_panel"))
        flash("Invalid username or password.", "error")
    return render_template_string(ADMIN_LOGIN_PAGE)


@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("admin_login"))


# ---------------------------------------------------------------------------
# Admin: tracker records
# ---------------------------------------------------------------------------
@app.route("/admin")
@admin_required
def admin_panel():
    ensure_workbook()
    dates = list_available_dates()
    today = now().strftime("%Y-%m-%d")
    selected_date = request.args.get("date") or (dates[0] if dates else today)
    records = get_records_for_date(selected_date)
    target_lookup = {p["Process"]: p for p in get_process_list() if p.get("Process")}
    for r in records:
        r["_breakdown"] = process_breakdown(r, target_lookup)
    user_counts = entry_counts_by_user(records)
    return render_template_string(
        ADMIN_PAGE, dates=dates, selected_date=selected_date,
        records=records, fields=TRACKER_FIELDS, user_counts=user_counts
    )


@app.route("/admin/edit/<date>/<int:row>", methods=["GET", "POST"])
@admin_required
def admin_edit(date, row):
    if request.method == "POST":
        data = {key: request.form.get(key, "").strip() for key, _ in TRACKER_FIELDS}
        if update_record(date, row, data):
            flash("Record updated.")
        else:
            flash("Could not update record — sheet not found.", "error")
        return redirect(url_for("admin_panel", date=date))

    records = get_records_for_date(date)
    record = next((r for r in records if r["_row"] == row), None)
    if record is None:
        abort(404)
    return render_template_string(EDIT_PAGE, date=date, record=record, fields=TRACKER_FIELDS)


@app.route("/admin/delete/<date>/<int:row>", methods=["POST"])
@admin_required
def admin_delete(date, row):
    if delete_record(date, row):
        flash("Record deleted.")
    else:
        flash("Could not delete record — sheet not found.", "error")
    return redirect(url_for("admin_panel", date=date))


@app.route("/admin/download")
@admin_required
def download_excel():
    ensure_workbook()
    return send_file(EXCEL_FILE, as_attachment=True, download_name="productivity_tracker.xlsx")


# ---------------------------------------------------------------------------
# Admin: employee master list
# ---------------------------------------------------------------------------
@app.route("/admin/master")
@admin_required
def admin_master():
    ensure_workbook()
    return render_template_string(MASTER_PAGE, master=get_master_list(), columns=MASTER_COLUMNS)


@app.route("/admin/master/add", methods=["POST"])
@admin_required
def admin_master_add():
    data = {c: request.form.get(c, "").strip() for c in MASTER_COLUMNS}
    add_master(data)
    flash(f"Added {data.get('Emp_Name')} to the master list.")
    return redirect(url_for("admin_master"))


@app.route("/admin/master/edit/<int:row>", methods=["GET", "POST"])
@admin_required
def admin_master_edit(row):
    if request.method == "POST":
        data = {c: request.form.get(c, "").strip() for c in MASTER_COLUMNS}
        update_master(row, data)
        flash("Employee updated.")
        return redirect(url_for("admin_master"))
    record = next((m for m in get_master_list() if m["_row"] == row), None)
    if record is None:
        abort(404)
    return render_template_string(MASTER_EDIT_PAGE, record=record, columns=MASTER_COLUMNS)


@app.route("/admin/master/delete/<int:row>", methods=["POST"])
@admin_required
def admin_master_delete(row):
    delete_master(row)
    flash("Employee removed from master list.")
    return redirect(url_for("admin_master"))


# ---------------------------------------------------------------------------
# Admin: process list
# ---------------------------------------------------------------------------
@app.route("/admin/process")
@admin_required
def admin_process():
    ensure_workbook()
    return render_template_string(PROCESS_PAGE, processes=get_process_list())


@app.route("/admin/process/add", methods=["POST"])
@admin_required
def admin_process_add():
    name = request.form.get("Process", "").strip()
    target_hr = request.form.get("Target_Hr", "").strip()
    target_pct = request.form.get("Target_Pct", "").strip()
    if name:
        add_process(name, target_hr, target_pct)
        flash(f"Added process '{name}'.")
    return redirect(url_for("admin_process"))


@app.route("/admin/process/edit/<int:row>", methods=["GET", "POST"])
@admin_required
def admin_process_edit(row):
    if request.method == "POST":
        name = request.form.get("Process", "").strip()
        target_hr = request.form.get("Target_Hr", "").strip()
        target_pct = request.form.get("Target_Pct", "").strip()
        update_process(row, name, target_hr, target_pct)
        flash("Process updated.")
        return redirect(url_for("admin_process"))
    record = next((p for p in get_process_list() if p["_row"] == row), None)
    if record is None:
        abort(404)
    return render_template_string(PROCESS_EDIT_PAGE, record=record)


@app.route("/admin/process/delete/<int:row>", methods=["POST"])
@admin_required
def admin_process_delete(row):
    delete_process(row)
    flash("Process removed.")
    return redirect(url_for("admin_process"))


# ---------------------------------------------------------------------------
# Admin: user login accounts
# ---------------------------------------------------------------------------
@app.route("/admin/users")
@admin_required
def admin_users():
    ensure_workbook()
    return render_template_string(USERS_PAGE, users=get_users())


@app.route("/admin/users/add", methods=["POST"])
@admin_required
def admin_users_add():
    email = request.form.get("Email", "").strip()
    password = request.form.get("Password", "").strip()
    name = request.form.get("Name", "").strip()
    if find_user_by_email(email):
        flash("An account with that email already exists.", "error")
    else:
        add_user(email, password, name)
        flash(f"Created login for {name} ({email}).")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/edit/<int:row>", methods=["GET", "POST"])
@admin_required
def admin_users_edit(row):
    if request.method == "POST":
        email = request.form.get("Email", "").strip()
        password = request.form.get("Password", "").strip()
        name = request.form.get("Name", "").strip()
        update_user(row, email, password, name)
        flash("Account updated.")
        return redirect(url_for("admin_users"))
    record = next((u for u in get_users() if u["_row"] == row), None)
    if record is None:
        abort(404)
    return render_template_string(USER_EDIT_PAGE, record=record)


@app.route("/admin/users/delete/<int:row>", methods=["POST"])
@admin_required
def admin_users_delete(row):
    delete_user(row)
    flash("Account removed.")
    return redirect(url_for("admin_users"))


# ---------------------------------------------------------------------------
ensure_workbook()
_start_reminder_scheduler()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
