from flask import Flask, render_template_string, request, redirect, url_for, send_file, flash, jsonify
import sqlite3
import traceback
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import re
import os
import json
import calendar
from werkzeug.security import generate_password_hash, check_password_hash
import zipfile
import io
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
load_dotenv()

app = Flask(__name__)
import secrets as _secrets
_secret_key_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.secret_key')
if os.path.exists(_secret_key_file):
    with open(_secret_key_file, 'r') as f:
        app.secret_key = f.read().strip()
else:
    app.secret_key = _secrets.token_hex(32)
    with open(_secret_key_file, 'w') as f:
        f.write(app.secret_key)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "onetrack.db")
EXCEL_FILE = os.path.join(BASE_DIR, "OneTrack_Generated.xlsx")
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_SECURE'] = bool(os.environ.get('PORT'))

SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_EMAIL = os.environ.get("SMTP_EMAIL", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")

def send_email(to, subject, body):
    if not SMTP_EMAIL or not SMTP_PASSWORD:
        return
    try:
        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = SMTP_EMAIL
        msg["To"] = to
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.send_message(msg)
    except Exception as e:
        pass

# ============== CONSTANTS ==============
STAFF_ORDER_DEFAULT = ["AMIN","SAIDUL","ANIK","SIRAZUL","ALIFF","NORAINI","DIMAS","AIREL","IBRAHIM","HADIF"]

def get_staff_order():
    """Dynamic staff order from DB, falling back to default."""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT staff_id FROM team_members ORDER BY id")
    db_staff = [row[0] for row in c.fetchall()]
    conn.close()
    return db_staff if db_staff else STAFF_ORDER_DEFAULT

def get_staff_for_period(daily):
    """Return staff who have data in the period, current members first, then extras."""
    current = get_staff_order()
    seen = set()
    extras = []
    for day_data in daily.values():
        for name in day_data:
            if name not in seen:
                seen.add(name)
                if name not in current:
                    extras.append(name)
    return current + extras

TASK_KEYS = ["lubes", "combo", "pastry", "shellapp"]
TASK_LABELS = {"lubes": "LUBES", "combo": "COMBO", "pastry": "PASTRY", "shellapp": "SHELL APP"}
TASK_UNITS = {"lubes": "LITRE", "combo": "BAG", "pastry": "PCS", "shellapp": "REG"}
STAFF_TASKS = {}
STAFF_GENDER = {"AMIN":"M","SAIDUL":"M","ANIK":"F","SIRAZUL":"M","ALIFF":"M","NORAINI":"F","DIMAS":"M","AIREL":"M","IBRAHIM":"M","HADIF":"M"}
for s in STAFF_ORDER_DEFAULT:
    STAFF_TASKS[s] = ["lubes", "combo", "shellapp"] if s in ["AMIN","SAIDUL","ANIK","SIRAZUL"] else ["combo", "pastry", "shellapp"]
DAILY_TARGETS_DEFAULT = {}
for s in STAFF_ORDER_DEFAULT:
    if s in ["AMIN","SAIDUL","SIRAZUL"]:
        DAILY_TARGETS_DEFAULT[s] = {"lubes": 5, "combo": 2, "shellapp": 2}
    elif s == "ANIK":
        DAILY_TARGETS_DEFAULT[s] = {"lubes": 4, "combo": 2, "shellapp": 2}
    else:
        DAILY_TARGETS_DEFAULT[s] = {"combo": 2, "pastry": 2, "shellapp": 2}

def get_daily_targets():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT staff_id, targets FROM team_members")
    result = {}
    for row in c.fetchall():
        try:
            result[row["staff_id"]] = json.loads(row["targets"]) if row["targets"] else {}
        except:
            result[row["staff_id"]] = {}
    conn.close()
    for s in get_staff_order():
        if s not in result or not result[s]:
            result[s] = DAILY_TARGETS_DEFAULT.get(s, {})
    return result

def get_monthly_targets(year, month):
    days = calendar.monthrange(year, month)[1]
    dt = get_daily_targets()
    return {s: {tk: dt[s].get(tk, 0) * days for tk in dt[s]} for s in get_staff_order()}

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS responses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT, date TEXT,
        q1 TEXT, q2 TEXT, q3 TEXT, q4 TEXT,
        q5 TEXT, q6 TEXT, q7 TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS team_members (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        staff_id TEXT UNIQUE, name TEXT, gender TEXT,
        picture TEXT, targets TEXT, tasks TEXT,
        created_at TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'user',
        display_name TEXT,
        created_at TEXT,
        email TEXT DEFAULT '',
        reset_code TEXT DEFAULT '',
        reset_expiry TEXT DEFAULT ''
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS tasks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        task_key TEXT UNIQUE NOT NULL,
        label TEXT NOT NULL,
        unit TEXT NOT NULL,
        sort_order INTEGER DEFAULT 0,
        active INTEGER DEFAULT 1,
        created_at TEXT
    )''')
    # Seed users if empty
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        users = [
            ("admin", generate_password_hash("admin123"), "admin", "Administrator"),
            ("staff", generate_password_hash("staff123"), "user", "Staff Member"),
        ]
        for uname, pw_hash, role, display in users:
            c.execute("INSERT INTO users (username,password_hash,role,display_name,created_at) VALUES (?,?,?,?,?)",
                      (uname, pw_hash, role, display, datetime.now().isoformat()))
    # Seed tasks if empty
    c.execute("SELECT COUNT(*) FROM tasks")
    if c.fetchone()[0] == 0:
        default_tasks = [
            ("lubes", "LUBES", "LITRE", 1),
            ("combo", "COMBO", "BAG", 2),
            ("pastry", "PASTRY", "PCS", 3),
            ("shellapp", "SHELL APP", "REG", 4),
        ]
        for tk, label, unit, order in default_tasks:
            c.execute("INSERT INTO tasks (task_key,label,unit,sort_order,active,created_at) VALUES (?,?,?,?,1,?)",
                      (tk, label, unit, order, datetime.now().isoformat()))
    # Seed team_members if empty
    c.execute("SELECT COUNT(*) FROM team_members")
    if c.fetchone()[0] == 0:
        for s in get_staff_order():
            t = json.dumps(DAILY_TARGETS_DEFAULT.get(s, {}))
            tasks_list = STAFF_TASKS.get(s, [])
            c.execute("INSERT INTO team_members (staff_id,name,gender,targets,tasks,created_at) VALUES (?,?,?,?,?,?)",
                      (s, s.title(), STAFF_GENDER.get(s,"M"), t, json.dumps(tasks_list), datetime.now().isoformat()))
    else:
        c.execute("SELECT staff_id, targets FROM team_members")
        for sid, tstr in c.fetchall():
            try:
                t = json.loads(tstr) if tstr else {}
            except:
                t = {}
            needs_fix = not t or any(v > 10 for v in t.values() if isinstance(v, (int, float)))
            if needs_fix:
                new_t = DAILY_TARGETS_DEFAULT.get(sid, {})
                c.execute("UPDATE team_members SET targets=? WHERE staff_id=?", (json.dumps(new_t), sid))
    conn.commit()
    conn.close()
    # Add email/reset columns if missing (migration)
    conn2 = sqlite3.connect(DB_FILE)
    c2 = conn2.cursor()
    for col in ['email', 'reset_code', 'reset_expiry']:
        try:
            c2.execute(f"ALTER TABLE users ADD COLUMN {col} TEXT DEFAULT ''")
        except:
            pass
    # Drop password_plain column if it exists
    try:
        c2.execute("PRAGMA table_info(users)")
        columns = [row[1] for row in c2.fetchall()]
        if 'password_plain' in columns:
            c2.execute("ALTER TABLE users DROP COLUMN password_plain")
    except:
        pass
    conn2.commit()
    conn2.close()
    # Migrate old SHA-256 passwords to werkzeug password hashing
    _known_pws = {"admin": "admin123", "staff": "staff123", "testadmin": "test1234"}
    conn3 = sqlite3.connect(DB_FILE)
    c3 = conn3.cursor()
    c3.execute("SELECT id, username, password_hash FROM users")
    for uid, uname, pw_hash in c3.fetchall():
        if pw_hash and not pw_hash.startswith('pbkdf2:') and not pw_hash.startswith('scrypt:'):
            real_pw = _known_pws.get(uname, "")
            if real_pw:
                c3.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(real_pw), uid))
            else:
                c3.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash("changeme"), uid))
    conn3.commit()
    conn3.close()

def get_tasks():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM tasks WHERE active=1 ORDER BY sort_order")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows

def get_all_tasks():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM tasks ORDER BY sort_order")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows

def get_task_dicts():
    tasks = get_tasks()
    keys = [t["task_key"] for t in tasks]
    labels = {t["task_key"]: t["label"] for t in tasks}
    units = {t["task_key"]: t["unit"] for t in tasks}
    return keys, labels, units

def get_staff_tasks_from_db():
    """Load staff task assignments from DB (tasks column authoritative, targets fallback)."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT staff_id, tasks, targets FROM team_members")
    result = {}
    for row in c.fetchall():
        sid = row["staff_id"]
        if row["tasks"] is not None and row["tasks"] != "":
            try:
                t = json.loads(row["tasks"])
                if isinstance(t, list):
                    result[sid] = t
                    continue
            except:
                pass
        try:
            t = json.loads(row["targets"]) if row["targets"] else {}
        except:
            t = {}
        if t:
            result[sid] = list(t.keys())
        else:
            result[sid] = STAFF_TASKS.get(sid, [])
    conn.close()
    for s in get_staff_order():
        if s not in result:
            result[s] = STAFF_TASKS.get(s, [])
    return result

def get_all_responses():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM responses ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()
    return rows

def add_response(data):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('INSERT INTO responses (timestamp,date,q1,q2,q3,q4,q5,q6,q7) VALUES (?,?,?,?,?,?,?,?,?)',
              (datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
               data.get('date',''), data.get('q1',''), data.get('q2',''),
               data.get('q3',''), data.get('q4',''), data.get('q5',''),
               data.get('q6',''), data.get('q7','')))
    conn.commit()
    conn.close()

def parse_answer(text):
    if not text or not str(text).strip():
        return []
    return [(n.strip().upper(), float(q)) for n, q in re.findall(r'([A-Za-z]+)\s*-\s*(\d+(?:\.\d+)?)', str(text))]

def process_responses(month=None, year=None):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM responses ORDER BY id")
    rows = c.fetchall()
    conn.close()
    daily = {}
    monthly_totals = {}
    task_keys_db, _, _ = get_task_dicts()
    answer_keys = task_keys_db[:7]
    # Pad to 7 if fewer tasks
    while len(answer_keys) < 7:
        answer_keys.append("")
    for row in rows:
        try:
            dt = datetime.strptime(row[2], "%d/%m/%Y")
        except:
            try:
                dt = datetime.strptime(row[2], "%Y-%m-%d")
            except:
                continue
        day = dt.day
        rmonth = dt.month
        ryear = dt.year
        if month and rmonth != month:
            continue
        if year and ryear != year:
            continue
        for i, answer in enumerate(row[3:10]):
            if not answer or not str(answer).strip():
                continue
            for name, qty in parse_answer(answer):
                daily.setdefault(day, {}).setdefault(name, {})
                daily[day][name][answer_keys[i]] = daily[day][name].get(answer_keys[i], 0) + qty
                task_key = answer_keys[i]
                monthly_totals.setdefault(name, {})
                monthly_totals[name][task_key] = monthly_totals[name].get(task_key, 0) + qty
                monthly_totals[name]["_total"] = monthly_totals[name].get("_total", 0) + qty
    return daily, monthly_totals

def calc_scores(daily, year=None, month=None):
    if not year or not month:
        now = datetime.now()
        year, month = now.year, now.month
    targets = get_monthly_targets(year, month)
    staff_tasks = get_staff_tasks_from_db()
    scores = {}
    for s in get_staff_order():
        scores[s] = {"total": 0, "tasks": {}, "days_active": 0, "achievement": {}}
        for day in range(1, 32):
            day_data = daily.get(day, {}).get(s, {})
            if day_data:
                scores[s]["days_active"] += 1
            for tk in staff_tasks.get(s, STAFF_TASKS.get(s, [])):
                val = day_data.get(tk, 0)
                t = targets.get(s, {}).get(tk, 1)
                scores[s]["tasks"][tk] = scores[s]["tasks"].get(tk, 0) + val
                pct = min((val / t) * 100, 150) if t > 0 else 0
                scores[s]["achievement"][tk] = scores[s]["achievement"].get(tk, 0) + pct
        for tk in staff_tasks.get(s, STAFF_TASKS.get(s, [])):
            scores[s]["total"] += scores[s]["tasks"].get(tk, 0)
    ranked = sorted(scores.items(), key=lambda x: x[1]["total"], reverse=True)
    for i, (s, data) in enumerate(ranked):
        scores[s]["rank"] = i + 1
    return scores

def get_team_members():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM team_members ORDER BY id")
    rows = c.fetchall()
    conn.close()
    return rows

def get_chart_data():
    """Get monthly category totals, staff rankings, and KPIs for the dashboard chart."""
    now = datetime.now()
    task_keys_db, _, _ = get_task_dicts()
    answer_keys = task_keys_db[:7]
    while len(answer_keys) < 7:
        answer_keys.append("")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM responses ORDER BY id")
    rows = c.fetchall()
    conn.close()
    # Collect up to 10 months of data ending at current month
    monthly_cat = {}  # {month_idx: {task_key: total}}
    monthly_staff = {}  # {month_idx: {staff: total}}
    all_months = []
    seen_months = set()
    for row in rows:
        try:
            dt = datetime.strptime(row[2], "%d/%m/%Y")
        except:
            try:
                dt = datetime.strptime(row[2], "%Y-%m-%d")
            except:
                continue
        ym = (dt.year, dt.month)
        if ym not in seen_months:
            seen_months.add(ym)
            all_months.append(ym)
    all_months.sort()
    recent_months = all_months[-10:] if len(all_months) > 10 else all_months
    month_idx_map = {ym: i for i, ym in enumerate(recent_months)}
    for row in rows:
        try:
            dt = datetime.strptime(row[2], "%d/%m/%Y")
        except:
            try:
                dt = datetime.strptime(row[2], "%Y-%m-%d")
            except:
                continue
        ym = (dt.year, dt.month)
        if ym not in month_idx_map:
            continue
        idx = month_idx_map[ym]
        for i, answer in enumerate(row[3:10]):
            if not answer or not str(answer).strip():
                continue
            for name, qty in parse_answer(answer):
                tk = answer_keys[i]
                monthly_cat.setdefault(idx, {})
                monthly_cat[idx][tk] = monthly_cat[idx].get(tk, 0) + qty
                monthly_staff.setdefault(idx, {})
                monthly_staff[idx][name] = monthly_staff[idx].get(name, 0) + qty
    month_labels = [datetime(y, m, 1).strftime("%b").upper() for y, m in recent_months]
    # Current month data - find the index matching now.month/now.year
    cur_idx = None
    for i, (y, m) in enumerate(recent_months):
        if y == now.year and m == now.month:
            cur_idx = i
            break
    # If current month not in data, fall back to last month with data
    if cur_idx is None:
        cur_idx = len(recent_months) - 1 if recent_months else 0
    cur_cat = monthly_cat.get(cur_idx, {})
    cur_staff = monthly_staff.get(cur_idx, {})
    prev_idx = cur_idx - 1 if cur_idx > 0 else None
    prev_cat = monthly_cat.get(prev_idx, {}) if prev_idx is not None else {}
    # Staff rankings for current month
    ranked_staff = sorted(cur_staff.items(), key=lambda x: x[1], reverse=True)[:10]
    max_score = ranked_staff[0][1] if ranked_staff else 1
    # KPIs
    total_sales = sum(cur_cat.values())
    prev_total = sum(prev_cat.values())
    pct_change = round(((total_sales - prev_total) / prev_total * 100), 1) if prev_total > 0 else 0
    top_name = ranked_staff[0][0] if ranked_staff else "—"
    top_score = ranked_staff[0][1] if ranked_staff else 0
    # Average daily: count unique days with data in current month
    days_with_data = len([i for i in range(len(recent_months)) if i == cur_idx])
    active_days_count = 0
    for row in rows:
        try:
            dt = datetime.strptime(row[2], "%d/%m/%Y")
        except:
            try:
                dt = datetime.strptime(row[2], "%Y-%m-%d")
            except:
                continue
        ym = (dt.year, dt.month)
        if ym in recent_months:
            idx = month_idx_map[ym]
            if idx == cur_idx:
                day_key = dt.day
                active_days_count = max(active_days_count, day_key)
    avg_daily = round(total_sales / max(active_days_count, 1), 0)
    # Category totals for current month
    cat_totals = {tk: int(cur_cat.get(tk, 0)) for tk in TASK_KEYS if cur_cat.get(tk, 0) > 0}
    # Previous month totals per category (for trend arrows per category)
    prev_cat_totals = {tk: int(prev_cat.get(tk, 0)) for tk in TASK_KEYS}
    # Compute y_max for chart scaling
    all_vals = []
    for idx in monthly_cat:
        for tk in ["lubes","combo","pastry","shellapp"]:
            all_vals.append(monthly_cat[idx].get(tk, 0))
    y_max_val = max(all_vals) * 1.2 if all_vals and max(all_vals) > 0 else 2000
    return {
        "month_labels": month_labels,
        "monthly_cat": monthly_cat,
        "cur_cat": cur_cat,
        "cur_month_label": datetime(now.year, now.month, 1).strftime("%B %Y").upper(),
        "ranked_staff": ranked_staff,
        "max_score": max_score,
        "total_sales": int(total_sales),
        "pct_change": pct_change,
        "top_name": top_name,
        "top_score": int(top_score),
        "avg_daily": int(avg_daily),
        "cat_totals": cat_totals,
        "num_categories": len(cat_totals),
        "recent_months": recent_months,
        "month_idx_map": month_idx_map,
        "y_max_val": y_max_val,
    }

# ============== AUTH HELPERS ==============
def hash_pw(pw):
    return generate_password_hash(pw)

def generate_csrf_token():
    if '_csrf_token' not in session:
        session['_csrf_token'] = _secrets.token_hex(32)
    return session['_csrf_token']

def validate_csrf_token():
    token = request.form.get('_csrf_token', '') or request.headers.get('X-CSRF-Token', '')
    if not token or token != session.get('_csrf_token', ''):
        return False
    return True

app.jinja_env.globals['csrf_token'] = generate_csrf_token

# ============== RATE LIMITING ==============
import time as _time
_rate_limit_store = {}

def _check_rate_limit(key, max_attempts=5, window=60):
    now = _time.time()
    if key not in _rate_limit_store:
        _rate_limit_store[key] = []
    _rate_limit_store[key] = [t for t in _rate_limit_store[key] if now - t < window]
    if len(_rate_limit_store[key]) >= max_attempts:
        return False
    _rate_limit_store[key].append(now)
    return True

def get_user(username):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE username=? OR email=?", (username, username))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            flash("Please log in first", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            flash("Please log in first", "warning")
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            flash("Admin access required", "error")
            return redirect(url_for("index"))
        return f(*args, **kwargs)
    return decorated

from flask import session

@app.before_request
def csrf_protect():
    if request.method in ('POST', 'PUT', 'DELETE'):
        if request.path in ('/login', '/forgot-password', '/verify-code', '/reset-password', '/profile/send-code', '/profile/verify', '/profile/reset-password', '/api/sync'):
            return
        if not validate_csrf_token():
            flash("Session expired. Please try again.", "error")
            return redirect(request.referrer or url_for("index"))

@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' https://cdn.tailwindcss.com; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com; "
        "img-src 'self' data:; "
        "connect-src 'self'"
    )
    return response

# ============== HTML TEMPLATES ==============
LOGIN_TPL = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>OneTrack Login</title>
<link rel="icon" type="image/x-icon" href="/static/images/favicon.ico">
<link rel="icon" type="image/png" sizes="64x64" href="/static/images/favicon-64x64.png">
<link rel="icon" type="image/png" sizes="32x32" href="/static/images/favicon-32x32.png">
<link rel="icon" type="image/png" sizes="16x16" href="/static/images/favicon-16x16.png">
<link rel="manifest" href="/static/manifest.json">
<meta name="theme-color" content="#1A1A1A">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<link rel="apple-touch-icon" href="/static/images/icon-192x192.png">
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<script>tailwind.config={theme:{extend:{colors:{dark:'#1A1A1A','dark-card':'#262626','dark-border':'#333333','dark-hover':'#2A2A2A',ny:'#FFD700','ny-bright':'#FFE44D',red:'#FF0000','on-dark':'#E5E5E5','on-dark-muted':'#A3A3A3','on-dark-dim':'#555555'},fontFamily:{sans:['Inter','sans-serif']}}}}
</script>
<style>
body{font-family:'Inter',sans-serif;color:#E5E5E5;min-height:100vh;min-height:100dvh}
.login-bg{background:#1A1A1A;position:relative;overflow:hidden}
.login-bg::before{content:'';position:fixed;inset:0;background-size:cover;background-position:center;opacity:0.15;z-index:0}
@media(min-width:1024px){.login-bg::before{background-image:url('/static/images/login-bg-desktop.png')}}
@media(max-width:1023px){.login-bg::before{background-image:url('/static/images/login-bg-mobile.png')}}
.login-bg::after{content:'';position:fixed;inset:0;background:linear-gradient(180deg,rgba(26,26,26,0.3) 0%,rgba(26,26,26,0.95) 100%);z-index:0}
.login-content{position:relative;z-index:1}
input:focus{outline:none;border-color:#FFD700;box-shadow:0 0 0 2px rgba(255,215,0,0.15)}
.step-pill{display:inline-flex;align-items:center;gap:4px;padding:4px 8px;border-radius:6px;font-size:10px;font-weight:700;white-space:nowrap}
.step-active{background:#FFD700;color:#1A1A1A}
.step-done{background:#FFD700;color:#1A1A1A}
.step-pending{background:#333;color:#666}
.step-arrow{color:#555;font-size:10px}
</style>
</head>
<body class="login-bg flex items-center justify-center px-4 py-6">
<div class="login-content w-full max-w-xs">
<!-- Shell Logo -->
<div class="text-center mb-3">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-10 h-10 rounded-full object-cover mx-auto mb-1.5 border-2 border-ny/30">
<h1 class="text-lg font-black text-white">OneTrack</h1>
<p class="text-on-dark-muted text-[10px] mt-0.5">Shell Bandar Mahkota Cheras</p>
</div>

<!-- Login Card -->
<div class="bg-dark-card/90 backdrop-blur-sm border border-dark-border rounded-lg p-3">
<h2 class="text-xs font-bold text-white mb-2">Sign In</h2>

{% with messages = get_flashed_messages(with_categories=true) %}
{% if messages %}
{% for cat, msg in messages %}
<div class="mb-4 px-3 py-2 rounded text-sm {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}
{% endif %}
{% endwith %}

<form method="POST" action="/login" class="space-y-2">
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">Username or Email</label>
<input type="text" name="username" required autocomplete="username" placeholder="Enter username or email"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs placeholder-on-dark-dim" style="min-height:38px">
</div>
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">Password</label>
<input type="password" name="password" required autocomplete="current-password" placeholder="Enter password"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs placeholder-on-dark-dim" style="min-height:38px">
</div>
<div class="flex items-center justify-between">
<label class="flex items-center gap-1.5 cursor-pointer">
<input type="checkbox" name="remember" class="w-3.5 h-3.5 rounded border-dark-border bg-dark accent-ny">
<span class="text-[11px] text-on-dark-dim">Remember me</span>
</label>
</div>
<button type="submit" class="w-full bg-ny text-dark font-bold py-2 rounded text-xs hover:bg-ny-bright transition mt-1" style="min-height:40px">
SIGN IN
</button>
</form>

<div class="text-center mt-2">
<a href="/forgot-password" class="text-ny text-[11px] hover:underline">Forgot Username or Password?</a>
</div>

<div class="flex items-center gap-3 my-2">
<div class="flex-1 h-px bg-dark-border"></div>
<span class="text-on-dark-dim text-[9px] uppercase">or</span>
<div class="flex-1 h-px bg-dark-border"></div>
</div>
<a href="/" class="block text-center text-on-dark-muted text-[11px] hover:text-ny transition-colors py-0.5">
Continue as Guest
</a>
</div>

<p class="text-center text-on-dark-dim text-[9px] mt-2">&copy; 2026 Shell Prosper Niaga</p>
</div>
</body>
</html>'''

AUTH_LAYOUT = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>OneTrack - Shell Prosper Niaga</title>
<link rel="icon" type="image/x-icon" href="/static/images/favicon.ico">
<link rel="icon" type="image/png" sizes="64x64" href="/static/images/favicon-64x64.png">
<link rel="icon" type="image/png" sizes="32x32" href="/static/images/favicon-32x32.png">
<link rel="icon" type="image/png" sizes="16x16" href="/static/images/favicon-16x16.png">
<link rel="manifest" href="/static/manifest.json">
<meta name="theme-color" content="#1A1A1A">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<link rel="apple-touch-icon" href="/static/images/icon-192x192.png">
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
<script>tailwind.config={theme:{extend:{colors:{dark:'#1A1A1A','dark-card':'#262626','dark-border':'#333333','dark-hover':'#2A2A2A',ny:'#FFD700','ny-bright':'#FFE44D',red:'#FF0000','on-dark':'#E5E5E5','on-dark-muted':'#A3A3A3','on-dark-dim':'#555555'},fontFamily:{sans:['Inter','sans-serif']}}}}
</script>
<style>
body{font-family:'Inter',sans-serif;color:#E5E5E5;min-height:100vh;min-height:100dvh}
.login-bg{background:#1A1A1A;position:relative;overflow:hidden}
.login-bg::before{content:'';position:fixed;inset:0;background-size:cover;background-position:center;opacity:0.15;z-index:0}
@media(min-width:1024px){.login-bg::before{background-image:url('/static/images/login-bg-desktop.png')}}
@media(max-width:1023px){.login-bg::before{background-image:url('/static/images/login-bg-mobile.png')}}
.login-bg::after{content:'';position:fixed;inset:0;background:linear-gradient(180deg,rgba(26,26,26,0.3) 0%,rgba(26,26,26,0.95) 100%);z-index:0}
.login-content{position:relative;z-index:1}
input:focus{outline:none;border-color:#FFD700;box-shadow:0 0 0 2px rgba(255,215,0,0.15)}
.step-pill{display:inline-flex;align-items:center;gap:4px;padding:4px 8px;border-radius:6px;font-size:10px;font-weight:700;white-space:nowrap}
.step-active{background:#FFD700;color:#1A1A1A}
.step-done{background:#FFD700;color:#1A1A1A}
.step-pending{background:#333;color:#666}
.step-arrow{color:#555;font-size:10px}
</style>
</head>
<body class="login-bg flex items-center justify-center px-4 py-6">
<div class="login-content w-full max-w-xs">
{{ content|safe }}
</div>
</body>
</html>'''

LAYOUT = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
{% if autorefresh %}<meta http-equiv="refresh" content="5">{% endif %}
<title>OneTrack - Shell Prosper Niaga</title>
<link rel="icon" type="image/x-icon" href="/static/images/favicon.ico">
<link rel="icon" type="image/png" sizes="64x64" href="/static/images/favicon-64x64.png">
<link rel="icon" type="image/png" sizes="32x32" href="/static/images/favicon-32x32.png">
<link rel="icon" type="image/png" sizes="16x16" href="/static/images/favicon-16x16.png">
<link rel="manifest" href="/static/manifest.json">
<meta name="theme-color" content="#1A1A1A">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<link rel="apple-touch-icon" href="/static/images/icon-192x192.png">
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet">
<script>
tailwind.config={theme:{extend:{colors:{
"ny":"#FFD700","ny-dim":"#B8960F","ny-bright":"#FFED4A",
"red":"#FF0000","red-dim":"#CC0000",
"dark":"#1A1A1A","dark-card":"#262626","dark-border":"#333333","dark-hover":"#2A2A2A",
"green-neon":"#39FF14",
"on-dark":"#FFFFFF","on-dark-muted":"#AAAAAA","on-dark-dim":"#666666"
},fontFamily:{sans:["Inter","sans-serif"],mono:["JetBrains Mono","monospace"]}}}}
</script>
<style>
.material-symbols-outlined{font-variation-settings:'FILL' 0,'wght' 400,'GRAD' 0,'opsz' 24}
*{scrollbar-width:thin;scrollbar-color:#333 #1A1A1A}
::-webkit-scrollbar{width:6px}::-webkit-scrollbar-track{background:#1A1A1A}::-webkit-scrollbar-thumb{background:#333;border-radius:3px}
input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}
.glow{box-shadow:0 0 20px rgba(255,215,0,0.15)}
.glow-strong{box-shadow:0 0 40px rgba(255,215,0,0.25),0 0 80px rgba(255,215,0,0.1)}
.card-hover{transition:all 0.2s}.card-hover:hover{transform:translateY(-2px);box-shadow:0 8px 25px rgba(255,215,0,0.1)}
.nav-active{background:rgba(255,215,0,0.1);border-left:3px solid #FFD700;color:#FFD700}
input:focus,select:focus{outline:none;border-color:#FFD700;box-shadow:0 0 0 2px rgba(255,215,0,0.2)}
.sidebar-open{transform:translateX(0) !important}
@media(max-width:1023px){
.sidebar-open~#sidebar-overlay{display:block !important}
}
.day-grid{display:flex;flex-wrap:wrap;gap:4px}
.day-btn{width:32px;height:32px;display:flex;align-items:center;justify-content:center;border-radius:6px;font-size:11px;font-family:'JetBrains Mono',monospace;font-weight:600;cursor:pointer;border:1px solid #333;color:#666;flex-shrink:0}
.day-btn.active{background:rgba(255,215,0,0.2);color:#FFD700;border-color:rgba(255,215,0,0.3)}
.day-btn.has-data{background:rgba(57,255,20,0.1);color:#39FF14;border-color:rgba(57,255,20,0.3)}
@media(max-width:640px){
.day-grid{overflow-x:auto;flex-wrap:nowrap;padding-bottom:4px;-webkit-overflow-scrolling:touch}
.day-btn{width:36px;height:36px;font-size:12px;flex-shrink:0}
}
.table-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch}
.table-scroll::-webkit-scrollbar{height:4px}
.table-scroll::-webkit-scrollbar-thumb{background:#FFD700;border-radius:2px}
.podium-1{transform:scale(1.05);z-index:3}
.podium-2{transform:scale(0.92);z-index:2}
.podium-3{transform:scale(0.92);z-index:2}
@media(max-width:640px){
.podium-1,.podium-2,.podium-3{transform:none}
.podium-wrap{gap:4px!important;padding:0 4px!important}
.podium-card{padding:10px 8px 16px!important;border-radius:12px!important}
.podium-card .w-28,.podium-card .h-28{width:56px!important;height:56px!important}
.podium-card .w-36,.podium-card .h-36{width:56px!important;height:56px!important}
.podium-card .text-3xl,.podium-card .text-5xl{font-size:16px!important;margin-bottom:4px!important}
.podium-card .text-xl,.podium-card .text-3xl{font-size:12px!important}
.podium-card .text-lg,.podium-card .text-2xl{font-size:12px!important}
.podium-card .text-sm{font-size:10px!important}
.podium-card .text-base{font-size:11px!important}
.podium-card .pod-pct{font-size:13px!important}
.podium-card .task-row{font-size:7px!important;padding:2px 6px!important}
}
@media(max-width:640px){
.podium-wrap{flex-wrap:nowrap!important}
}
@keyframes fadeInUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
@keyframes pulse-glow{0%,100%{box-shadow:0 0 20px rgba(255,215,0,0.2)}50%{box-shadow:0 0 40px rgba(255,215,0,0.4)}}
.animate-fadeInUp{animation:fadeInUp 0.5s ease-out}
.animate-pulse-glow{animation:pulse-glow 2s ease-in-out infinite}
@keyframes toast-in{from{transform:translateX(100%);opacity:0}to{transform:translateX(0);opacity:1}}
@keyframes toast-out{from{transform:translateX(0);opacity:1}to{transform:translateX(100%);opacity:0}}
.toast-show{animation:toast-in 0.3s ease-out}
.toast-hide{animation:toast-out 0.3s ease-in forwards}
</style>
</head>
<body class="bg-dark font-sans text-on-dark min-h-screen">
<!-- Toast Container -->
<div id="toast-container" class="fixed top-4 right-4 z-[200] space-y-2"></div>
<!-- Delete Confirmation Modal -->
<div id="deleteModal" class="fixed inset-0 z-[300] flex items-center justify-center" style="display:none">
<div class="fixed inset-0 bg-black/60 backdrop-blur-sm" onclick="closeDeleteModal(false)"></div>
<div class="relative bg-[#1A1A1A] border border-dark-border rounded-xl p-6 max-w-sm w-full mx-4 shadow-2xl">
<div class="flex items-center gap-3 mb-4">
<div class="w-10 h-10 rounded-full bg-red/10 flex items-center justify-center shrink-0">
<span class="material-symbols-outlined text-red text-xl">delete</span>
</div>
<div>
<h3 class="text-white font-bold text-sm">Confirm Delete</h3>
<p id="deleteModalMsg" class="text-on-dark-muted text-xs mt-1">Are you sure you want to delete this?</p>
</div>
</div>
<div class="flex gap-3 justify-end">
<button onclick="closeDeleteModal(false)" class="px-4 py-2 rounded-lg text-xs font-semibold text-on-dark-dim bg-dark border border-dark-border hover:bg-dark-hover transition">Cancel</button>
<button id="deleteModalYes" onclick="closeDeleteModal(true)" class="px-4 py-2 rounded-lg text-xs font-bold text-white bg-red hover:bg-red/80 transition">Yes, Delete</button>
</div>
</div>
</div>
<script>
var _deleteCb=null;
function showDeleteModal(msg,cb){
document.getElementById('deleteModalMsg').textContent=msg||'Are you sure you want to delete this?';
_deleteCb=cb;
document.getElementById('deleteModal').style.display='flex';
}
function closeDeleteModal(confirmed){
document.getElementById('deleteModal').style.display='none';
if(_deleteCb){_deleteCb(confirmed);_deleteCb=null}
}
</script>
{% with messages = get_flashed_messages(with_categories=true) %}
{% if messages %}
<script>
(function(){
var msgs={{ messages|tojson }};
msgs.forEach(function(m){
var cat=m[0], txt=m[1];
var isErr=cat==='error';
var el=document.createElement('div');
el.className='toast-show flex items-center gap-2 px-4 py-3 rounded-lg shadow-lg border text-sm font-medium max-w-sm';
el.style.cssText=isErr?'background:#2A1A1A;border-color:rgba(255,0,0,0.3);color:#FF6666':'background:#1A2A1A;border-color:rgba(57,255,20,0.3);color:#39FF14';
var icon=document.createElement('span');
icon.className='material-symbols-outlined text-lg';
icon.textContent=isErr?'error':'check_circle';
var msgSpan=document.createElement('span');
msgSpan.textContent=txt;
el.appendChild(icon);
el.appendChild(msgSpan);
document.getElementById('toast-container').appendChild(el);
setTimeout(function(){el.classList.remove('toast-show');el.classList.add('toast-hide');setTimeout(function(){el.remove()},300)},3500);
});
})();
</script>
{% endif %}
{% endwith %}
<script>
var sidebarOpen=false;
function toggleSidebar(){
sidebarOpen=!sidebarOpen;
var sb=document.getElementById('sidebar');
var ov=document.getElementById('sidebar-overlay');
if(sidebarOpen){sb.classList.add('sidebar-open');ov.style.display='block'}
else{sb.classList.remove('sidebar-open');ov.style.display='none'}
}
function closeSidebar(){
sidebarOpen=false;
document.getElementById('sidebar').classList.remove('sidebar-open');
document.getElementById('sidebar-overlay').style.display='none';
}
window.addEventListener('resize',function(){
if(window.innerWidth>=1024){closeSidebar();document.getElementById('sidebar-overlay').style.display='none'}
});
</script>
<!-- Mobile Header -->
<header class="lg:hidden fixed top-0 left-0 right-0 z-50 bg-dark-card border-b border-dark-border h-14 flex items-center px-4 justify-between">
<div class="flex items-center gap-2">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-8 h-8 rounded-lg object-cover">
<span class="font-bold text-ny text-sm">Shell Prosper Niaga</span>
</div>
<button onclick="toggleSidebar()" class="p-2 hover:bg-dark-hover rounded-lg" aria-label="Toggle menu">
<span class="material-symbols-outlined text-on-dark">menu</span>
</button>
</header>
<!-- Sidebar -->
<aside id="sidebar" class="fixed top-0 left-0 h-full w-64 bg-dark-card border-r border-dark-border z-[60] flex flex-col -translate-x-full lg:translate-x-0 transition-transform duration-200">
<div class="p-4 border-b border-dark-border">
<div class="flex items-center gap-3">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-10 h-10 rounded-lg object-cover">
<div>
<h1 class="font-bold text-ny text-lg leading-tight">Shell</h1>
<p class="text-xs text-on-dark-muted">Prosper Niaga - BMC</p>
</div>
</div>
</div>
<div class="px-3 pt-4 pb-2">
<span class="text-xs font-semibold text-on-dark-dim uppercase tracking-wider px-3">Navigation</span>
</div>
<nav class="flex-1 p-3 flex flex-col gap-1">
<a href="/" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='dashboard' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">dashboard</span>Dashboard
</a>
<a href="/leaderboard" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='leaderboard' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">leaderboard</span>Leaderboard
</a>
<a href="/daily" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='daily' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">calendar_month</span>Daily Tracking
</a>
<a href="/form" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='form' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">edit_note</span>Data Entry
</a>
{% if session.get('role')=='admin' %}
<a href="/team" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='team' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">groups</span>Team Members
</a>
<a href="/admin/tasks" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='admin_tasks' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">task_alt</span>Task Editor
</a>
<a href="/admin/records" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='admin_records' or page=='admin_record_edit' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">edit_document</span>Records
</a>
<a href="/analysis" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all
{% if page=='analysis' %}nav-active{% else %}text-on-dark-muted hover:bg-dark-hover{% endif %}">
<span class="material-symbols-outlined text-lg">analytics</span>Sales Analysis
</a>
{% endif %}
</nav>
<div class="p-3 border-t border-dark-border">
{% if session.get('user') %}
<a href="/profile" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-2 mb-2 rounded-lg hover:bg-dark-hover transition-all cursor-pointer">
<div class="w-8 h-8 rounded-full bg-ny/20 flex items-center justify-center">
<span class="text-ny text-sm font-bold">{{ session.get('display_name','U')[0] }}</span>
</div>
<div>
<p class="text-xs font-semibold text-on-dark">{{ session.get('display_name', session.get('user')) }}</p>
<p class="text-xs text-on-dark-muted capitalize">{{ session.get('role','user') }}</p>
</div>
</a>
<a href="/logout" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-2 rounded-lg text-sm font-medium text-red hover:bg-red/10 transition-all">
<span class="material-symbols-outlined text-lg">logout</span>Sign Out
</a>
{% else %}
<a href="/login" onclick="closeSidebar()" class="flex items-center gap-3 px-4 py-2 rounded-lg text-sm font-medium text-ny hover:bg-ny/10 transition-all">
<span class="material-symbols-outlined text-lg">login</span>Sign In
</a>
{% endif %}
</div>
</aside>
<!-- Overlay for mobile sidebar -->
<div id="sidebar-overlay" onclick="closeSidebar()" class="fixed inset-0 bg-black/60 z-[55] hidden lg:hidden" style="display:none"></div>
<!-- Main Content -->
<main class="lg:ml-64 min-h-screen pt-14 lg:pt-0">
{{ content|safe }}
</main>
<!-- Offline Overlay -->
<div id="offline-overlay" class="hidden fixed inset-0 z-[9999] bg-dark/95 backdrop-blur-sm flex items-center justify-center p-4">
  <div class="w-full max-w-sm text-center">
    <div class="mb-6">
      <div class="relative inline-block">
        <span class="material-symbols-outlined text-7xl text-red">wifi_off</span>
        <span class="absolute -top-1 -right-1 w-4 h-4 bg-red rounded-full animate-ping"></span>
        <span class="absolute -top-1 -right-1 w-4 h-4 bg-red rounded-full"></span>
      </div>
    </div>
    <h2 class="text-xl font-bold text-white mb-2">No Internet Connection</h2>
    <p class="text-on-dark-muted text-sm mb-6">Please check your network settings and try again.</p>
    <div class="bg-dark-card border border-dark-border rounded-lg p-4 mb-6 text-left">
      <div class="space-y-2 text-xs font-mono">
        <div class="flex justify-between"><span class="text-on-dark-dim">WiFi / Cellular</span><span class="text-red">No Link</span></div>
        <div class="flex justify-between"><span class="text-on-dark-dim">Gateway</span><span class="text-red">Unreachable</span></div>
        <div class="flex justify-between"><span class="text-on-dark-dim">Server</span><span class="text-red">Offline</span></div>
      </div>
    </div>
    <button onclick="location.reload()" class="w-full bg-ny text-dark py-3 rounded-lg font-bold text-sm hover:bg-ny-bright transition flex items-center justify-center gap-2 mb-3">
      <span class="material-symbols-outlined text-base">refresh</span>
      Try Again
    </button>
    <a href="/" class="text-ny text-sm hover:underline inline-flex items-center gap-1">
      <span class="material-symbols-outlined text-base">dashboard</span>Go to Dashboard (Offline Mode)
    </a>
  </div>
</div>
<script>
(function(){
  var o=document.getElementById('offline-overlay');
  function up(){if(!navigator.onLine){o.classList.remove('hidden');}else{o.classList.add('hidden');}}
  window.addEventListener('online',up);window.addEventListener('offline',up);up();
})();
</script>
<script>
document.querySelectorAll('form[method="POST"],form[method="post"]').forEach(function(f){
  if(!f.querySelector('input[name="_csrf_token"]')){
    var inp=document.createElement('input');
    inp.type='hidden';inp.name='_csrf_token';
    inp.value='{{ csrf_token() }}';
    f.appendChild(inp);
  }
});
</script>
<script>
if('serviceWorker' in navigator){
  navigator.serviceWorker.register('/static/sw.js').then(function(reg){
    reg.onupdatefound=function(){var nw=reg.installing;nw.onstatechange=function(){if(nw.state==='installed'&&navigator.serviceWorker.controller){nw.postMessage('retry-sync')}}
    }
  }).catch(function(){});
}
window.addEventListener('online',function(){if(navigator.serviceWorker&&navigator.serviceWorker.controller){navigator.serviceWorker.controller.postMessage('retry-sync')}});
</script>
</body></html>'''

# ============== DASHBOARD ==============
DASHBOARD = '''
{% set cd = chart_data %}
{% set cat_colors = {"lubes":"#3B82F6","combo":"#10B981","pastry":"#FFD700","shellapp":"#A855F7"} %}
{% set cat_text = {"lubes":"#60A5FA","combo":"#39FF14","pastry":"#FFD700","shellapp":"#C084FC"} %}
{% set cat_id = {"lubes":"grad-lubes","combo":"grad-combo","pastry":"grad-pastry","shellapp":"grad-app"} %}
<style>
.live-beacon{animation:pulse-fast 2s infinite ease-in-out}
@keyframes pulse-fast{0%,100%{opacity:1;transform:scale(1)}50%{opacity:0.4;transform:scale(0.9)}}
.hud-grid{background-size:32px 32px;background-image:linear-gradient(to right,rgba(255,255,255,0.025) 1px,transparent 1px),linear-gradient(to bottom,rgba(255,255,255,0.025) 1px,transparent 1px)}
.chart-tip{position:absolute;pointer-events:none;background:#20201f;border:1px solid #FFD700;border-radius:6px;padding:6px 10px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#e5e2e1;z-index:50;white-space:nowrap;opacity:0;transition:opacity 0.15s}
.chart-tip.show{opacity:1}
</style>
<div class="p-4 lg:p-8 max-w-7xl mx-auto space-y-4 sm:space-y-6 hud-grid min-h-screen" style="font-family:'Inter',sans-serif">
<!-- SUBHEADER -->
<div class="flex flex-col md:flex-row md:items-center justify-between gap-3 bg-[#20201f] border border-[#4d4732] rounded-lg p-3">
<div class="flex items-center gap-3">
<div class="p-1.5 bg-[#20201f] border border-[#4d4732] rounded">
<span class="material-symbols-outlined text-[#FFD700] text-base">query_stats</span>
</div>
<div>
<h1 class="text-[18px] font-semibold text-[#FFD700] tracking-wide" style="font-family:'Inter',sans-serif">OneTrack // Sales Analytics Telemetry</h1>
<p class="text-[11px] font-semibold tracking-wider text-[#888]" style="font-family:'JetBrains Mono',monospace">DIAGNOSTIC PIPELINE // REAL-TIME UNIT STREAM // {{ current_month }} CYCLE</p>
</div>
</div>
<div class="flex flex-wrap items-center gap-3">
<div class="flex items-center gap-1.5 px-2.5 py-1 bg-[#10b981]/10 border border-[#10b981]/40 rounded text-[11px] font-semibold text-[#39FF14]" style="font-family:'JetBrains Mono',monospace">
<span class="w-1.5 h-1.5 rounded-full bg-[#39FF14] live-beacon"></span>
<span>100% TELEMETRY FIDELITY</span>
</div>
</div>
</div>
<!-- MAIN CHART + RANKINGS -->
<div class="grid grid-cols-1 lg:grid-cols-12 gap-4 items-start">
<!-- 80% CHART PANEL -->
<section class="lg:col-span-8 flex flex-col bg-[#20201f] border border-[#4d4732] rounded-xl p-4 md:p-6 relative overflow-hidden">
<div class="flex flex-col sm:flex-row sm:items-center justify-between pb-4 border-b border-[#4d4732] gap-3">
<div class="flex items-center gap-2">
<span class="w-2 h-2 rounded-full bg-[#FFD700]"></span>
<h2 class="text-[18px] font-semibold text-[#FFD700] font-bold">{% if view_mode == 'month' %}Daily Sales Trend{% else %}Monthly Sales Trend{% endif %}</h2>
</div>
<form method="GET" class="flex items-end gap-2 flex-wrap">
<div class="flex flex-col gap-1">
<span class="text-[10px] text-[#888] uppercase tracking-wider" style="font-family:'JetBrains Mono',monospace">Year</span>
<select name="year" class="bg-[#111] border border-[#333] rounded-md px-2 py-1.5 text-[12px] text-white outline-none" style="font-family:'JetBrains Mono',monospace">
{% for y in available_years %}
<option value="{{ y }}" {% if y==sel_year %}selected{% endif %}>{{ y }}</option>
{% endfor %}
</select>
</div>
<div class="flex flex-col gap-1">
<span class="text-[10px] text-[#888] uppercase tracking-wider" style="font-family:'JetBrains Mono',monospace">Month</span>
<select name="month" class="bg-[#111] border border-[#333] rounded-md px-2 py-1.5 text-[12px] text-white outline-none" style="font-family:'JetBrains Mono',monospace">
<option value="">ALL MONTHS</option>
{% set month_names = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'] %}
{% for m in range(1, 13) %}
<option value="{{ m }}" {% if view_mode=='month' and m==sel_month %}selected{% endif %}>{{ month_names[m] }}</option>
{% endfor %}
</select>
</div>
<button type="submit" class="bg-[#FFD700] text-[#000] border-none rounded-md px-4 py-1.5 text-[11px] font-bold cursor-pointer hover:bg-[#FFE44D] transition" style="font-family:'JetBrains Mono',monospace">Apply</button>
</form>
</div>
<!-- Legend Pills -->
<div class="grid grid-cols-2 sm:grid-cols-4 gap-2 my-4">
{% for tk in all_tasks %}
{% set val = cd.cat_totals.get(tk, 0) %}
<div class="flex items-center justify-between p-2 bg-[#1c1b1b] border border-[#4d4732] rounded transition-colors cursor-pointer group">
<div class="flex items-center gap-2">
<span class="w-2.5 h-2.5 rounded-full group-hover:scale-110 transition-transform" style="background-color:{{ cat_colors[tk] }}"></span>
<span class="text-[11px] font-semibold tracking-wider text-[#e5e2e1]" style="font-family:'JetBrains Mono',monospace">{{ task_labels.get(tk, tk)|upper }}</span>
</div>
<span class="text-[12px] font-medium" style="font-family:'JetBrains Mono',monospace;color:{{ cat_text[tk] }}">{{ "{:,.0f}".format(val) }}</span>
</div>
{% endfor %}
</div>
<!-- SVG Chart -->
{% set num_points = chart_labels|length %}
{% set cw = 800 %}{% set ch = 320 %}{% set pl = 50 %}{% set pr = 780 %}{% set pt = 35 %}{% set pb = 270 %}
{% set ym = cd.y_max_val %}
<div class="w-full bg-[#0e0e0e] border border-[#4d4732] rounded-lg p-2 md:p-4 relative" id="chart-wrap">
<svg class="w-full h-auto overflow-visible select-none" viewBox="0 0 {{ cw }} {{ ch }}" xmlns="http://www.w3.org/2000/svg" id="chart-svg">
<defs>
<linearGradient id="grad-lubes" x1="0%" x2="0%" y1="0%" y2="100%"><stop offset="0%" stop-color="#3B82F6" stop-opacity="0.32"/><stop offset="100%" stop-color="#3B82F6" stop-opacity="0.0"/></linearGradient>
<linearGradient id="grad-combo" x1="0%" x2="0%" y1="0%" y2="100%"><stop offset="0%" stop-color="#10B981" stop-opacity="0.28"/><stop offset="100%" stop-color="#10B981" stop-opacity="0.0"/></linearGradient>
<linearGradient id="grad-pastry" x1="0%" x2="0%" y1="0%" y2="100%"><stop offset="0%" stop-color="#FFD700" stop-opacity="0.28"/><stop offset="100%" stop-color="#FFD700" stop-opacity="0.0"/></linearGradient>
<linearGradient id="grad-app" x1="0%" x2="0%" y1="0%" y2="100%"><stop offset="0%" stop-color="#A855F7" stop-opacity="0.25"/><stop offset="100%" stop-color="#A855F7" stop-opacity="0.0"/></linearGradient>
</defs>
<!-- Y-Axis -->
{% for i in range(5) %}
{% set y = pb - (i * (pb - pt) / 4) %}
{% set val = (ym / 4 * i)|int %}
<line stroke="#2c2c2c" stroke-dasharray="3,3" stroke-width="1" x1="{{ pl }}" x2="{{ pr }}" y1="{{ y }}" y2="{{ y }}"/>
<text fill="#888" font-family="JetBrains Mono" font-size="10" text-anchor="end" x="{{ pl - 8 }}" y="{{ y + 4 }}">{{ "{:,}".format(val) }}</text>
{% endfor %}
<!-- X positions -->
{% set x_positions = [] %}
{% for i in range(num_points) %}
{% set x = pl + (i * (pr - pl) / ([num_points - 1, 1]|max)) %}
{% if x_positions.append(x) %}{% endif %}
{% endfor %}
<!-- Area fills -->
{% for tk in all_tasks %}
{% set apts = [(pl|string) ~ "," ~ (pb|string)] %}
{% for i in range(num_points) %}
{% set key = (i + 1) if view_mode == 'month' else (i + 1) %}
{% set val = chart_points.get(key, {}).get(tk, 0) %}
{% set x = x_positions[i] %}
{% set y = pb - (val / ym * (pb - pt)) %}
{% set _ = apts.append((x|int|string) ~ "," ~ (y|int|string)) %}
{% endfor %}
{% set _ = apts.append((pr|string) ~ "," ~ (pb|string)) %}
<polygon fill="url(#{{ cat_id[tk] }})" points="{{ apts|join(' ') }}"/>
{% endfor %}
<!-- Spline lines -->
{% for tk in all_tasks %}
{% set pts = [] %}
{% for i in range(num_points) %}
{% set key = (i + 1) if view_mode == 'month' else (i + 1) %}
{% set val = chart_points.get(key, {}).get(tk, 0) %}
{% set x = x_positions[i] %}
{% set y = pb - (val / ym * (pb - pt)) %}
{% set _ = pts.append([x|int, y|int]) %}
{% endfor %}
{% if pts|length > 1 %}
<path d="M{{ pts[0][0] }},{{ pts[0][1] }}{% for i in range(1, pts|length) %}{% set px = pts[i-1][0] %}{% set py = pts[i-1][1] %}{% set cx = pts[i][0] %}{% set cy = pts[i][1] %}{% set cp = ((cx - px) / 3)|int %} C{{ px + cp }},{{ py }} {{ cx - cp }},{{ cy }} {{ cx }},{{ cy }}{% endfor %}" fill="none" stroke="{{ cat_colors[tk] }}" stroke-linecap="round" stroke-width="2.5"/>
{% endif %}
{% endfor %}
<!-- X-Axis labels -->
{% for i in range(num_points) %}
{% set x = x_positions[i] %}
<text fill="{% if i == num_points - 1 %}#FFD700{% else %}#888{% endif %}" font-family="JetBrains Mono" font-size="{% if view_mode == 'year' %}10{% else %}9{% endif %}" text-anchor="middle" x="{{ x }}" y="{{ pb + 22 }}" {% if i == num_points - 1 %}font-weight="700"{% endif %}>{{ chart_labels[i] }}</text>
{% endfor %}
<!-- Hover line (hidden by default) -->
<line id="hover-line" x1="0" x2="0" y1="{{ pt }}" y2="{{ pb }}" stroke="#FFD700" stroke-width="1" stroke-dasharray="4,4" opacity="0"/>
<!-- Hover dots (hidden) -->
{% for tk in all_tasks %}
<circle id="dot-{{ tk }}" cx="0" cy="0" r="4" fill="{{ cat_colors[tk] }}" stroke="#131313" stroke-width="2" opacity="0"/>
{% endfor %}
<!-- Hover data points (invisible wider targets) -->
{% for i in range(num_points) %}
{% set x = x_positions[i] %}
<rect x="{{ x - 15 }}" y="{{ pt }}" width="30" height="{{ pb - pt }}" fill="transparent" class="hover-zone" data-idx="{{ i }}"/>
{% endfor %}
</svg>
<div class="chart-tip" id="chart-tip"></div>
</div>
<!-- Readout bar -->
<div class="mt-4 p-3 bg-[#1c1b1b] border border-[#4d4732] rounded flex flex-wrap items-center justify-between gap-2">
<div class="flex items-center gap-2">
<span class="text-[11px] font-semibold tracking-wider text-[#888]" style="font-family:'JetBrains Mono',monospace">{{ current_month }} AUDIT:</span>
<span class="text-[12px] font-medium text-[#FFD700] font-semibold" style="font-family:'JetBrains Mono',monospace">Total Category Index: {{ "{:,}".format(cd.total_sales|int) }} Units</span>
</div>
<div class="flex items-center gap-4 text-[12px] font-medium" style="font-family:'JetBrains Mono',monospace">
{% for tk in all_tasks %}
<span style="color:{{ cat_text[tk] }}">{{ task_labels.get(tk, tk)|upper }}: {{ "{:,}".format(cd.cat_totals.get(tk, 0)) }}</span>
{% endfor %}
</div>
</div>
</section>
<!-- 20% RANKINGS -->
<section class="lg:col-span-4 bg-[#20201f] border border-[#4d4732] rounded-xl p-4 flex flex-col justify-between h-full">
<div>
<div class="flex items-center justify-between pb-3 border-b border-[#4d4732] mb-3">
<div class="flex items-center gap-2">
<span class="material-symbols-outlined text-[#FFD700] text-base">leaderboard</span>
<h2 class="text-[18px] font-semibold text-[#FFD700] font-bold">Staff Rankings</h2>
</div>
<span class="px-2 py-0.5 bg-[#131313] border border-[#4d4732] rounded text-[11px] font-semibold tracking-wider text-[#FFD700]" style="font-family:'JetBrains Mono',monospace">{{ current_month|upper }}</span>
</div>
<div class="flex items-center justify-between text-[11px] font-semibold tracking-wider text-[#888] px-1 mb-2" style="font-family:'JetBrains Mono',monospace">
<span>OPERATOR // CREW</span>
</div>
<div class="space-y-2">
{% for name, score in cd.ranked_staff %}
{% set pct = (score / cd.max_score * 100)|round(1) if cd.max_score > 0 else 0 %}
{% if loop.index == 1 %}
<div class="p-1.5 bg-[#1c1b1b] border border-[#FFD700]/60 rounded group hover:border-[#FFD700] transition-all">
<div class="flex items-center justify-between text-[12px] font-medium mb-1" style="font-family:'JetBrains Mono',monospace">
<div class="flex items-center gap-1.5">
<span class="px-1.5 py-0.2 bg-[#FFD700] text-[#131313] font-bold rounded text-[11px]">#{{ loop.index }}</span>
<span class="text-[#FFD700] font-bold">{{ name }}</span>
<span class="material-symbols-outlined text-[#FFD700] text-sm">hotel_class</span>
</div>
<span class="text-[#FFD700] font-semibold">{{ "{:,.0f}".format(score) }} sales</span>
</div>
<div class="w-full h-2 bg-[#0e0e0e] rounded-full overflow-hidden border border-[#4d4732]/40">
<div class="h-full bg-[#FFD700] rounded-full" style="width:{{ pct }}%"></div>
</div>
</div>
{% elif loop.index == 2 %}
<div class="p-1.5 bg-[#1c1b1b] border border-[#10B981]/40 rounded group hover:border-[#10B981] transition-all">
<div class="flex items-center justify-between text-[12px] font-medium mb-1" style="font-family:'JetBrains Mono',monospace">
<div class="flex items-center gap-1.5">
<span class="px-1.5 py-0.2 bg-[#10B981]/20 text-[#39FF14] font-bold rounded text-[11px]">#{{ loop.index }}</span>
<span class="text-[#e5e2e1] font-semibold">{{ name }}</span>
</div>
<span class="text-[#39FF14] font-semibold">{{ "{:,.0f}".format(score) }} sales</span>
</div>
<div class="w-full h-1.5 bg-[#0e0e0e] rounded-full overflow-hidden border border-[#4d4732]/30">
<div class="h-full bg-[#10B981] rounded-full" style="width:{{ pct }}%"></div>
</div>
</div>
{% elif loop.index == 3 %}
<div class="p-1.5 bg-[#1c1b1b] border border-[#4d4732] rounded group hover:border-[#FFD700] transition-all">
<div class="flex items-center justify-between text-[12px] font-medium mb-1" style="font-family:'JetBrains Mono',monospace">
<div class="flex items-center gap-1.5">
<span class="px-1.5 py-0.2 bg-[#131313] text-[#FFD700] font-semibold rounded text-[11px]">#{{ loop.index }}</span>
<span class="text-[#e5e2e1]">{{ name }}</span>
</div>
<span class="text-[#FFD700]">{{ "{:,.0f}".format(score) }} sales</span>
</div>
<div class="w-full h-1.5 bg-[#0e0e0e] rounded-full overflow-hidden border border-[#4d4732]/30">
<div class="h-full bg-[#B8960F] rounded-full" style="width:{{ pct }}%"></div>
</div>
</div>
{% else %}
<div class="p-1.5 bg-[#1c1b1b] border border-[#4d4732] rounded hover:border-[#999077] transition-all">
<div class="flex items-center justify-between text-[12px] font-medium mb-1" style="font-family:'JetBrains Mono',monospace">
<div class="flex items-center gap-1.5">
<span class="text-[#888] text-[11px]">#{{ loop.index }}</span>
<span class="text-[#e5e2e1]">{{ name }}</span>
</div>
<span class="text-[#888]">{{ "{:,.0f}".format(score) }} sales</span>
</div>
<div class="w-full h-1 bg-[#0e0e0e] rounded-full overflow-hidden">
<div class="h-full bg-[#4d4732] rounded-full" style="width:{{ pct }}%"></div>
</div>
</div>
{% endif %}
{% endfor %}
{% if not cd.ranked_staff %}
<p class="text-[#888] text-xs text-center py-4">No data yet</p>
{% endif %}
</div>
</div>
<div class="pt-3 mt-3 border-t border-[#4d4732] flex items-center justify-between text-[11px] font-semibold tracking-wider text-[#888]" style="font-family:'JetBrains Mono',monospace">
<span>CREW SIZE: {{ cd.ranked_staff|length }} ACTIVE</span>
<span class="text-[#FFD700]">QUOTA SURPASSED</span>
</div>
</section>
</div>
<!-- METRIC CARDS -->
<section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
<div class="bg-[#20201f] border border-[#4d4732] rounded-xl p-4 hover:border-[#FFD700] transition-all group relative overflow-hidden">
<div class="flex items-center justify-between mb-2">
<span class="text-[11px] font-semibold tracking-wider text-[#888]" style="font-family:'JetBrains Mono',monospace">METRIC // TOTAL SALES</span>
<span class="material-symbols-outlined text-[#888] group-hover:text-[#FFD700] transition-colors">show_chart</span>
</div>
<div class="text-[24px] font-bold tracking-tight mb-1 text-[#FFD700]" style="font-family:'JetBrains Mono',monospace">{{ "{:,}".format(cd.total_sales|int) }} <span class="text-[11px] font-semibold tracking-wider text-[#888] font-normal">UNITS</span></div>
<div class="flex items-center gap-1.5 text-[12px] font-medium text-[#39FF14]" style="font-family:'JetBrains Mono',monospace">
<span class="material-symbols-outlined text-sm">trending_up</span>
<span>{% if cd.pct_change >= 0 %}+{% endif %}{{ "%+.1f"|format(cd.pct_change) }}% vs {% if view_mode == 'month' %}last month{% else %}last year{% endif %}</span>
</div>
</div>
<div class="bg-[#20201f] border border-[#4d4732] rounded-xl p-4 hover:border-[#FFD700] transition-all group relative overflow-hidden">
<div class="flex items-center justify-between mb-2">
<span class="text-[11px] font-semibold tracking-wider text-[#888]" style="font-family:'JetBrains Mono',monospace">LEADING OPERATOR</span>
<span class="material-symbols-outlined text-[#FFD700]">workspace_premium</span>
</div>
<div class="text-[24px] font-bold tracking-tight mb-1 truncate text-[#FFD700]" style="font-family:'JetBrains Mono',monospace">{{ cd.top_name or "—" }}</div>
<div class="flex items-center justify-between text-[12px] font-medium" style="font-family:'JetBrains Mono',monospace">
<span class="text-[#FFD700] font-semibold">{{ "{:,.0f}".format(cd.top_score|int) }} Sales</span>
{% if cd.ranked_staff %}<span class="px-2 py-0.5 bg-[#FFD700]/10 border border-[#FFD700]/30 text-[#FFD700] text-[11px] font-semibold tracking-wider rounded">TOP BADGE</span>{% endif %}
</div>
</div>
<div class="bg-[#20201f] border border-[#4d4732] rounded-xl p-4 hover:border-[#FFD700] transition-all group relative overflow-hidden">
<div class="flex items-center justify-between mb-2">
<span class="text-[11px] font-semibold tracking-wider text-[#888]" style="font-family:'JetBrains Mono',monospace">DAILY RUN RATE</span>
<span class="material-symbols-outlined text-[#888] group-hover:text-[#FFD700] transition-colors">speed</span>
</div>
<div class="text-[24px] font-bold tracking-tight mb-1 text-[#FFD700]" style="font-family:'JetBrains Mono',monospace">{{ "{:,}".format(cd.avg_daily|int) }} <span class="text-[11px] font-semibold tracking-wider text-[#888] font-normal">UNITS/DAY</span></div>
<div class="flex items-center gap-1.5 text-[12px] font-medium text-[#39FF14]" style="font-family:'JetBrains Mono',monospace">
<span class="material-symbols-outlined text-sm">bolt</span>
<span>{{ active_days }} active days tracked</span>
</div>
</div>
<div class="bg-[#20201f] border border-[#4d4732] rounded-xl p-4 hover:border-[#FFD700] transition-all group relative overflow-hidden">
<div class="flex items-center justify-between mb-2">
<span class="text-[11px] font-semibold tracking-wider text-[#888]" style="font-family:'JetBrains Mono',monospace">ACTIVE TELEMETRY TASKS</span>
<span class="material-symbols-outlined text-[#888] group-hover:text-[#FFD700] transition-colors">checklist</span>
</div>
<div class="text-[24px] font-bold tracking-tight mb-1 text-[#FFD700]" style="font-family:'JetBrains Mono',monospace">{{ cd.num_categories }} <span class="text-[11px] font-semibold tracking-wider text-[#888] font-normal">CATEGORIES</span></div>
<div class="flex items-center gap-1.5 text-[12px] font-medium text-[#39FF14]" style="font-family:'JetBrains Mono',monospace">
<span class="material-symbols-outlined text-sm">verified</span>
<span>100% Tracking Fidelity</span>
</div>
</div>
</section>
<!-- FORM MODAL -->
{{ form_modal|safe }}
<!-- HOVER TOOLTIP JS -->
<script>
var CHART_DATA={
{% for tk in all_tasks %}
"{{ tk }}":[{% for i in range(num_points) %}{% set key = (i+1) if view_mode=='month' else (i+1) %}{{ chart_points.get(key, {}).get(tk, 0) }}{% if not loop.last %},{% endif %}{% endfor %}]{% if not loop.last %},{% endif %}
{% endfor %}
};
var CHART_LABELS={{ chart_labels|tojson }};
var CHART_COLORS={lubes:"#3B82F6",combo:"#10B981",pastry:"#FFD700",shellapp:"#A855F7"};
var CHART_NAMES={% for tk in all_tasks %}"{{ tk }}":"{{ task_labels.get(tk, tk)|upper }}"{% if not loop.last %},{% endif %}{% endfor %};
(function(){
var svg=document.getElementById("chart-svg");
var tip=document.getElementById("chart-tip");
if(!svg||!tip)return;
var zones=svg.querySelectorAll(".hover-zone");
var hl=document.getElementById("hover-line");
var vb=svg.viewBox.baseVal;
zones.forEach(function(z){
z.addEventListener("mouseenter",function(){
var idx=parseInt(z.getAttribute("data-idx"));
var xPos=parseFloat(z.getAttribute("x"))+15;
hl.setAttribute("x1",xPos);hl.setAttribute("x2",xPos);hl.setAttribute("opacity","0.8");
var html="<div style='margin-bottom:4px;color:#FFD700;font-weight:700'>"+CHART_LABELS[idx]+"</div>";
var tasks=["lubes","combo","pastry","shellapp"];
var hasAny=false;
for(var t=0;t<tasks.length;t++){
var tk=tasks[t];
var v=CHART_DATA[tk]?CHART_DATA[tk][idx]:0;
if(v>0){html+="<div style='color:"+CHART_COLORS[tk]+"'>"+CHART_NAMES[tk]+": "+v+"</div>";hasAny=true;}
}
if(hasAny){tip.innerHTML=html;tip.classList.add("show");}
tip.style.left=(xPos+10)+"px";tip.style.top="10px";
});
z.addEventListener("mouseleave",function(){
hl.setAttribute("opacity","0");tip.classList.remove("show");
});
});
})();
</script>
<!-- FOOTER -->
<footer class="w-full flex flex-col items-center justify-center py-4 px-4 text-center bg-[#0e0e0e] border-t border-[#4d4732]">
<div class="flex items-center justify-center gap-6 mb-2">
<span class="text-[11px] font-semibold tracking-wider text-[#FFD700] underline cursor-pointer" style="font-family:'JetBrains Mono',monospace">SYS_STATUS: ONLINE</span>
<span class="text-[11px] font-semibold tracking-wider text-[#888] cursor-pointer" style="font-family:'JetBrains Mono',monospace">SECURITY POLICY</span>
<span class="text-[11px] font-semibold tracking-wider text-[#888] cursor-pointer" style="font-family:'JetBrains Mono',monospace">TERMINAL DOCS</span>
</div>
<p class="text-[12px] font-medium text-[#888]" style="font-family:'JetBrains Mono',monospace">&copy; 2025 PT SHELL PROSPER NIAGA. SECURE SCADA ACCESS PROTOCOL. ALL RIGHTS RESERVED.</p>
</footer>
</div>'''

# ============== LEADERBOARD ==============
LEADERBOARD_TPL = '''
<style>
.lb-print-header{display:none}
.print-rank-header{display:none}
@media print{
@page{size:A4 portrait;margin:0.15in 0.3in}
*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}
body{margin:0;padding:0;background:#fff!important}
.print-hide,.mobile-header,.sidebar-wrapper,#sidebar-overlay,#toast-container,aside,nav,header{display:none!important}
main{margin:0!important;padding:0!important;overflow:visible!important}
main>div:first-child{display:none!important}
.lb-print-header{display:block!important;margin-bottom:3px;font-family:Arial,sans-serif}
.lb-print-header *{all:unset;display:block}
.lb-print-header .ph-banner{background:#000;color:#F1C232;text-align:center;padding:7px 0;font-weight:700;font-size:24px;letter-spacing:6px;font-family:Arial,sans-serif}
.lb-print-header .ph-title{color:#1F4E78;font-size:20px;font-weight:800;text-align:center;letter-spacing:1px;margin-top:4px;font-family:Arial,sans-serif}
.lb-print-header .ph-star{background:#000;color:#FF0000;text-align:center;padding:5px 0;font-size:16px;letter-spacing:2px;font-weight:700;margin-top:3px;font-family:Arial,sans-serif}
.lb-print-header .ph-date{color:#1F4E78;font-size:22px;text-align:center;margin-top:8px;font-weight:700;font-style:italic;font-family:Arial,sans-serif}
.lb-print-header hr{border:none;border-top:1.5px solid #000;margin:3px 0 0 0}
.podium-wrap{display:flex!important;justify-content:center!important;align-items:flex-end!important;gap:4px!important;padding:0 4px!important}
.podium-card,.podium-card[class*="bg-dark-card"]{background:#fff!important;border:1px solid #999!important;border-radius:20px!important;padding:3px 2px!important;box-shadow:none!important;animation:none!important;overflow:visible!important;display:flex!important;flex-direction:column!important;align-items:center!important;justify-content:flex-end!important}
.podium-card [class*="bg-gradient"]{display:none!important}
.podium-card [class*="absolute"]{display:none!important}
.podium-card .rounded-full,.podium-card [class*="rounded-full"]{border-radius:50%!important;overflow:hidden!important;display:flex!important;align-items:center!important;justify-content:center!important}
.podium-card [class*="w-18"],.podium-card [class*="h-18"]{width:90px!important;height:90px!important}
.podium-card [class*="w-24"],.podium-card [class*="h-24"]{width:90px!important;height:90px!important}
.podium-card [class*="w-32"],.podium-card [class*="h-32"]{width:90px!important;height:90px!important}
.podium-card img{width:100%!important;height:100%!important;object-fit:cover!important;display:block!important;margin:0!important;padding:0!important;border:none!important}
.podium-card .pod-1st-name,.podium-card .pod-2nd-name,.podium-card .pod-3rd-name{font-size:16px!important;font-weight:900!important;color:#000!important;font-family:Arial,sans-serif!important;padding:1px 6px!important;display:inline!important;border-radius:2px!important}
.podium-card .pod-1st-name{background:#FFD700!important}
.podium-card .pod-2nd-name{background:#C0C0C0!important}
.podium-card .pod-3rd-name{background:#CD7F32!important}
.podium-card [class*="text-sm"]{font-size:14px!important;color:#000!important}
.podium-card .pod-pct{font-size:18px!important;font-weight:900!important;color:#000!important;font-family:Arial,sans-serif!important}
.podium-card [class*="text-3xl"],.podium-card [class*="text-5xl"]{font-size:26px!important;color:#000!important}
.ppodium-card [class*="text-on-dark"],.podium-card [class*="text-gray"],.podium-card [class*="text-ny"],.podium-card [class*="text-orange"]{color:#000!important}
.podium-card [class*="text-on-dark-muted"],.podium-card [class*="text-on-dark-dim"]{font-size:14px!important;font-weight:700!important;color:#333!important}
.podium-card [class*="uppercase"][class*="tracking-widest"]{font-size:16px!important;font-weight:800!important;color:#000!important;letter-spacing:0px!important;white-space:nowrap!important}
.podium-card .pod-tasks,.podium-card [class*="pod-tasks"]{display:flex!important;flex-direction:column!important;flex-wrap:nowrap!important;gap:3px!important;justify-content:center!important;margin-top:6px!important;margin-left:-10px!important;margin-right:-10px!important;width:calc(100% + 20px)!important}
.podium-card .task-row,.podium-card [class*="task-row"]{display:block!important;width:100%!important;text-align:left!important;padding:4px 10px!important;border-radius:6px!important;font-size:14px!important;font-family:monospace!important;font-weight:700!important}
.podium-card .task-row.bg-blue-500{background:rgba(59,130,246,0.2)!important;color:#3b82f6!important;border:1px solid rgba(59,130,246,0.3)!important}
.podium-card .task-row.bg-purple-500{background:rgba(139,92,246,0.2)!important;color:#8b5cf6!important;border:1px solid rgba(139,92,246,0.3)!important}
.podium-card .task-row.bg-pink-500{background:rgba(236,72,153,0.2)!important;color:#ec4899!important;border:1px solid rgba(236,72,153,0.3)!important}
.podium-card .task-row.bg-emerald-500{background:rgba(16,185,129,0.2)!important;color:#10b981!important;border:1px solid rgba(16,185,129,0.3)!important}
.rankings-box{border:1px solid #999!important;background:#fff!important;border-radius:0!important;margin-top:4px!important;padding:0!important}
.rankings-box .p-4,.rankings-box [class*="p-6"],.rankings-box [class*="lg:p-"]{padding:0!important}
.rankings-box .border-b,.rankings-box [class*="border-b"]{border-bottom:none!important}
.rankings-box h3{display:none!important}
.rankings-box table{width:100%!important;border-collapse:collapse!important}
.rankings-box thead.print-hide{display:none!important}
.rankings-box .print-rank-header{display:table-header-group!important}
.rankings-box .print-rank-header tr{background:#e0e0e0!important}
.rankings-box .print-rank-header th{color:#000!important;font-weight:700!important;border-bottom:1.5px solid #000!important;font-size:13px!important;padding:5px 10px!important;font-family:Arial,sans-serif!important;text-transform:uppercase!important;text-align:left!important}
.rankings-box .print-rank-header th:last-child{text-align:right!important}
.rankings-box tbody tr.print-rank-top3{display:none!important}
.rankings-box tbody tr{display:table-row!important}
.rankings-box tbody td{color:#000!important;font-size:13px!important;padding:5px 10px!important;border-bottom:0.5px solid #ddd!important;font-family:Arial,sans-serif!important;text-align:left!important}
.rankings-box tbody td:last-child{text-align:right!important}
.rankings-box tbody td:last-child span{font-size:16px!important;font-weight:900!important;color:#000!important}
.rankings-box tbody tr:nth-child(even) td{background:#f8f8f8!important}
.rankings-box tbody tr:nth-child(odd) td{background:#fff!important}
.rankings-box [class*="text-ny"],.rankings-box [class*="text-gray"],.rankings-box [class*="text-on-dark"],.rankings-box [class*="text-orange"]{color:#000!important}
.rankings-box [class*="hover"]{background:transparent!important}
.rankings-box [class*="font-mono"]{font-family:Arial,sans-serif!important}
.rankings-box [class*="font-black"],.rankings-box [class*="font-bold"],.rankings-box [class*="font-semibold"]{font-weight:700!important}
.rankings-box [class*="text-base"],.rankings-box [class*="text-lg"]{font-size:9px!important}
.rankings-box .table-scroll{overflow:visible!important}
.max-w-6xl{max-width:100%!important}
}
</style>
{% if is_admin %}
<!-- PRINT HEADER (hidden on screen, shown on print) -->
<div class="lb-print-header">
<div class="ph-banner">PROSPER NIAGA</div>
<div class="ph-title">LEADERBOARD</div>
<div class="ph-star">✦&nbsp;&nbsp;SHINING STAR OF THE MONTH&nbsp;&nbsp;✦</div>
<div class="ph-date">{{ current_month }}</div>
<hr>
</div>
{% endif %}
<!-- SCREEN HEADER -->
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border print-hide">
<div class="max-w-7xl mx-auto">
<div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 sm:gap-4">
<div>
<h2 class="text-xl sm:text-2xl lg:text-3xl font-black text-ny">SITE HERO SALES LEADERBOARD</h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">Shell Bandar Mahkota Cheras - Prosper Niaga | {{ current_month }}</p>
</div>
<form method="GET" class="flex items-center gap-2 print-hide">
<select name="month" class="bg-dark border border-dark-border rounded-lg px-3 py-2 text-on-dark text-sm font-mono">
{% for m in range(1,13) %}
<option value="{{ m }}" {% if m==sel_month %}selected{% endif %}>{{ datetime(sel_year, m, 1).strftime('%B') }}</option>
{% endfor %}
</select>
<select name="year" class="bg-dark border border-dark-border rounded-lg px-3 py-2 text-on-dark text-sm font-mono">
{% for y in available_years %}
<option value="{{ y }}" {% if y==sel_year %}selected{% endif %}>{{ y }}</option>
{% endfor %}
</select>
<button type="submit" class="bg-ny text-dark px-4 py-2 rounded-lg font-bold text-sm hover:bg-ny-bright transition">Go</button>
</form>
</div>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-6xl mx-auto">
<!-- Top 3 Podium: 2nd(left) - 1st(center,tallest) - 3rd(right) -->
{% if podium|length > 0 %}
<div class="mb-8 lg:mb-12">
<div class="podium-wrap flex items-end justify-center gap-2 sm:gap-3 lg:gap-5 px-2">
<!-- 2nd Place (Left) -->
{% if podium|length >= 2 %}
<div class="flex-1 max-w-[200px] sm:max-w-[280px] animate-fadeInUp" style="animation-delay:0.1s">
<div class="podium-card bg-dark-card border border-gray-400/30 rounded-[20px] p-5 sm:p-8 pb-8 sm:pb-12 text-center relative overflow-hidden">
<div class="absolute inset-0 bg-gradient-to-b from-gray-400/5 to-transparent"></div>
<div class="relative">
<div class="w-28 h-28 sm:w-36 sm:h-36 rounded-full mx-auto mb-4 flex items-center justify-center overflow-hidden border-2 border-gray-400/40">
{% if podium[1].picture %}<img src="/static/uploads/{{ podium[1].picture }}" class="w-full h-full object-cover" alt="{{ podium[1].name }}">{% else %}<span class="text-xl sm:text-3xl font-black text-gray-300">{{ podium[1].name[0] }}</span>{% endif %}
</div>
<div class="text-[9px] sm:text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1">2ND PLACE</div>
<div class="pod-2nd-name text-sm sm:text-lg font-black text-white mb-1">{{ podium[1].name }}</div>
<div class="pod-pct text-base sm:text-2xl font-black font-mono text-gray-300 mb-1">{{ podium[1].pct }}%</div>
<div class="text-[9px] sm:text-[10px] text-on-dark-dim mb-3">{{ podium[1].score|round(0)|int }} sales</div>
<div class="pod-tasks flex flex-col gap-1.5 justify-center w-full">
{% for tk in podium[1].task_details %}
<span class="task-row text-[10px] sm:text-[11px] px-3 py-1.5 rounded-lg font-mono font-bold text-left w-full
{% if tk.key == 'lubes' %}bg-blue-500/20 text-blue-400 border border-blue-500/30
{% elif tk.key == 'combo' %}bg-purple-500/20 text-purple-400 border border-purple-500/30
{% elif tk.key == 'pastry' %}bg-pink-500/20 text-pink-400 border border-pink-500/30
{% elif tk.key == 'shellapp' %}bg-emerald-500/20 text-emerald-400 border border-emerald-500/30
{% else %}bg-gray-500/20 text-gray-400 border border-gray-500/30{% endif %}">
{{ tk.label }}: {{ tk.value|int }}
</span>
{% endfor %}
</div>
</div>
</div>
</div>
{% endif %}
<!-- 1st Place (Center, Tallest) -->
{% if podium|length >= 1 %}
<div class="flex-1 max-w-[200px] sm:max-w-[280px] animate-fadeInUp" style="animation-delay:0s">
<div class="podium-card bg-dark-card border-2 border-ny/40 rounded-[20px] p-5 sm:p-8 pb-8 sm:pb-12 text-center glow-strong animate-pulse-glow relative overflow-hidden">
<div class="absolute inset-0 bg-gradient-to-b from-ny/10 to-transparent"></div>
<div class="relative">
<div class="text-3xl sm:text-5xl mb-3 sm:mb-4">&#x1F3C6;</div>
<div class="w-28 h-28 sm:w-36 sm:h-36 rounded-full mx-auto mb-4 sm:mb-5 flex items-center justify-center overflow-hidden border-3 border-ny/50 shadow-lg shadow-ny/20">
{% if podium[0].picture %}<img src="/static/uploads/{{ podium[0].picture }}" class="w-full h-full object-cover" alt="{{ podium[0].name }}">{% else %}<span class="text-3xl sm:text-5xl font-black text-ny">{{ podium[0].name[0] }}</span>{% endif %}
</div>
<div class="text-[10px] sm:text-xs font-bold text-ny uppercase tracking-widest mb-1">&#x2B50; SHINING STAR &#x2B50;</div>
<div class="pod-1st-name text-lg sm:text-2xl font-black text-white mb-1">{{ podium[0].name }}</div>
<div class="pod-pct text-2xl sm:text-4xl font-black font-mono text-ny mb-1">{{ podium[0].pct }}%</div>
<div class="text-[10px] sm:text-xs text-on-dark-muted mb-2 sm:mb-3">{{ podium[0].score|round(0)|int }} sales</div>
<div class="pod-tasks flex flex-col gap-1.5 sm:gap-2 justify-center w-full">
{% for tk in podium[0].task_details %}
<span class="task-row text-[10px] sm:text-[11px] px-3 py-1.5 rounded-lg font-mono font-bold text-left w-full
{% if tk.key == 'lubes' %}bg-blue-500/20 text-blue-400 border border-blue-500/30
{% elif tk.key == 'combo' %}bg-purple-500/20 text-purple-400 border border-purple-500/30
{% elif tk.key == 'pastry' %}bg-pink-500/20 text-pink-400 border border-pink-500/30
{% elif tk.key == 'shellapp' %}bg-emerald-500/20 text-emerald-400 border border-emerald-500/30
{% else %}bg-gray-500/20 text-gray-400 border border-gray-500/30{% endif %}">
{{ tk.label }}: {{ tk.value|int }}
</span>
{% endfor %}
</div>
</div>
</div>
</div>
{% endif %}
<!-- 3rd Place (Right) -->
{% if podium|length >= 3 %}
<div class="flex-1 max-w-[200px] sm:max-w-[280px] animate-fadeInUp" style="animation-delay:0.2s">
<div class="podium-card bg-dark-card border border-orange-400/30 rounded-[20px] p-5 sm:p-8 pb-8 sm:pb-12 text-center relative overflow-hidden">
<div class="absolute inset-0 bg-gradient-to-b from-orange-400/5 to-transparent"></div>
<div class="relative">
<div class="w-28 h-28 sm:w-36 sm:h-36 rounded-full mx-auto mb-4 flex items-center justify-center overflow-hidden border-2 border-orange-400/40">
{% if podium[2].picture %}<img src="/static/uploads/{{ podium[2].picture }}" class="w-full h-full object-cover" alt="{{ podium[2].name }}">{% else %}<span class="text-xl sm:text-3xl font-black text-orange-300">{{ podium[2].name[0] }}</span>{% endif %}
</div>
<div class="text-[9px] sm:text-[10px] font-bold text-orange-400 uppercase tracking-widest mb-1">3RD PLACE</div>
<div class="pod-3rd-name text-sm sm:text-lg font-black text-white mb-1">{{ podium[2].name }}</div>
<div class="pod-pct text-base sm:text-2xl font-black font-mono text-orange-300 mb-1">{{ podium[2].pct }}%</div>
<div class="text-[9px] sm:text-[10px] text-on-dark-dim mb-3">{{ podium[2].score|round(0)|int }} sales</div>
<div class="pod-tasks flex flex-col gap-1.5 justify-center w-full">
{% for tk in podium[2].task_details %}
<span class="task-row text-[10px] sm:text-[11px] px-3 py-1.5 rounded-lg font-mono font-bold text-left w-full
{% if tk.key == 'lubes' %}bg-blue-500/20 text-blue-400 border border-blue-500/30
{% elif tk.key == 'combo' %}bg-purple-500/20 text-purple-400 border border-purple-500/30
{% elif tk.key == 'pastry' %}bg-pink-500/20 text-pink-400 border border-pink-500/30
{% elif tk.key == 'shellapp' %}bg-emerald-500/20 text-emerald-400 border border-emerald-500/30
{% else %}bg-gray-500/20 text-gray-400 border border-gray-500/30{% endif %}">
{{ tk.label }}: {{ tk.value|int }}
</span>
{% endfor %}
</div>
</div>
</div>
</div>
{% endif %}
</div>
</div>
{% endif %}
<!-- Full Rankings Table: Rank, Name (left), Score% (right) -->
<div class="rankings-box bg-dark-card border border-dark-border rounded-xl overflow-hidden">
<div class="p-4 lg:p-6 border-b border-dark-border print-hide">
<h3 class="text-base sm:text-lg font-bold text-on-dark">FULL RANKINGS</h3>
</div>
<div class="table-scroll">
<table class="w-full text-xs sm:text-sm">
<thead class="print-hide"><tr class="border-b border-dark-border">
<th class="py-3 px-4 sm:px-6 text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase text-left w-16">Rank</th>
<th class="py-3 px-4 sm:px-6 text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase text-left">Staff</th>
<th class="py-3 px-4 sm:px-6 text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase text-right">Score</th>
</tr></thead>
<thead class="print-rank-header"><tr class="border-b border-dark-border">
<th class="py-2 px-4 text-xs font-bold uppercase text-left w-12" style="color:#000;border-bottom:2px solid #000">Rank</th>
<th class="py-2 px-4 text-xs font-bold uppercase text-left" style="color:#000;border-bottom:2px solid #000">Site Hero</th>
<th class="py-2 px-4 text-xs font-bold uppercase text-right" style="color:#000;border-bottom:2px solid #000">Total Score</th>
</tr></thead>
<tbody class="divide-y divide-dark-border">
{% for r in rankings %}
<tr class="hover:bg-dark-hover transition {% if r.rank<=3 %}print-rank-top3{% endif %}">
<td class="py-3 px-4 sm:px-6 font-mono font-bold text-xs sm:text-sm
{% if r.rank==1 %}text-ny rank-1{% elif r.rank==2 %}text-gray-300 rank-2{% elif r.rank==3 %}text-orange-300 rank-3{% else %}text-on-dark-dim{% endif %}">
{% if r.rank<=3 %}{{ ['🥇','🥈','🥉'][r.rank-1] }}{% endif %} #{{ r.rank }}</td>
<td class="py-3 px-4 sm:px-6">
<span class="font-semibold text-on-dark text-xs sm:text-sm">{{ r.name }}</span>
</td>
<td class="py-3 px-4 sm:px-6 text-right">
<span class="text-sm sm:text-base font-black font-mono
{% if r.rank==1 %}text-ny{% elif r.rank<=3 %}text-on-dark{% else %}text-on-dark-muted{% endif %}">
{{ r.pct }}%</span>
</td>
</tr>
{% endfor %}
</tbody></table></div></div>
{% if is_admin %}
<div class="mt-4 text-center print-hide">
<button onclick="window.print()" class="bg-ny text-dark px-6 py-3 rounded-lg font-bold text-sm hover:bg-ny-bright transition inline-flex items-center gap-2">
<span class="material-symbols-outlined text-base">print</span>Print A4
</button>
</div>
{% endif %}
</div>'''

# ============== DAILY TRACKING ==============
DAILY_TPL = '''
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-7xl mx-auto">
<div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 sm:gap-4">
<div>
<h2 class="text-xl sm:text-2xl lg:text-3xl font-black text-ny">DAILY TRACKING</h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">OneTrack sales grid - {{ current_month }}</p>
</div>
<!-- Month/Year Selector -->
<form method="GET" class="flex items-center gap-2">
<select name="month" class="bg-dark border border-dark-border rounded-lg px-3 py-2 text-on-dark text-sm font-mono">
{% for m in range(1,13) %}
<option value="{{ m }}" {% if m==sel_month %}selected{% endif %}>{{ datetime(sel_year, m, 1).strftime('%B') }}</option>
{% endfor %}
</select>
<select name="year" class="bg-dark border border-dark-border rounded-lg px-3 py-2 text-on-dark text-sm font-mono">
{% for y in available_years %}
<option value="{{ y }}" {% if y==sel_year %}selected{% endif %}>{{ y }}</option>
{% endfor %}
</select>
<button type="submit" class="bg-ny text-dark px-4 py-2 rounded-lg font-bold text-sm hover:bg-ny-bright transition">Go</button>
</form>
</div>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-7xl mx-auto">
<!-- Monthly Sales Summary Cards -->
<div class="mb-6">
<h3 class="text-xs font-semibold text-on-dark-dim uppercase tracking-wider mb-3">Monthly Sales Summary</h3>
<div class="grid grid-cols-2 lg:grid-cols-4 gap-3">
{% for tk in all_tasks %}
<div class="bg-dark-card border border-dark-border rounded-xl p-3 sm:p-4">
<div class="text-[10px] sm:text-xs text-on-dark-muted uppercase tracking-wider mb-1">{{ task_labels[tk] }} ({{ task_units[tk] }})</div>
<div class="text-xl sm:text-2xl font-black font-mono text-ny">{{ task_grand_totals[tk] }}</div>
</div>
{% endfor %}
</div>
</div>
<!-- Per-Staff Monthly Totals -->
<div class="mb-6">
<h3 class="text-xs font-semibold text-on-dark-dim uppercase tracking-wider mb-3">Per Staff Monthly Totals</h3>
<div class="table-scroll">
<table class="w-full text-xs sm:text-sm">
<thead><tr class="border-b border-dark-border">
<th class="py-2 px-2 sm:px-3 text-[10px] sm:text-xs text-on-dark-dim uppercase text-left">Staff</th>
{% for tk in all_tasks %}
<th class="py-2 px-2 sm:px-3 text-[10px] sm:text-xs text-on-dark-dim uppercase text-center">{{ tk[:3]|upper }}</th>
{% endfor %}
<th class="py-2 px-2 sm:px-3 text-[10px] sm:text-xs text-on-dark-dim uppercase text-center font-bold">Total</th>
</tr></thead>
<tbody class="divide-y divide-dark-border">
{% for s in staff %}
{% set mt = monthly_totals.get(s, {}) %}
<tr class="hover:bg-dark-hover transition">
<td class="py-2 px-2 sm:px-3 font-semibold text-on-dark text-xs sm:text-sm">{{ s }}</td>
{% for tk in all_tasks %}
{% set val = mt.get(tk, 0)|int %}
{% set tgt = targets.get(s, {}).get(tk, 0) %}
<td class="py-2 px-2 sm:px-3 text-center font-mono text-xs sm:text-sm
{% if tgt>0 and val>=tgt %}text-green-neon{% elif val>0 %}text-ny{% else %}text-on-dark-dim{% endif %}">
{{ val }}</td>
{% endfor %}
<td class="py-2 px-2 sm:px-3 text-center font-mono font-bold text-on-dark text-xs sm:text-sm">{{ mt.get('_total', 0)|int }}</td>
</tr>
{% endfor %}
</tbody></table></div>
</div>
<!-- Daily Grid -->
<div class="mb-3">
<h3 class="text-xs font-semibold text-on-dark-dim uppercase tracking-wider">Daily Grid</h3>
</div>
<div class="table-scroll">
<table class="w-full text-xs sm:text-sm border-collapse">
<thead><tr class="border-b-2 border-ny/30">
<th class="py-2 px-1 sm:px-2 text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase sticky left-0 bg-dark-card z-10 min-w-[40px] border-r border-dark-border">Day</th>
{% set ns = namespace(staff_idx=0) %}
{% for s in staff %}
{% set staff_idx = loop.index0 %}
<th colspan="{{ staff_tasks.get(s, [])|length }}" class="py-2 px-1 sm:px-2 text-[10px] sm:text-xs font-semibold text-center border-l-2 border-ny/20
{% if s in top3 %}text-ny bg-ny/5{% else %}text-on-dark-muted{% if staff_idx % 2 == 0 %} bg-dark-card{% else %} bg-dark-card{% endif %}{% endif %} min-w-[60px]">
{{ s }}{% if s in top3 %} ⭐{% endif %}</th>
{% endfor %}
</tr>
<tr class="border-b border-dark-border">
<th class="py-1 px-1 sm:px-2 text-[10px] sm:text-xs text-on-dark-dim sticky left-0 bg-dark-card z-10 border-r border-dark-border"></th>
{% for s in staff %}
{% for tk in staff_tasks.get(s, []) %}
<th class="py-1 px-1 sm:px-2 text-[10px] sm:text-xs text-on-dark-dim text-center border-l {% if loop.first %}border-l-2 border-ny/20{% else %}border-dark-border{% endif %} font-mono min-w-[40px]">{{ tk[:3]|upper }}</th>
{% endfor %}
{% endfor %}
</tr></thead>
<tbody>
{% for day in range(1, days_in_month+1) %}
<tr id="day-{{ day }}" class="border-b border-dark-border hover:bg-dark-hover transition">
<td class="py-1.5 sm:py-2 px-1 sm:px-2 font-mono font-bold text-on-dark sticky left-0 bg-dark-card text-center z-10 min-w-[40px] border-r border-dark-border
{% if day==today and sel_month==today_month and sel_year==today_year %}text-ny bg-ny/5{% endif %}">{{ day }}</td>
{% for s in staff %}
{% for tk in staff_tasks.get(s, []) %}
{% set val = daily.get(day,{}).get(s,{}).get(tk,0)|int %}
{% set tgt = daily_targets.get(s, {}).get(tk,1) %}
<td class="py-1.5 sm:py-2 px-1 sm:px-2 text-center border-l {% if loop.first %}border-l-2 border-ny/20{% else %}border-dark-border{% endif %} font-mono text-xs sm:text-sm
{% if val > 0 and val >= tgt %}text-green-neon bg-green-neon/5
{% elif val > 0 and val < tgt %}text-red bg-red/5
{% else %}text-on-dark-dim{% endif %}">
{{ val }}</td>
{% endfor %}
{% endfor %}
</tr>
{% endfor %}
</tbody></table></div>
<!-- Legend -->
<div class="mt-4 sm:mt-6 flex flex-wrap gap-3 sm:gap-4 text-[10px] sm:text-xs text-on-dark-muted">
<div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 sm:w-3 sm:h-3 rounded bg-green-neon/30"></span>Target Met (≥ target)</div>
<div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 sm:w-3 sm:h-3 rounded bg-red/30"></span>Below Target</div>
<div class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 sm:w-3 sm:h-3 rounded bg-dark-border"></span>No Data (0)</div>
</div></div>'''

# ============== FORM ==============
FORM_TPL = '''
<!-- Modal Backdrop -->
<div id="entryModal" class="fixed inset-0 z-[100] flex items-center justify-center p-4" style="display:none;">
<div class="absolute inset-0 bg-black/70" onclick="closeModal()"></div>
<div class="relative bg-dark-card border border-dark-border rounded-lg w-full max-w-lg max-h-[90vh] overflow-y-auto">
<!-- Modal Header -->
<div class="flex items-center justify-between px-4 py-3 border-b border-dark-border sticky top-0 bg-dark-card z-10">
<h3 class="text-white font-bold text-sm">New Entry</h3>
<button onclick="closeModal()" class="text-on-dark-dim hover:text-white text-lg">&times;</button>
</div>
<!-- Modal Body -->
<div class="p-4">
<form method="POST" id="entryForm" action="/form">
<div class="mb-4">
<label class="block text-[10px] text-on-dark-dim uppercase tracking-wider mb-1">Date</label>
<input type="date" name="date" value="{{ today }}" required
class="w-full bg-dark border border-dark-border rounded px-3 py-2 text-white font-mono text-sm">
</div>
<!-- Staff Quick Select -->
<div class="mb-4">
<label class="block text-[10px] text-on-dark-dim uppercase tracking-wider mb-2">Select Staff</label>
<div class="flex flex-wrap gap-1.5" id="staffChips">
{% for s in staff %}
<button type="button" onclick="toggleStaff('{{ s }}')"
class="staff-chip px-2.5 py-1 rounded text-xs font-semibold border border-dark-border text-on-dark-muted transition"
data-staff="{{ s }}">{{ s }}</button>
{% endfor %}
</div>
<div id="selectedStaff" class="text-[10px] text-on-dark-dim mt-1"></div>
</div>
<!-- Dynamic Task Inputs -->
<div id="taskInputs" class="space-y-3 mb-4"></div>
<!-- Submit -->
<div class="flex gap-2">
<button type="submit" class="flex-1 bg-ny text-dark py-2.5 rounded font-bold text-sm">Submit</button>
<button type="button" onclick="closeModal()" class="px-4 py-2.5 rounded text-on-dark-muted text-sm border border-dark-border">Cancel</button>
</div>
</form>
</div>
</div>
</div>
<!-- Trigger Button -->
<button onclick="openModal()" class="bg-ny text-dark px-4 py-2 rounded font-bold text-sm hover:bg-ny-bright transition">
+ New Entry
</button>
<script>
const STAFF_TASKS = {{ staff_tasks_json|safe }};
const ALL_TASKS = {{ task_keys|tojson }};
const TASK_LABELS = {{ task_labels|tojson }};
const TASK_UNITS = {{ task_units|tojson }};
let selected = new Set();
function openModal(){
document.getElementById('entryModal').style.display='flex';
selected.clear();renderChips();renderInputs();
}
function closeModal(){
document.getElementById('entryModal').style.display='none';
}
function toggleStaff(name){
if(selected.has(name)){selected.delete(name)}else{selected.add(name)}
renderChips();renderInputs();
}
function renderChips(){
document.querySelectorAll('.staff-chip').forEach(c=>{
const isActive=selected.has(c.dataset.staff);
c.classList.toggle('bg-ny/20',isActive);
c.classList.toggle('border-ny',isActive);
c.classList.toggle('text-dark',isActive);
c.classList.toggle('text-on-dark-muted',!isActive);
});
document.getElementById('selectedStaff').textContent=selected.size+' selected';
}
function renderInputs(){
const container=document.getElementById('taskInputs');
container.innerHTML='';
selected.forEach(s=>{
const tasks=STAFF_TASKS[s]||[];
let html='<div class="border border-dark-border rounded p-3"><div class="flex items-center gap-2 mb-2">';
html+='<div class="w-6 h-6 rounded bg-ny/20 flex items-center justify-center text-ny text-xs font-bold">'+s[0]+'</div>';
html+='<span class="text-white text-sm font-semibold">'+s+'</span></div><div class="grid grid-cols-2 gap-2">';
tasks.forEach(tk=>{
const lbl=TASK_LABELS[tk]||tk.toUpperCase();
const unit=TASK_UNITS[tk]||'';
html+='<div><label class="text-[10px] text-on-dark-dim block mb-0.5">'+lbl+'</label>';
html+='<input type="number" name="'+s.toLowerCase()+'_'+tk+'" min="0" value="0" class="w-full bg-dark border border-dark-border rounded px-2 py-1.5 text-white font-mono text-sm text-center">';
html+='</div>';
});
html+='</div></div>';
container.innerHTML+=html;
});
}
</script>'''

FORM_PAGE_TPL = '''
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-3xl mx-auto">
<h2 class="text-xl sm:text-2xl lg:text-3xl font-black text-ny">DATA ENTRY</h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">Submit daily sales data for each staff member</p>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-3xl mx-auto">
{% with messages = get_flashed_messages(with_categories=true) %}
{% if messages %}
{% for cat, msg in messages %}
<div class="px-4 py-3 rounded text-sm mb-4 {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}
{% endif %}
{% endwith %}
<div class="bg-dark-card border border-dark-border rounded-xl p-4 sm:p-6">
<form method="POST" id="entryForm" action="/form">
<div class="mb-4">
<label class="block text-[10px] text-on-dark-dim uppercase tracking-wider mb-1">Date</label>
<input type="date" name="date" value="{{ today }}" required
class="w-full bg-dark border border-dark-border rounded px-3 py-3 text-white font-mono text-base" style="min-height:44px">
</div>
<div class="mb-4">
<label class="block text-[10px] text-on-dark-dim uppercase tracking-wider mb-2">Select Staff</label>
<div class="flex flex-wrap gap-2" id="staffChips">
{% for s in staff %}
<button type="button" onclick="toggleStaff('{{ s }}')"
class="staff-chip px-3 py-2 rounded-lg text-sm font-semibold border border-dark-border text-on-dark-muted transition"
style="min-height:44px;min-width:44px"
data-staff="{{ s }}">{{ s }}</button>
{% endfor %}
</div>
<div id="selectedStaff" class="text-[10px] text-on-dark-dim mt-1"></div>
</div>
<div id="taskInputs" class="space-y-3 mb-4"></div>
<div class="flex gap-2">
<button type="submit" class="flex-1 bg-ny text-dark py-3 rounded-lg font-bold text-base" style="min-height:48px">Submit</button>
<a href="/" class="px-4 py-3 rounded-lg text-on-dark-muted text-sm border border-dark-border" style="min-height:48px">Cancel</a>
</div>
</form>
</div>
</div>
<script>
const STAFF_TASKS = {{ staff_tasks_json|safe }};
const ALL_TASKS = {{ task_keys|tojson }};
const TASK_LABELS = {{ task_labels|tojson }};
const TASK_UNITS = {{ task_units|tojson }};
let selected = new Set();
function toggleStaff(name){
if(selected.has(name)){selected.delete(name)}else{selected.add(name)}
renderChips();renderInputs();
}
function renderChips(){
document.querySelectorAll('.staff-chip').forEach(c=>{
const isActive=selected.has(c.dataset.staff);
c.classList.toggle('bg-ny/20',isActive);
c.classList.toggle('border-ny',isActive);
c.classList.toggle('text-dark',isActive);
c.classList.toggle('text-on-dark-muted',!isActive);
});
document.getElementById('selectedStaff').textContent=selected.size+' selected';
}
function renderInputs(){
const container=document.getElementById('taskInputs');
container.innerHTML='';
selected.forEach(s=>{
const tasks=STAFF_TASKS[s]||[];
let html='<div class="border border-dark-border rounded-lg p-3"><div class="flex items-center gap-2 mb-2">';
html+='<div class="w-8 h-8 rounded-full bg-ny/20 flex items-center justify-center text-ny text-sm font-bold">'+s[0]+'</div>';
html+='<span class="text-white text-base font-semibold">'+s+'</span></div><div class="grid grid-cols-2 gap-3">';
tasks.forEach(tk=>{
const lbl=TASK_LABELS[tk]||tk.toUpperCase();
const unit=TASK_UNITS[tk]||'';
html+='<div><label class="text-xs text-on-dark-dim block mb-1">'+lbl+(unit?' ('+unit+')':'')+'</label>';
html+='<input type="number" inputmode="numeric" pattern="[0-9]*" name="'+s.toLowerCase()+'_'+tk+'" min="0" value="0" class="w-full bg-dark border border-dark-border rounded-lg px-3 py-2.5 text-white font-mono text-base text-center" style="min-height:44px">';
html+='</div>';
});
html+='</div></div>';
container.innerHTML+=html;
});
}
</script>'''

# ============== TEAM MEMBERS ==============
TEAM_TPL = '''
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-5xl mx-auto flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 sm:gap-4">
<div>
<h2 class="text-xl sm:text-2xl lg:text-3xl font-black text-ny">TEAM MEMBERS</h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">Manage staff profiles and targets</p>
</div>
<a href="/team/add" class="bg-ny text-dark px-3 sm:px-4 py-2 rounded-lg font-bold hover:bg-ny-bright transition flex items-center gap-2 text-xs sm:text-sm">
<span class="material-symbols-outlined text-base sm:text-lg">person_add</span>Add Member</a>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-5xl mx-auto">
<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-4">
{% for m in members %}
<div class="bg-dark-card border border-dark-border rounded-xl p-4 sm:p-5 card-hover">
<div class="flex items-center gap-3 sm:gap-4 mb-3 sm:mb-4">
<div class="w-11 h-11 sm:w-14 sm:h-14 rounded-full bg-ny/20 flex items-center justify-center text-ny text-lg sm:text-xl font-black overflow-hidden border border-ny/30">
{% if m.picture %}<img src="/static/uploads/{{ m.picture }}" class="w-full h-full object-cover">{% else %}{{ m.name[0] }}{% endif %}</div>
<div>
<h4 class="font-bold text-on-dark text-base sm:text-lg">{{ m.name }}</h4>
<p class="text-[10px] sm:text-xs text-on-dark-dim">{{ m.staff_id }} | {{ m.gender }}</p>
</div>
</div>
<div class="space-y-1.5 sm:space-y-2 mb-3 sm:mb-4">
{% for tk, tv in (m.targets_json if m.targets_json else {}).items() %}
<div class="flex justify-between items-center text-xs sm:text-sm">
<span class="text-on-dark-muted">{{ tk|upper }}</span>
<span class="font-mono font-bold text-ny">{{ tv }}/day</span>
</div>
{% endfor %}
{% if not m.targets_json %}
<p class="text-on-dark-dim text-xs italic">No targets set</p>
{% endif %}
</div>
<div class="flex gap-2">
<a href="/team/edit/{{ m.staff_id }}" class="flex-1 bg-dark-border text-on-dark py-2 rounded-lg text-center text-xs sm:text-sm font-semibold hover:bg-dark-hover transition">Edit</a>
<form method="POST" action="/team/delete/{{ m.staff_id }}" style="display:inline" onsubmit="event.preventDefault();var form=this;showDeleteModal('Remove {{ m.name }} from team?',function(yes){if(yes)form.submit()})">
<button type="submit" class="bg-red/10 text-red py-2 px-3 sm:px-4 rounded-lg text-xs sm:text-sm font-semibold hover:bg-red/20 transition w-full">Remove</button>
</form>
</div>
</div>
{% endfor %}
</div></div>'''

TEAM_FORM_TPL = '''
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-2xl mx-auto">
<h2 class="text-xl sm:text-2xl lg:text-3xl font-black text-ny">{{ "EDIT" if edit else "ADD" }} TEAM MEMBER</h2>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-2xl mx-auto">
<form method="POST" enctype="multipart/form-data" class="bg-dark-card border border-dark-border rounded-xl p-4 sm:p-6 space-y-3 sm:space-y-4">
<div>
<label class="text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase tracking-wider block mb-2">Staff ID</label>
<input type="text" name="staff_id" value="{{ member.staff_id if member else '' }}" {{ "readonly" if edit else "" }}
class="w-full bg-dark border border-dark-border rounded-lg px-3 sm:px-4 py-2.5 sm:py-3 text-on-dark font-mono uppercase text-sm sm:text-base">
</div>
<div>
<label class="text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase tracking-wider block mb-2">Name</label>
<input type="text" name="name" value="{{ member.name if member else '' }}" required
class="w-full bg-dark border border-dark-border rounded-lg px-3 sm:px-4 py-2.5 sm:py-3 text-on-dark text-sm sm:text-base">
</div>
<div>
<label class="text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase tracking-wider block mb-2">Gender</label>
<select name="gender" class="w-full bg-dark border border-dark-border rounded-lg px-3 sm:px-4 py-2.5 sm:py-3 text-on-dark text-sm sm:text-base">
<option value="M" {{ 'selected' if member and member.gender=='M' }}>Male</option>
<option value="F" {{ 'selected' if member and member.gender=='F' }}>Female</option>
</select>
</div>
<div>
<label class="text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase tracking-wider block mb-2">Picture</label>
{% if edit and member.picture %}
<div class="mb-2 overflow-hidden rounded-full w-20 h-20 border-2 border-ny/30">
<img src="/static/uploads/{{ member.picture }}" class="w-full h-full object-cover object-center">
</div>
{% endif %}
<input type="file" name="picture_file" accept="image/*" id="pictureInput"
class="w-full bg-dark border border-dark-border rounded-lg px-3 sm:px-4 py-2.5 sm:py-3 text-on-dark text-sm sm:text-base file:mr-4 file:py-1 file:px-3 file:rounded file:border-0 file:bg-ny file:text-dark file:font-bold file:text-xs"
onchange="previewCrop(this)">
<div id="cropPreview" class="mt-2 hidden">
<p class="text-[10px] text-on-dark-dim mb-1">Preview (will be cropped to circle):</p>
<div class="w-20 h-20 rounded-full overflow-hidden border-2 border-ny/30">
<img id="cropImg" class="w-full h-full object-cover object-center">
</div>
</div>
</div>
<!-- Daily Targets as Form Fields -->
<div>
<label class="text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase tracking-wider block mb-2">Assigned Tasks</label>
<div class="grid grid-cols-2 gap-2 mb-4">
{% for t in all_tasks_db %}
<label class="flex items-center gap-2 bg-dark border border-dark-border rounded-lg px-3 py-2 cursor-pointer hover:border-ny/40 transition">
<input type="checkbox" name="assigned_tasks" value="{{ t.task_key }}"
class="accent-[#FFD700] w-4 h-4"
{% if member and member.tasks_list and t.task_key in member.tasks_list %}checked{% endif %}
{% if not member and default_tasks and t.task_key in default_tasks %}checked{% endif %}>
<span class="text-xs text-on-dark-muted font-semibold">{{ t.label }}</span>
</label>
{% endfor %}
</div>
</div>
<div>
<label class="text-[10px] sm:text-xs font-semibold text-on-dark-dim uppercase tracking-wider block mb-2">Daily Targets</label>
<div class="grid grid-cols-2 gap-2">
{% for t in all_tasks_db %}
<div class="flex items-center gap-2 bg-dark border border-dark-border rounded-lg px-3 py-2">
<span class="text-xs text-on-dark-muted font-semibold">{{ t.label }}</span>
<input type="number" name="target_{{ t.task_key }}" min="0"
value="{{ member.targets_json.get(t.task_key, 0) if member and member.targets_json else (default_targets.get(t.task_key, 0) if default_targets else 0) }}"
class="w-16 bg-dark border border-dark-border rounded px-2 py-1 text-ny font-mono text-xs text-center">
<span class="text-on-dark-dim text-[9px]">{{ t.unit }}/day</span>
</div>
{% endfor %}
</div>
</div>
<div class="flex flex-col sm:flex-row gap-2 sm:gap-3 pt-2">
<button type="submit" class="bg-ny text-dark px-6 py-3 rounded-lg font-bold hover:bg-ny-bright transition flex items-center justify-center gap-2 text-sm">
<span class="material-symbols-outlined">save</span>Save</button>
<a href="/team" class="bg-dark-border text-on-dark px-6 py-3 rounded-lg font-medium hover:bg-dark-hover transition text-center text-sm">Cancel</a>
</div>
</form>
</div>
<script>
function previewCrop(input){
if(input.files&&input.files[0]){
var reader=new FileReader();
reader.onload=function(e){
document.getElementById('cropImg').src=e.target.result;
document.getElementById('cropPreview').classList.remove('hidden');
};
reader.readAsDataURL(input.files[0]);
}
}
</script>'''

# ============== SALES ANALYSIS ==============
ANALYSIS_TPL = '''
<style>
.dp-wrap{display:flex;align-items:flex-end;gap:8px;flex-wrap:wrap}
.dp-group{display:flex;flex-direction:column;gap:3px}
.dp-label{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px}
.dp-input{background:#111;border:1px solid #333;border-radius:6px;padding:6px 10px;color:#fff;font-family:monospace;font-size:12px;outline:none;transition:border-color 0.2s;width:140px}
.dp-input:focus{border-color:#FFD700}
.dp-input::-webkit-calendar-picker-indicator{filter:invert(0.7);cursor:pointer}
.dp-go{background:#FFD700;color:#000;border:none;border-radius:6px;padding:6px 16px;font-size:11px;font-weight:700;cursor:pointer;transition:background 0.2s}
.dp-go:hover{background:#FFE44D}
</style>
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-7xl mx-auto">
<div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 sm:gap-4">
<div>
<h2 class="text-xl sm:text-2xl lg:text-3xl font-black text-ny">SALES ANALYSIS</h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">{{ period_label }}</p>
</div>
<form method="GET" class="dp-wrap">
<div class="dp-group">
<span class="dp-label">From</span>
<input type="date" name="start_date" id="a-start" value="{{ start_date }}" class="dp-input">
</div>
<div class="dp-group">
<span class="dp-label">To</span>
<input type="date" name="end_date" id="a-end" value="{{ end_date }}" class="dp-input">
</div>
<button type="submit" class="dp-go">Apply</button>
</form>
</div>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-7xl mx-auto">
<!-- Range Info -->
<div class="bg-dark-card border border-dark-border rounded-xl p-3 sm:p-4 mb-4 sm:mb-6">
<div class="flex flex-wrap gap-4 sm:gap-6 text-xs">
<div><span class="text-on-dark-dim">Range:</span> <span class="text-ny font-mono font-bold">{{ start_date }} to {{ end_date }}</span></div>
<div><span class="text-on-dark-dim">Days:</span> <span class="text-on-dark font-mono font-bold">{{ range_days }}</span></div>
<div><span class="text-on-dark-dim">Period Target:</span> <span class="text-green-neon font-mono font-bold">per-day &times; {{ range_days }} days</span></div>
</div>
</div>
<!-- Summary Cards -->
<div class="grid grid-cols-2 gap-2 sm:gap-4 mb-6 lg:mb-8">
{% for tk in task_totals %}
<div class="bg-dark-card border border-dark-border rounded-xl p-3 sm:p-4 card-hover">
<div class="text-[10px] sm:text-xs text-on-dark-dim uppercase tracking-wider mb-1">{{ tk|upper }} ({{ task_units[tk] }})</div>
<div class="text-xl sm:text-2xl font-black font-mono text-ny">{{ task_totals[tk]|int }}</div>
<div class="text-[10px] sm:text-xs text-on-dark-muted">across all staff</div>
</div>
{% endfor %}
</div>
<!-- Staff Breakdown -->
<div class="bg-dark-card border border-dark-border rounded-xl overflow-hidden mb-6 lg:mb-8">
<div class="p-3 sm:p-4 lg:p-6 border-b border-dark-border">
<h3 class="text-sm sm:text-lg font-bold text-on-dark">PERFORMANCE BY STAFF</h3>
</div>
<div class="table-scroll">
<table class="w-full text-xs sm:text-sm">
<thead><tr class="border-b border-dark-border">
<th class="py-2 sm:py-3 px-2 sm:px-4 text-[10px] sm:text-xs text-on-dark-dim uppercase text-left">Staff</th>
{% for tk in all_tasks %}
<th class="py-2 sm:py-3 px-2 sm:px-4 text-[10px] sm:text-xs text-on-dark-dim uppercase text-center">{{ tk[:4]|upper }}</th>
{% endfor %}
<th class="py-2 sm:py-3 px-2 sm:px-4 text-[10px] sm:text-xs text-on-dark-dim uppercase text-center">Total</th>
<th class="py-2 sm:py-3 px-2 sm:px-4 text-[10px] sm:text-xs text-on-dark-dim uppercase text-center hidden sm:table-cell">Achievement</th>
</tr></thead>
<tbody class="divide-y divide-dark-border">
{% for s in staff %}
<tr class="hover:bg-dark-hover transition">
<td class="py-2.5 sm:py-3 px-2 sm:px-4">
<div class="flex items-center gap-1.5 sm:gap-2">
<div class="w-7 h-7 sm:w-8 sm:h-8 rounded-full bg-ny/20 flex items-center justify-center text-ny text-xs sm:text-sm font-bold">{{ s[0] }}</div>
<span class="font-semibold text-on-dark text-xs sm:text-sm">{{ s }}</span>
</div>
</td>
{% for tk in all_tasks %}
{% set val = staff_totals.get(s,{}).get(tk,0)|int %}
{% set tgt = targets.get(s, {}).get(tk,0) %}
<td class="py-2.5 sm:py-3 px-2 sm:px-4 text-center font-mono
{% if tgt>0 and val>=tgt %}text-green-neon{% elif val>0 %}text-ny{% else %}text-on-dark-dim{% endif %}">
{{ val }}{% if tgt>0 %}<span class="text-on-dark-dim">/{{ tgt }}</span>{% endif %}</td>
{% endfor %}
<td class="py-2.5 sm:py-3 px-2 sm:px-4 text-center font-mono font-bold text-on-dark">{{ staff_totals.get(s,{}).get('total',0)|int }}</td>
<td class="py-2.5 sm:py-3 px-2 sm:px-4 text-center hidden sm:table-cell">
{% set pct = achievement_pct.get(s,0) %}
<div class="w-full bg-dark-border rounded-full h-1.5 sm:h-2 mb-1">
<div class="h-1.5 sm:h-2 rounded-full {% if pct>=100 %}bg-green-neon{% elif pct>=50 %}bg-ny{% else %}bg-red{% endif %}" style="width:{{ [pct,100]|min }}%"></div>
</div>
<span class="text-[10px] sm:text-xs font-mono {% if pct>=100 %}text-green-neon{% elif pct>=50 %}text-ny{% else %}text-red{% endif %}">{{ pct|int }}%</span>
</td>
</tr>
{% endfor %}
</tbody></table></div></div>
<!-- Export -->
<div class="flex flex-col sm:flex-row gap-2 sm:gap-3">
<a href="/generate?from={{ start_date }}&to={{ end_date }}" class="bg-ny text-dark px-4 sm:px-6 py-3 rounded-lg font-bold hover:bg-ny-bright transition flex items-center justify-center gap-2 text-sm">
<span class="material-symbols-outlined">download</span>Export to Excel</a>
<a href="/" class="bg-dark-border text-on-dark px-4 sm:px-6 py-3 rounded-lg font-medium hover:bg-dark-hover transition text-center text-sm">Back to Dashboard</a>
</div>
</div>'''

# ============== RENDER HELPER ==============
def render(content, page="dashboard"):
    autorefresh = request.args.get("autorefresh") == "1"
    return render_template_string(LAYOUT, content=render_template_string(content), page=page, autorefresh=autorefresh)

# ============== AUTH ROUTES ==============
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        if not _check_rate_limit(f"login:{request.remote_addr}", 5, 60):
            flash("Too many login attempts. Try again in 1 minute.", "error")
            return redirect(url_for("login"))
        username = request.form.get("username","").strip().lower()
        password = request.form.get("password","")
        user = get_user(username)
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user"] = user["username"]
            session["role"] = user["role"]
            session["display_name"] = user["display_name"] or user["username"]
            flash(f"Welcome back, {session['display_name']}!", "success")
            return redirect(url_for("index"))
        flash("Invalid username/email or password", "error")
    return render_template_string(AUTH_LAYOUT, content=render_template_string(LOGIN_TPL))

@app.route("/logout")
def logout():
    session.clear()
    flash("Signed out successfully", "success")
    return redirect(url_for("login"))

# ============== PROFILE & PASSWORD RESET ==============
import random as _random

PROFILE_TPL = '''
<div class="p-3 sm:p-4 lg:p-8 max-w-2xl mx-auto">
  <h2 class="text-xl sm:text-2xl font-bold text-white mb-4 sm:mb-6">Profile</h2>
  {% with msgs = get_flashed_messages(with_categories=true) %}
  {% for cat, msg in msgs %}
  <div class="mb-3 sm:mb-4 px-3 sm:px-4 py-2 sm:py-3 rounded-lg text-sm {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-green-neon/10 border border-green-neon/30 text-green-neon{% endif %}">{{ msg }}</div>
  {% endfor %}{% endwith %}
  <div class="bg-dark-card border border-dark-border rounded-xl p-4 sm:p-6 mb-4 sm:mb-6" style="border-top:3px solid #FFD700">
    <h3 class="text-base sm:text-lg font-bold text-white mb-3 sm:mb-4"><span class="material-symbols-outlined align-middle mr-1">person</span> Account Information</h3>
    <form method="POST" action="/profile">
      <input type="hidden" name="action" value="update_profile">
      <div class="space-y-3 sm:space-y-4">
        <div><label class="block text-xs text-on-dark-muted mb-1 uppercase tracking-wider">Username</label>
        <input type="text" value="{{ user.username }}" disabled class="w-full bg-dark border border-dark-border rounded-lg px-3 py-2.5 text-on-dark-dim font-mono text-sm" style="min-height:44px"></div>
        <div><label class="block text-xs text-on-dark-muted mb-1 uppercase tracking-wider">Display Name</label>
        <input type="text" name="display_name" value="{{ user.display_name or '' }}" class="w-full bg-dark border border-dark-border rounded-lg px-3 py-2.5 text-white text-sm focus:border-ny" style="min-height:44px"></div>
        <div><label class="block text-xs text-on-dark-muted mb-1 uppercase tracking-wider">Email</label>
        <input type="email" name="email" value="{{ user.email or '' }}" placeholder="user@example.com" class="w-full bg-dark border border-dark-border rounded-lg px-3 py-2.5 text-white text-sm focus:border-ny" style="min-height:44px"></div>
        <div><label class="block text-xs text-on-dark-muted mb-1 uppercase tracking-wider">Password</label>
        <div class="flex items-center gap-2"><input type="password" value="••••••••" disabled class="w-full bg-dark border border-dark-border rounded-lg px-3 py-2.5 text-on-dark-dim font-mono text-sm" style="min-height:44px"><a href="/forgot-password" class="text-ny text-xs whitespace-nowrap hover:underline">Change</a></div></div>
        <div><label class="block text-xs text-on-dark-muted mb-1 uppercase tracking-wider">Role</label>
        <span class="inline-block px-3 py-1 rounded text-xs font-bold {% if user.role=='admin' %}bg-red text-white{% else %}bg-dark-border text-on-dark-muted{% endif %}">{{ user.role|upper }}</span></div>
      </div>
      <button type="submit" class="mt-4 bg-ny text-dark px-6 py-2.5 rounded-lg font-bold text-sm hover:bg-ny-bright transition" style="min-height:44px">Save Changes</button>
    </form>
  </div>
  <div class="bg-dark-card border border-dark-border rounded-xl p-4 sm:p-6" style="border-top:3px solid #FFD700">
    <h3 class="text-base sm:text-lg font-bold text-white mb-3 sm:mb-4"><span class="material-symbols-outlined align-middle mr-1">lock</span> Update Password</h3>
    <p class="text-on-dark-muted text-xs sm:text-sm mb-3 sm:mb-4">To update your password, we'll send a verification code to your email address.</p>
    <form method="POST" action="/profile">
      <input type="hidden" name="action" value="send_password_code">
      <button type="submit" class="bg-ny text-dark px-6 py-2.5 rounded-lg font-bold text-sm hover:bg-ny-bright transition" style="min-height:44px">Update Password</button>
    </form>
  </div>
</div>
<script>
function togglePw(id,btn){
  var inp=document.getElementById(id);
  var icon=btn.querySelector('.material-symbols-outlined');
  if(inp.type==='password'){inp.type='text';icon.textContent='visibility_off';}
  else{inp.type='password';icon.textContent='visibility';}
}
</script>
'''

PROFILE_VERIFY_TPL = '''
<div class="text-center mb-3">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-10 h-10 rounded-full object-cover mx-auto mb-1.5 border-2 border-ny/30">
<h1 class="text-lg font-black text-white">OneTrack</h1>
<p class="text-on-dark-muted text-[10px] mt-0.5">Shell Bandar Mahkota Cheras</p>
</div>
<div class="bg-dark-card/90 backdrop-blur-sm border border-dark-border rounded-lg p-3">
<h2 class="text-xs font-bold text-white mb-2">Update Password</h2>
{% with msgs = get_flashed_messages(with_categories=true) %}
{% for cat, msg in msgs %}
<div class="mb-2 px-2.5 py-1.5 rounded text-[11px] {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}{% endwith %}
<div class="flex items-center justify-center gap-1 mb-3">
<span class="step-pill step-active">1 Verify Email</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-pending">2 Set Password</span>
</div>
<p class="text-on-dark-muted text-[11px] mb-2">Enter the 6-digit code sent to <strong class="text-white">{{ email }}</strong></p>
<form method="POST" action="/profile/verify" class="space-y-2">
<input type="hidden" name="email" value="{{ email }}">
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">Verification Code</label>
<input type="text" name="code" maxlength="6" pattern="[0-9]{6}" required placeholder="000000"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs text-center font-mono tracking-[0.3em] placeholder-on-dark-dim"
autocomplete="off" inputmode="numeric" style="min-height:40px;font-size:1.1rem">
</div>
<button type="submit" class="w-full bg-ny text-dark font-bold py-2 rounded text-xs hover:bg-ny-bright transition mt-1" style="min-height:40px">
VERIFY CODE
</button>
</form>
<div class="text-center mt-2">
<a href="/profile" class="text-ny text-[11px] hover:underline">Back to Profile</a>
</div>
</div>
<p class="text-center text-on-dark-dim text-[9px] mt-2">&copy; 2026 Shell Prosper Niaga</p>
'''

PROFILE_RESET_TPL = '''
<div class="text-center mb-3">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-10 h-10 rounded-full object-cover mx-auto mb-1.5 border-2 border-ny/30">
<h1 class="text-lg font-black text-white">OneTrack</h1>
<p class="text-on-dark-muted text-[10px] mt-0.5">Shell Bandar Mahkota Cheras</p>
</div>
<div class="bg-dark-card/90 backdrop-blur-sm border border-dark-border rounded-lg p-3">
<h2 class="text-xs font-bold text-white mb-2">Update Password</h2>
{% with msgs = get_flashed_messages(with_categories=true) %}
{% for cat, msg in msgs %}
<div class="mb-2 px-2.5 py-1.5 rounded text-[11px] {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}{% endwith %}
<div class="flex items-center justify-center gap-1 mb-3">
<span class="step-pill step-done">&#10003; Verify Email</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-active">2 Set Password</span>
</div>
<p class="text-on-dark-muted text-[11px] mb-2">Create a new password for <strong class="text-white">{{ email }}</strong></p>
<form method="POST" action="/profile/reset-password" class="space-y-2">
<input type="hidden" name="email" value="{{ email }}">
<input type="hidden" name="code" value="{{ code }}">
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">New Password</label>
<input type="password" name="new_password" required minlength="6" placeholder="Min. 6 characters"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs placeholder-on-dark-dim" style="min-height:38px">
</div>
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">Confirm Password</label>
<input type="password" name="confirm_password" required placeholder="Re-enter password"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs placeholder-on-dark-dim" style="min-height:38px">
</div>
<button type="submit" class="w-full bg-ny text-dark font-bold py-2 rounded text-xs hover:bg-ny-bright transition mt-1" style="min-height:40px">
SET NEW PASSWORD
</button>
</form>
<div class="text-center mt-2">
<a href="/profile" class="text-ny text-[11px] hover:underline">Back to Profile</a>
</div>
</div>
<p class="text-center text-on-dark-dim text-[9px] mt-2">&copy; 2026 Shell Prosper Niaga</p>
'''

FORGOT_TPL = '''
<div class="text-center mb-3">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-10 h-10 rounded-full object-cover mx-auto mb-1.5 border-2 border-ny/30">
<h1 class="text-lg font-black text-white">OneTrack</h1>
<p class="text-on-dark-muted text-[10px] mt-0.5">Shell Bandar Mahkota Cheras</p>
</div>
<div class="bg-dark-card/90 backdrop-blur-sm border border-dark-border rounded-lg p-3">
<h2 class="text-xs font-bold text-white mb-2">Forgot Password?</h2>
{% with msgs = get_flashed_messages(with_categories=true) %}
{% for cat, msg in msgs %}
<div class="mb-2 px-2.5 py-1.5 rounded text-[11px] {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}{% endwith %}
<p class="text-on-dark-muted text-[11px] mb-2">Enter your email and we'll send a verification code.</p>
<div class="flex items-center justify-center gap-1 mb-3">
<span class="step-pill step-active">1 Email</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-pending">2 Verify</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-pending">3 Password</span>
</div>
<form method="POST" action="/forgot-password" class="space-y-2">
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">Email Address</label>
<input type="email" name="email" required placeholder="user@example.com"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs placeholder-on-dark-dim" style="min-height:38px">
</div>
<button type="submit" class="w-full bg-ny text-dark font-bold py-2 rounded text-xs hover:bg-ny-bright transition mt-1" style="min-height:40px">
SEND VERIFICATION CODE
</button>
</form>
<div class="text-center mt-2">
<a href="/login" class="text-ny text-[11px] hover:underline">Back to Sign In</a>
</div>
</div>
<p class="text-center text-on-dark-dim text-[9px] mt-2">&copy; 2026 Shell Prosper Niaga</p>
'''

VERIFY_TPL = '''
<div class="text-center mb-3">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-10 h-10 rounded-full object-cover mx-auto mb-1.5 border-2 border-ny/30">
<h1 class="text-lg font-black text-white">OneTrack</h1>
<p class="text-on-dark-muted text-[10px] mt-0.5">Shell Bandar Mahkota Cheras</p>
</div>
<div class="bg-dark-card/90 backdrop-blur-sm border border-dark-border rounded-lg p-3">
<h2 class="text-xs font-bold text-white mb-2">Enter Verification Code</h2>
{% with msgs = get_flashed_messages(with_categories=true) %}
{% for cat, msg in msgs %}
<div class="mb-2 px-2.5 py-1.5 rounded text-[11px] {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}{% endwith %}
<div class="flex items-center justify-center gap-1 mb-3">
<span class="step-pill step-done">&#10003; Email</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-active">2 Verify</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-pending">3 Password</span>
</div>
<p class="text-on-dark-muted text-[11px] mb-2">Enter the 6-digit code sent to <strong class="text-white">{{ email }}</strong></p>
<form method="POST" action="/verify-code" class="space-y-2">
<input type="hidden" name="email" value="{{ email }}">
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">Verification Code</label>
<input type="text" name="code" maxlength="6" pattern="[0-9]{6}" required placeholder="000000"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs text-center font-mono tracking-[0.3em] placeholder-on-dark-dim"
autocomplete="off" inputmode="numeric" style="min-height:40px;font-size:1.1rem">
</div>
<button type="submit" class="w-full bg-ny text-dark font-bold py-2 rounded text-xs hover:bg-ny-bright transition mt-1" style="min-height:40px">
VERIFY CODE
</button>
</form>
<div class="text-center mt-2">
<a href="/forgot-password" class="text-ny text-[11px] hover:underline">Back to Email</a>
</div>
</div>
<p class="text-center text-on-dark-dim text-[9px] mt-2">&copy; 2026 Shell Prosper Niaga</p>
'''

RESET_TPL = '''
<div class="text-center mb-3">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-10 h-10 rounded-full object-cover mx-auto mb-1.5 border-2 border-ny/30">
<h1 class="text-lg font-black text-white">OneTrack</h1>
<p class="text-on-dark-muted text-[10px] mt-0.5">Shell Bandar Mahkota Cheras</p>
</div>
<div class="bg-dark-card/90 backdrop-blur-sm border border-dark-border rounded-lg p-3">
<h2 class="text-xs font-bold text-white mb-2">Set New Password</h2>
{% with msgs = get_flashed_messages(with_categories=true) %}
{% for cat, msg in msgs %}
<div class="mb-2 px-2.5 py-1.5 rounded text-[11px] {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}{% endwith %}
<div class="flex items-center justify-center gap-1 mb-3">
<span class="step-pill step-done">&#10003; Email</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-done">&#10003; Verify</span>
<span class="step-arrow">&rarr;</span>
<span class="step-pill step-active">3 Password</span>
</div>
<p class="text-on-dark-muted text-[11px] mb-2">Create a new password for <strong class="text-white">{{ email }}</strong></p>
<form method="POST" action="/reset-password" class="space-y-2">
<input type="hidden" name="email" value="{{ email }}">
<input type="hidden" name="code" value="{{ code }}">
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">New Password</label>
<input type="password" name="new_password" required minlength="6" placeholder="Min. 6 characters"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs placeholder-on-dark-dim" style="min-height:38px">
</div>
<div>
<label class="block text-[9px] font-semibold text-on-dark-dim uppercase tracking-wider mb-0.5">Confirm Password</label>
<input type="password" name="confirm_password" required placeholder="Re-enter password"
class="w-full bg-dark border border-dark-border rounded px-2.5 py-2 text-white text-xs placeholder-on-dark-dim" style="min-height:38px">
</div>
<button type="submit" class="w-full bg-ny text-dark font-bold py-2 rounded text-xs hover:bg-ny-bright transition mt-1" style="min-height:40px">
SET NEW PASSWORD
</button>
</form>
<div class="text-center mt-2">
<a href="/forgot-password" class="text-ny text-[11px] hover:underline">Back to Email</a>
</div>
</div>
<p class="text-center text-on-dark-dim text-[9px] mt-2">&copy; 2026 Shell Prosper Niaga</p>
'''

@app.route("/profile", methods=["GET","POST"])
@login_required
def profile():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE username=?", (session["user"],))
    user = dict(c.fetchone())
    if request.method == "POST":
        action = request.form.get("action","")
        if action == "update_profile":
            display_name = request.form.get("display_name","").strip()
            email = request.form.get("email","").strip()
            c.execute("UPDATE users SET display_name=?, email=? WHERE username=?", (display_name, email, session["user"]))
            conn.commit()
            session["display_name"] = display_name
            flash("Profile updated successfully", "success")
        elif action == "send_password_code":
            email = user.get("email","")
            if not email:
                flash("No email address on file. Please add an email first.", "error")
            else:
                code = str(_random.randint(100000, 999999))
                from datetime import timedelta
                expiry = (datetime.now() + timedelta(minutes=10)).isoformat()
                c.execute("UPDATE users SET reset_code=?, reset_expiry=? WHERE username=?", (code, expiry, session["user"]))
                conn.commit()
                send_email(email, "OneTrack - Password Update Code",
                           f"Your verification code is: {code}\nThis code expires in 10 minutes.")
                flash("Verification code sent! Check your email.", "success")
                conn.close()
                return redirect(url_for("profile_verify", email=email))
        c.execute("SELECT * FROM users WHERE username=?", (session["user"],))
        user = dict(c.fetchone())
    conn.close()
    content = render_template_string(PROFILE_TPL, user=user)
    return render(content, "profile")

@app.route("/profile/verify", methods=["GET","POST"])
@login_required
def profile_verify():
    email = request.args.get("email","") or request.form.get("email","")
    if request.method == "POST":
        code = request.form.get("code","").strip()
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE email=? AND reset_code=? AND username=?", (email, code, session["user"]))
        u = c.fetchone()
        if not u:
            flash("Invalid verification code", "error")
            conn.close()
            return redirect(url_for("profile_verify", email=email))
        from datetime import datetime as dt
        try:
            expiry = dt.fromisoformat(u["reset_expiry"])
            if dt.now() > expiry:
                flash("Code has expired. Please request a new one.", "error")
                conn.close()
                return redirect(url_for("profile"))
        except:
            pass
        conn.close()
        return render_template_string(AUTH_LAYOUT, content=render_template_string(PROFILE_RESET_TPL, email=email, code=code))
    return render_template_string(AUTH_LAYOUT, content=render_template_string(PROFILE_VERIFY_TPL, email=email))

@app.route("/profile/reset-password", methods=["POST"])
@login_required
def profile_reset_password():
    email = request.form.get("email","")
    code = request.form.get("code","")
    new_pw = request.form.get("new_password","")
    confirm = request.form.get("confirm_password","")
    if new_pw != confirm:
        flash("Passwords do not match", "error")
        return redirect(url_for("profile_reset_password", email=email, code=code))
    if len(new_pw) < 6:
        flash("Password must be at least 6 characters", "error")
        return redirect(url_for("profile_reset_password", email=email, code=code))
    if not re.search(r'[A-Z]', new_pw) or not re.search(r'[0-9]', new_pw):
        flash("Password must contain at least 1 uppercase letter and 1 number", "error")
        return redirect(url_for("profile_reset_password", email=email, code=code))
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE email=? AND reset_code=? AND username=?", (email, code, session["user"]))
    u = c.fetchone()
    if not u:
        flash("Invalid session", "error")
        conn.close()
        return redirect(url_for("profile"))
    c.execute("UPDATE users SET password_hash=?, reset_code='', reset_expiry='' WHERE username=?", (hash_pw(new_pw), session["user"]))
    conn.commit()
    conn.close()
    flash("Password updated successfully!", "success")
    return redirect(url_for("profile"))

@app.route("/forgot-password", methods=["GET","POST"])
def forgot_password():
    if request.method == "POST":
        if not _check_rate_limit(f"forgot:{request.remote_addr}", 3, 300):
            flash("Too many requests. Try again in 5 minutes.", "error")
            return redirect(url_for("forgot_password"))
        email = request.form.get("email","").strip().lower()
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE email=?", (email,))
        user = c.fetchone()
        if not user:
            flash("If an account exists with that email, a code has been sent.", "success")
            conn.close()
            return redirect(url_for("verify_code"))
        code = str(_random.randint(100000, 999999))
        from datetime import timedelta
        expiry = (datetime.now() + timedelta(minutes=10)).isoformat()
        c.execute("UPDATE users SET reset_code=?, reset_expiry=? WHERE email=?", (code, expiry, email))
        conn.commit()
        conn.close()
        send_email(email, "OneTrack - Password Reset Code",
                   f"Your verification code is: {code}\nThis code expires in 10 minutes.")
        flash("Verification code sent! Check your email.", "success")
        return redirect(url_for("verify_code", email=email))
    return render_template_string(AUTH_LAYOUT, content=render_template_string(FORGOT_TPL))

@app.route("/verify-code", methods=["GET","POST"])
def verify_code():
    email = request.args.get("email","") or request.form.get("email","")
    if request.method == "POST":
        if not _check_rate_limit(f"verify:{email}:{request.remote_addr}", 5, 300):
            flash("Too many attempts. Try again in 5 minutes.", "error")
            return redirect(url_for("verify_code", email=email))
        code = request.form.get("code","").strip()
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE email=? AND reset_code=?", (email, code))
        user = c.fetchone()
        if not user:
            flash("Invalid verification code", "error")
            conn.close()
            return redirect(url_for("verify_code", email=email))
        from datetime import datetime as dt
        try:
            expiry = dt.fromisoformat(user["reset_expiry"])
            if dt.now() > expiry:
                flash("Code has expired. Please request a new one.", "error")
                conn.close()
                return redirect(url_for("forgot_password"))
        except:
            pass
        conn.close()
        return render_template_string(AUTH_LAYOUT, content=render_template_string(RESET_TPL, email=email, code=code))
    return render_template_string(AUTH_LAYOUT, content=render_template_string(VERIFY_TPL, email=email))

@app.route("/reset-password", methods=["POST"])
def reset_password():
    email = request.form.get("email","")
    code = request.form.get("code","")
    new_pw = request.form.get("new_password","")
    confirm = request.form.get("confirm_password","")
    if new_pw != confirm:
        flash("Passwords do not match", "error")
        return redirect(url_for("reset_password", email=email, code=code))
    if len(new_pw) < 6:
        flash("Password must be at least 6 characters", "error")
        return redirect(url_for("reset_password", email=email, code=code))
    if not re.search(r'[A-Z]', new_pw) or not re.search(r'[0-9]', new_pw):
        flash("Password must contain at least 1 uppercase letter and 1 number", "error")
        return redirect(url_for("reset_password", email=email, code=code))
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE email=? AND reset_code=?", (email, code))
    user = c.fetchone()
    if not user:
        flash("Invalid session", "error")
        conn.close()
        return redirect(url_for("forgot_password"))
    c.execute("UPDATE users SET password_hash=?, reset_code='', reset_expiry='' WHERE email=?", (hash_pw(new_pw), email))
    conn.commit()
    conn.close()
    flash("Password updated! Please sign in.", "success")
    return redirect(url_for("login"))

# ============== ROUTES ==============
@app.route("/")
def index():
    now = datetime.now()
    sel_month = request.args.get("month", "", type=str)
    sel_year = request.args.get("year", "", type=str)
    view_mode = "month" if sel_month else "year"
    if sel_month:
        try:
            sel_month = int(sel_month)
        except:
            sel_month = now.month
            view_mode = "month"
    else:
        sel_month = now.month
    # Get available months from DB
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT date FROM responses")
    available_months = set()
    for row in c.fetchall():
        try:
            dt = datetime.strptime(row[0], "%d/%m/%Y")
            available_months.add((dt.year, dt.month))
        except:
            try:
                dt = datetime.strptime(row[0], "%Y-%m-%d")
                available_months.add((dt.year, dt.month))
            except:
                pass
    conn.close()
    available_months = sorted(available_months)
    available_years = sorted(set(y for y, m in available_months))
    # Default year to most recent year with data
    if not sel_year or not sel_year.isdigit():
        sel_year = available_years[-1] if available_years else now.year
    else:
        sel_year = int(sel_year)
    responses = get_all_responses()
    task_keys, task_labels, task_units = get_task_dicts()
    all_tasks = task_keys
    staff_tasks = get_staff_tasks_from_db()
    today = now.strftime("%Y-%m-%d")
    form_modal = render_template_string(FORM_TPL, today=today, staff=get_staff_order(),
        staff_tasks_json=json.dumps(staff_tasks), task_keys=task_keys,
        task_labels=task_labels, task_units=task_units)
    if view_mode == "month":
        daily, _ = process_responses(month=sel_month, year=sel_year)
        scores = calc_scores(daily, year=sel_year, month=sel_month)
        days_in_month = calendar.monthrange(sel_year, sel_month)[1]
        active_days = len([d for d in range(1, days_in_month+1) if daily.get(d)])
        current_month = datetime(sel_year, sel_month, 1).strftime("%B %Y")
        # Build daily chart data (all tasks, all staff)
        period_staff = set()
        for sdata in daily.values():
            for s in sdata:
                period_staff.add(s)
        period_staff = sorted(period_staff)
        chart_points = {}
        for d in range(1, days_in_month + 1):
            day_data = daily.get(d, {})
            chart_points[d] = {}
            for tk in all_tasks:
                total_tk = 0
                for s in period_staff:
                    total_tk += day_data.get(s, {}).get(tk, 0)
                chart_points[d][tk] = total_tk
        task_totals = {tk: 0 for tk in all_tasks}
        for day_data in daily.values():
            for s in period_staff:
                sday = day_data.get(s, {})
                for tk in all_tasks:
                    task_totals[tk] += sday.get(tk, 0)
        total_sales = sum(task_totals.values())
        # Use calc_scores for rankings (matches leaderboard)
        ranked_staff = [(s, data["total"]) for s, data in sorted(scores.items(), key=lambda x: x[1]["total"], reverse=True) if data["total"] > 0]
        max_score = ranked_staff[0][1] if ranked_staff else 1
        top_name = ranked_staff[0][0] if ranked_staff else "—"
        top_score = ranked_staff[0][1] if ranked_staff else 0
        avg_daily = round(total_sales / max(active_days, 1), 0) if active_days > 0 else 0
        prev_month = sel_month - 1 if sel_month > 1 else 12
        prev_year = sel_year if sel_month > 1 else sel_year - 1
        prev_daily, _ = process_responses(month=prev_month, year=prev_year)
        prev_total = 0
        for day_data in prev_daily.values():
            for sdata in day_data.values():
                for tk in all_tasks:
                    prev_total += sdata.get(tk, 0)
        pct_change = round(((total_sales - prev_total) / prev_total * 100), 1) if prev_total > 0 else 0
        cat_totals = {tk: int(task_totals[tk]) for tk in all_tasks if task_totals.get(tk, 0) > 0}
        chart_labels = [str(d) for d in range(1, days_in_month + 1)]
        chart_data = {
            "total_sales": int(total_sales), "pct_change": pct_change,
            "top_name": top_name, "top_score": int(top_score),
            "avg_daily": int(avg_daily), "ranked_staff": ranked_staff,
            "max_score": max_score, "cat_totals": cat_totals,
            "num_categories": len(cat_totals),
        }
    else:
        # Year view: monthly totals for Jan-Dec
        current_month = f"FY {sel_year}"
        chart_points = {}
        task_totals = {tk: 0 for tk in all_tasks}
        total_sales = 0
        active_days = 0
        # Aggregate calc_scores across all months for rankings
        yearly_scores = {}
        for s in get_staff_order():
            yearly_scores[s] = 0
        for m in range(1, 13):
            md, _ = process_responses(month=m, year=sel_year)
            ms = calc_scores(md, year=sel_year, month=m)
            for s, data in ms.items():
                yearly_scores[s] = yearly_scores.get(s, 0) + data["total"]
            chart_points[m] = {}
            for tk in all_tasks:
                mtotal = 0
                for day_data in md.values():
                    for sdata in day_data.values():
                        mtotal += sdata.get(tk, 0)
                chart_points[m][tk] = mtotal
                task_totals[tk] += mtotal
                total_sales += mtotal
            for day_data in md.values():
                for s, sdata in day_data.items():
                    active_days += 1
        ranked_staff = [(s, v) for s, v in sorted(yearly_scores.items(), key=lambda x: x[1], reverse=True) if v > 0]
        max_score = ranked_staff[0][1] if ranked_staff else 1
        top_name = ranked_staff[0][0] if ranked_staff else "—"
        top_score = ranked_staff[0][1] if ranked_staff else 0
        avg_daily = round(total_sales / max(active_days, 1), 0) if active_days > 0 else 0
        prev_total = 0
        for m in range(1, 13):
            pd, _ = process_responses(month=m, year=sel_year - 1)
            for day_data in pd.values():
                for sdata in day_data.values():
                    for tk in all_tasks:
                        prev_total += sdata.get(tk, 0)
        pct_change = round(((total_sales - prev_total) / prev_total * 100), 1) if prev_total > 0 else 0
        cat_totals = {tk: int(task_totals[tk]) for tk in all_tasks if task_totals.get(tk, 0) > 0}
        chart_labels = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
        chart_data = {
            "total_sales": int(total_sales), "pct_change": pct_change,
            "top_name": top_name, "top_score": int(top_score),
            "avg_daily": int(avg_daily), "ranked_staff": ranked_staff,
            "max_score": max_score, "cat_totals": cat_totals,
            "num_categories": len(cat_totals),
        }
    all_vals = []
    for key_data in chart_points.values():
        for tk in all_tasks:
            all_vals.append(key_data.get(tk, 0))
    y_max_val = max(all_vals) * 1.2 if all_vals and max(all_vals) > 0 else 200
    chart_data["y_max_val"] = y_max_val
    content = render_template_string(DASHBOARD, total=total_sales, active_days=active_days,
        current_month=current_month, responses="", form_modal=form_modal,
        chart_data=chart_data, task_labels=task_labels,
        sel_month=sel_month, sel_year=sel_year, available_months=available_months,
        available_years=available_years, chart_points=chart_points,
        chart_labels=chart_labels, all_tasks=all_tasks,
        view_mode=view_mode, days_in_month=len(chart_labels))
    return render(content, "dashboard")

@app.route("/leaderboard")
def leaderboard():
    now = datetime.now()
    sel_month = request.args.get("month", now.month, type=int)
    sel_year = request.args.get("year", "", type=str)
    # Get available years from DB
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT date FROM responses")
    available_months = set()
    for row in c.fetchall():
        try:
            dt = datetime.strptime(row[0], "%d/%m/%Y")
            available_months.add((dt.year, dt.month))
        except:
            try:
                dt = datetime.strptime(row[0], "%Y-%m-%d")
                available_months.add((dt.year, dt.month))
            except:
                pass
    conn.close()
    available_years = sorted(set(y for y, m in available_months))
    if not sel_year or not sel_year.isdigit():
        sel_year = available_years[-1] if available_years else now.year
    else:
        sel_year = int(sel_year)
    daily, _ = process_responses(month=sel_month, year=sel_year)
    scores = calc_scores(daily, year=sel_year, month=sel_month)
    month_name = datetime(sel_year, sel_month, 1).strftime("%B %Y")
    days_in_month = calendar.monthrange(sel_year, sel_month)[1]
    ranked = sorted(scores.items(), key=lambda x: x[1]["total"], reverse=True)
    member_map = {}
    for m in get_team_members():
        member_map[m["staff_id"]] = dict(m)
    task_keys, task_labels, _ = get_task_dicts()
    staff_tasks_dyn = get_staff_tasks_from_db()
    targets = get_monthly_targets(sel_year, sel_month)

    def calc_pct(s, data):
        total_actual = sum(data["tasks"].get(tk, 0) for tk in staff_tasks_dyn.get(s, STAFF_TASKS.get(s, [])))
        total_target = sum(targets.get(s, {}).get(tk, 0) for tk in staff_tasks_dyn.get(s, STAFF_TASKS.get(s, [])))
        return round((total_actual / total_target * 100), 1) if total_target > 0 else 0

    podium = []
    for i, (s, data) in enumerate(ranked[:3]):
        task_details = sorted(
            [{"key": tk, "label": task_labels.get(tk, tk.upper()), "value": data["tasks"].get(tk, 0)}
             for tk in staff_tasks_dyn.get(s, STAFF_TASKS.get(s, [])) if data["tasks"].get(tk, 0) > 0],
            key=lambda x: x["value"], reverse=True
        )
        podium.append({
            "rank": i+1, "name": s, "score": data["total"],
            "pct": calc_pct(s, data),
            "task_details": task_details,
            "picture": member_map.get(s, {}).get("picture", "")
        })
    rankings = []
    for i, (s, data) in enumerate(ranked):
        rankings.append({
            "rank": i+1, "name": s, "score": data["total"],
            "pct": calc_pct(s, data),
        })
    content = render_template_string(LEADERBOARD_TPL, podium=podium, rankings=rankings, current_month=month_name,
        sel_month=sel_month, sel_year=sel_year, datetime=datetime, now=datetime.now(), is_admin=session.get("role")=="admin",
        available_years=available_years)
    return render(content, "leaderboard")

@app.route("/daily")
def daily_view():
    sel_month = request.args.get("month", datetime.now().month, type=int)
    sel_year = request.args.get("year", "", type=str)
    # Get available years from DB
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT date FROM responses")
    available_months = set()
    for row in c.fetchall():
        try:
            dt = datetime.strptime(row[0], "%d/%m/%Y")
            available_months.add((dt.year, dt.month))
        except:
            try:
                dt = datetime.strptime(row[0], "%Y-%m-%d")
                available_months.add((dt.year, dt.month))
            except:
                pass
    conn.close()
    available_years = sorted(set(y for y, m in available_months))
    if not sel_year or not sel_year.isdigit():
        sel_year = available_years[-1] if available_years else datetime.now().year
    else:
        sel_year = int(sel_year)
    daily, monthly_totals = process_responses(month=sel_month, year=sel_year)
    month_name = datetime(sel_year, sel_month, 1).strftime("%B %Y")
    today = datetime.now().day
    days_in_month = calendar.monthrange(sel_year, sel_month)[1]
    active_days_set = set(d for d in range(1, days_in_month+1) if daily.get(d))
    top3_set = set()
    scores = calc_scores(daily, year=sel_year, month=sel_month)
    targets = get_monthly_targets(sel_year, sel_month)
    ranked = sorted(scores.items(), key=lambda x: x[1]["total"], reverse=True)
    for i, (s, _) in enumerate(ranked[:3]):
        top3_set.add(s)
    # Available months/years from DB
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT DISTINCT date FROM responses")
    available = set()
    for (dstr,) in c.fetchall():
        try:
            dt = datetime.strptime(dstr, "%d/%m/%Y")
            available.add((dt.month, dt.year))
        except:
            pass
    conn.close()
    available_months = sorted(available, key=lambda x: (x[1], x[0]), reverse=True)
    available_years = sorted(set(y for m, y in available))
    task_keys, task_labels, task_units = get_task_dicts()
    task_grand_totals = {tk: 0 for tk in task_keys}
    for s_data in monthly_totals.values():
        for tk in task_keys:
            task_grand_totals[tk] += s_data.get(tk, 0)
    period_staff = get_staff_for_period(daily)
    staff_tasks_dyn = get_staff_tasks_from_db()
    for s in period_staff:
        if s not in staff_tasks_dyn:
            staff_tasks_dyn[s] = [tk for tk in task_keys if monthly_totals.get(s, {}).get(tk, 0) > 0]
    content = render_template_string(DAILY_TPL, daily=daily, staff=period_staff, staff_tasks=staff_tasks_dyn,
        targets=targets, daily_targets=get_daily_targets(), current_month=month_name, today=today, sel_month=sel_month, sel_year=sel_year,
        days_in_month=days_in_month, active_days=active_days_set, top3=top3_set,
        monthly_totals=monthly_totals, available_months=available_months, available_years=available_years,
        task_grand_totals=task_grand_totals, all_tasks=task_keys, task_labels=task_labels, task_units=task_units,
        datetime=datetime, now=datetime.now(), today_month=datetime.now().month, today_year=datetime.now().year)
    return render(content, "daily")

@app.route("/form", methods=["GET","POST"])
@login_required
def form_view():
    task_keys, task_labels, task_units = get_task_dicts()
    staff_tasks = get_staff_tasks_from_db()
    if request.method == "POST":
        raw_date = request.form.get('date', '')
        try:
            dt = datetime.strptime(raw_date, '%Y-%m-%d')
            date_str = dt.strftime('%d/%m/%Y')
        except:
            date_str = raw_date

        task_buckets = {tk: [] for tk in task_keys}
        for s in get_staff_order():
            for tk in staff_tasks.get(s, task_keys):
                key = f"{s.lower()}_{tk}"
                val = request.form.get(key, '0')
                try:
                    qty = int(float(val))
                    if qty > 0:
                        task_buckets[tk].append(f"{s.title()} - {qty}")
                except:
                    pass

        q_fields = [""] * 7
        for i, tk in enumerate(task_keys[:7]):
            q_fields[i] = " ".join(task_buckets.get(tk, []))
        q1, q2, q3, q4, q5, q6, q7 = q_fields

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute('INSERT INTO responses (timestamp,date,q1,q2,q3,q4,q5,q6,q7) VALUES (?,?,?,?,?,?,?,?,?)',
                  (datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                   date_str, q1, q2, q3, q4, q5, q6, q7))
        conn.commit()
        conn.close()
        flash("Entry saved", "success")
        return redirect(url_for("index"))
    today = datetime.now().strftime("%Y-%m-%d")
    content = render_template_string(FORM_PAGE_TPL, today=today, staff=get_staff_order(),
        staff_tasks_json=json.dumps(staff_tasks), task_keys=task_keys,
        task_labels=task_labels, task_units=task_units)
    return render(content, "form")

@app.route("/api/sync", methods=["POST"])
@login_required
def api_sync():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "No JSON data"}), 400
    entries = data.get("entries", [])
    if not entries:
        return jsonify({"error": "No entries"}), 400
    task_keys, _, _ = get_task_dicts()
    staff_tasks = get_staff_tasks_from_db()
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    synced = 0
    for entry in entries:
        raw_date = entry.get("date", "")
        try:
            dt = datetime.strptime(raw_date, "%Y-%m-%d")
            date_str = dt.strftime("%d/%m/%Y")
        except:
            date_str = raw_date
        task_buckets = {tk: [] for tk in task_keys}
        items = entry.get("items", {})
        for staff_id, tasks in items.items():
            for tk, qty in tasks.items():
                try:
                    q = int(float(qty))
                    if q > 0:
                        task_buckets[tk].append(f"{staff_id.title()} - {q}")
                except:
                    pass
        q_fields = [""] * 7
        for i, tk in enumerate(task_keys[:7]):
            q_fields[i] = " ".join(task_buckets.get(tk, []))
        q1, q2, q3, q4, q5, q6, q7 = q_fields
        c.execute('INSERT INTO responses (timestamp,date,q1,q2,q3,q4,q5,q6,q7) VALUES (?,?,?,?,?,?,?,?,?)',
                  (datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                   date_str, q1, q2, q3, q4, q5, q6, q7))
        synced += 1
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "synced": synced})

@app.route("/team")
@admin_required
def team_list():
    members = get_team_members()
    enriched = []
    for m in members:
        d = dict(m)
        try:
            d["targets_json"] = json.loads(d["targets"]) if d["targets"] else {}
        except:
            d["targets_json"] = {}
        enriched.append(d)
    content = render_template_string(TEAM_TPL, members=enriched)
    return render(content, "team")

@app.route("/team/add", methods=["GET","POST"])
@admin_required
def team_add():
    all_tasks_db = get_all_tasks()
    default_targets = DAILY_TARGETS_DEFAULT.get("", {})
    if request.method == "POST":
        picture = ""
        f = request.files.get("picture_file")
        ALLOWED_EXT = {'jpg', 'jpeg', 'png', 'gif', 'webp'}
        if f and f.filename:
            staff_id_raw = re.sub(r'[^A-Za-z0-9_\-]', '', request.form.get('staff_id','').upper())
            ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
            if ext not in ALLOWED_EXT:
                ext = "jpg"
            picture = f"{staff_id_raw}.{ext}"
            save_path = os.path.join(UPLOAD_FOLDER, os.path.basename(picture))
            f.save(save_path)
        assigned_tasks = request.form.getlist("assigned_tasks")
        # Build targets only for assigned tasks
        targets = {}
        for t in all_tasks_db:
            if t['task_key'] in assigned_tasks:
                val = request.form.get(f"target_{t['task_key']}", "0").strip()
                try:
                    targets[t['task_key']] = int(val) if val else 0
                except:
                    targets[t['task_key']] = 0
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        try:
            c.execute("INSERT INTO team_members (staff_id,name,gender,picture,targets,tasks,created_at) VALUES (?,?,?,?,?,?,?)",
                (request.form.get("staff_id","").upper(), request.form.get("name",""),
                 request.form.get("gender","M"), picture,
                 json.dumps(targets), json.dumps(assigned_tasks), datetime.now().isoformat()))
            conn.commit()
            flash("Member added!","success")
        except sqlite3.IntegrityError:
            flash("Staff ID already exists!","error")
        conn.close()
        return redirect(url_for("team_list"))
    content = render_template_string(TEAM_FORM_TPL, edit=False, member=None,
        all_tasks_db=all_tasks_db, default_targets=default_targets,
        default_tasks=STAFF_TASKS.get("", []))
    return render(content, "team")

@app.route("/team/edit/<staff_id>", methods=["GET","POST"])
@admin_required
def team_edit(staff_id):
    all_tasks_db = get_all_tasks()
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    if request.method == "POST":
        picture = ""
        f = request.files.get("picture_file")
        ALLOWED_EXT = {'jpg', 'jpeg', 'png', 'gif', 'webp'}
        if f and f.filename:
            safe_id = re.sub(r'[^A-Za-z0-9_\-]', '', staff_id)
            ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
            if ext not in ALLOWED_EXT:
                ext = "jpg"
            picture = f"{safe_id}.{ext}"
            save_path = os.path.join(UPLOAD_FOLDER, os.path.basename(picture))
            f.save(save_path)
        else:
            c.execute("SELECT picture FROM team_members WHERE staff_id=?", (staff_id,))
            row = c.fetchone()
            picture = row[0] if row else ""
        assigned_tasks = request.form.getlist("assigned_tasks")
        # Build targets only for assigned tasks
        targets = {}
        for t in all_tasks_db:
            if t['task_key'] in assigned_tasks:
                val = request.form.get(f"target_{t['task_key']}", "0").strip()
                try:
                    targets[t['task_key']] = int(val) if val else 0
                except:
                    targets[t['task_key']] = 0
        c.execute("UPDATE team_members SET name=?,gender=?,picture=?,targets=?,tasks=? WHERE staff_id=?",
            (request.form.get("name",""), request.form.get("gender","M"),
             picture, json.dumps(targets), json.dumps(assigned_tasks), staff_id))
        conn.commit()
        conn.close()
        flash("Member updated!","success")
        return redirect(url_for("team_list"))
    c.execute("SELECT * FROM team_members WHERE staff_id=?", (staff_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        flash("Member not found","error")
        return redirect(url_for("team_list"))
    member = dict(row)
    try:
        member["targets_json"] = json.loads(member["targets"]) if member["targets"] else {}
    except:
        member["targets_json"] = {}
    try:
        member["tasks_list"] = json.loads(member["tasks"]) if member.get("tasks") else []
    except:
        member["tasks_list"] = []
    if not member["tasks_list"]:
        member["tasks_list"] = STAFF_TASKS.get(staff_id, [])
    conn.close()
    content = render_template_string(TEAM_FORM_TPL, edit=True, member=member,
        all_tasks_db=all_tasks_db, default_targets={}, default_tasks=[])
    return render(content, "team")

@app.route("/team/delete/<staff_id>", methods=["POST"])
@admin_required
def team_delete(staff_id):
    conn = sqlite3.connect(DB_FILE)
    conn.cursor().execute("DELETE FROM team_members WHERE staff_id=?", (staff_id,))
    conn.commit()
    conn.close()
    flash("Member removed!","success")
    return redirect(url_for("team_list"))

@app.route("/analysis")
@admin_required
def analysis():
    now = datetime.now()
    sd_str = request.args.get("start_date", "")
    ed_str = request.args.get("end_date", "")
    if sd_str and ed_str:
        try:
            sd = datetime.strptime(sd_str, "%Y-%m-%d")
            ed = datetime.strptime(ed_str, "%Y-%m-%d")
        except:
            sd = now.replace(day=1)
            ed = now
    else:
        sd = now.replace(day=1)
        ed = now
    range_days = (ed - sd).days + 1
    daily = {}
    cur = sd
    while cur <= ed:
        month_data, _ = process_responses(month=cur.month, year=cur.year)
        if cur.day in month_data:
            daily[(cur.year, cur.month, cur.day)] = month_data[cur.day]
        cur += timedelta(days=1)
    task_keys, task_labels, task_units = get_task_dicts()
    all_tasks = task_keys
    staff_tasks_dyn = get_staff_tasks_from_db()
    daily_targets = get_daily_targets()
    period_staff = set()
    for sdata in daily.values():
        for s in sdata:
            period_staff.add(s)
    period_staff = sorted(period_staff)
    targets = {}
    for s in period_staff:
        dt = daily_targets.get(s, {})
        targets[s] = {tk: dt.get(tk, 0) * range_days for tk in dt}
    task_totals = {tk: 0 for tk in all_tasks}
    staff_totals = {}
    for s in period_staff:
        staff_totals[s] = {"total": 0}
    for key, sdata in daily.items():
        for s in period_staff:
            sday = sdata.get(s, {})
            for tk in all_tasks:
                val = sday.get(tk, 0)
                staff_totals[s][tk] = staff_totals[s].get(tk, 0) + val
                staff_totals[s]["total"] += val
                task_totals[tk] += val
    achievement_pct = {}
    for s in period_staff:
        total_ach = 0
        count = 0
        for tk in staff_tasks_dyn.get(s, STAFF_TASKS.get(s, [])):
            t = targets.get(s, {}).get(tk, 0)
            v = staff_totals[s].get(tk, 0)
            if t > 0:
                total_ach += min((v / t) * 100, 150)
                count += 1
        achievement_pct[s] = total_ach / count if count else 0
    chart_daily = {}
    for d in range(1, range_days + 1):
        key = (sd.year, sd.month, d)
        day_data = daily.get(key, {})
        chart_daily[d] = {}
        for tk in all_tasks:
            total = 0
            for s in period_staff:
                total += day_data.get(s, {}).get(tk, 0)
            chart_daily[d][tk] = total
    period_label = f"{sd.strftime('%d/%m/%Y')} to {ed.strftime('%d/%m/%Y')} ({range_days} days)"
    content = render_template_string(ANALYSIS_TPL, task_totals=task_totals, task_units=task_units,
        all_tasks=all_tasks, staff=period_staff, staff_totals=staff_totals,
        targets=targets, achievement_pct=achievement_pct, period_label=period_label,
        range_days=range_days, now=now, datetime=datetime, task_labels=task_labels,
        start_date=sd.strftime("%Y-%m-%d"), end_date=ed.strftime("%Y-%m-%d"),
        chart_daily=chart_daily)
    return render(content, "analysis")

# ============== ADMIN: TASK EDITOR ==============
@app.route("/admin/tasks", methods=["GET","POST"])
@admin_required
def admin_tasks():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Handle task category CRUD
    action = request.args.get("action","")
    task_id = request.args.get("task_id","", type=str)

    if action == "delete" and task_id:
        c.execute("DELETE FROM tasks WHERE id=?", (task_id,))
        conn.commit()
        flash("Task category removed", "success")
        conn.close()
        return redirect(url_for("admin_tasks"))

    if request.method == "POST":
        form_type = request.form.get("form_type","")

        if form_type == "add_task":
            new_key = request.form.get("task_key","").strip().lower().replace(" ","")
            new_label = request.form.get("task_label","").strip().upper()
            new_unit = request.form.get("task_unit","").strip().upper()
            if new_key and new_label and new_unit:
                try:
                    c.execute("INSERT INTO tasks (task_key,label,unit,sort_order,active,created_at) VALUES (?,?,?,?,1,?)",
                              (new_key, new_label, new_unit, 99, datetime.now().isoformat()))
                    conn.commit()
                    flash(f"Task '{new_label}' added", "success")
                except sqlite3.IntegrityError:
                    flash(f"Task key '{new_key}' already exists", "error")
            else:
                flash("All fields required", "error")
            conn.close()
            return redirect(url_for("admin_tasks"))

        if form_type == "edit_task":
            eid = request.form.get("edit_task_id","")
            e_label = request.form.get("edit_label","").strip().upper()
            e_unit = request.form.get("edit_unit","").strip().upper()
            if eid and e_label and e_unit:
                c.execute("UPDATE tasks SET label=?, unit=? WHERE id=?", (e_label, e_unit, eid))
                conn.commit()
                flash("Task updated", "success")
            conn.close()
            return redirect(url_for("admin_tasks"))

        if form_type == "staff_tasks":
            staff_id = request.form.get("staff_id","").strip().upper()
            staff_tasks = request.form.getlist("tasks")
            targets = {}
            for tk in staff_tasks:
                val = request.form.get(f"target_{tk}","").strip()
                if val:
                    try:
                        targets[tk] = int(val)
                    except:
                        targets[tk] = 0
            c.execute("UPDATE team_members SET targets=?, tasks=? WHERE staff_id=?", (json.dumps(targets), json.dumps(staff_tasks), staff_id))
            conn.commit()
            flash(f"Updated tasks for {staff_id}", "success")
            conn.close()
            return redirect(url_for("admin_tasks"))

        if form_type == "bulk_manage":
            bulk_tasks = request.form.getlist("bulk_tasks")
            bulk_staff = request.form.getlist("bulk_staff")
            bulk_action = request.form.get("bulk_action", "apply")
            if bulk_tasks and bulk_staff:
                updated = 0
                for sid in bulk_staff:
                    c.execute("SELECT targets, tasks FROM team_members WHERE staff_id=?", (sid,))
                    row = c.fetchone()
                    if row:
                        try:
                            t = json.loads(row[0]) if row[0] else {}
                        except:
                            t = {}
                        try:
                            existing_tasks = json.loads(row[1]) if row[1] else []
                        except:
                            existing_tasks = []
                        for bt in bulk_tasks:
                            bt_val = request.form.get(f"bulk_target_{bt}", "0").strip()
                            try:
                                t[bt] = int(bt_val) if bt_val else 0
                            except:
                                t[bt] = 0
                        if bulk_action == "remove":
                            merged = [tk for tk in existing_tasks if tk not in bulk_tasks]
                        else:
                            merged = list(set(existing_tasks) | set(bulk_tasks))
                        cleaned_targets = {k: v for k, v in t.items() if k in merged}
                        c.execute("UPDATE team_members SET targets=?, tasks=? WHERE staff_id=?", (json.dumps(cleaned_targets), json.dumps(merged), sid))
                        updated += 1
                conn.commit()
                action_word = "Removed" if bulk_action == "remove" else "Updated"
                flash(f"{action_word} {', '.join(t.upper() for t in bulk_tasks)} for {updated} staff", "success")
            else:
                flash("Select at least one task and one staff member", "error")
            conn.close()
            return redirect(url_for("admin_tasks"))

        if form_type == "save_all_staff":
            c.execute("SELECT staff_id FROM team_members ORDER BY id")
            all_staff_ids = [r[0] for r in c.fetchall()]
            updated = 0
            for sid in all_staff_ids:
                form_key = f"tasks_{sid}"
                if form_key not in request.form:
                    continue
                checked_tasks = request.form.getlist(form_key)
                existing_targets = {}
                c.execute("SELECT targets FROM team_members WHERE staff_id=?", (sid,))
                row = c.fetchone()
                if row:
                    try:
                        existing_targets = json.loads(row[0]) if row[0] else {}
                    except:
                        existing_targets = {}
                targets = {}
                if checked_tasks:
                    for tk in checked_tasks:
                        val = request.form.get(f"target_{sid}_{tk}", "0").strip()
                        try:
                            targets[tk] = int(val) if val else 0
                        except:
                            targets[tk] = 0
                c.execute("UPDATE team_members SET targets=?, tasks=? WHERE staff_id=?", (json.dumps(targets), json.dumps(checked_tasks), sid))
                updated += 1
            conn.commit()
            flash(f"Saved tasks for {updated} staff members", "success")
            conn.close()
            return redirect(url_for("admin_tasks"))

    # GET: load data
    all_tasks_db = get_all_tasks()
    c.execute("SELECT * FROM team_members ORDER BY id")
    members = [dict(r) for r in c.fetchall()]
    for m in members:
        m["targets"] = json.loads(m["targets"]) if m["targets"] else {}
        try:
            m["tasks_list"] = json.loads(m["tasks"]) if m.get("tasks") else []
        except:
            m["tasks_list"] = []
        if not m["tasks_list"]:
            m["tasks_list"] = STAFF_TASKS.get(m["staff_id"], [])
    conn.close()
    content = render_template_string(ADMIN_TASKS_TPL, members=members,
        all_tasks_db=all_tasks_db)
    return render(content, "admin_tasks")

# ============== ADMIN: RECORD EDITOR ==============
@app.route("/admin/records")
@admin_required
def admin_records():
    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    start_date = request.args.get("start_date", now.strftime("%Y-%m-01"))
    end_date = request.args.get("end_date", today_str)
    sel_staff = request.args.get("staff", "")
    view_mode = request.args.get("view", "list")

    try:
        sd = datetime.strptime(start_date, "%Y-%m-%d")
        sd_db = sd.strftime("%d/%m/%Y")
    except:
        sd = now.replace(day=1)
        sd_db = sd.strftime("%d/%m/%Y")
        start_date = sd.strftime("%Y-%m-%d")
    try:
        ed = datetime.strptime(end_date, "%Y-%m-%d")
        ed_db = ed.strftime("%d/%m/%Y")
    except:
        ed = now
        ed_db = ed.strftime("%d/%m/%Y")
        end_date = ed.strftime("%Y-%m-%d")

    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    query = "SELECT * FROM responses WHERE (substr(date,7,4)||substr(date,4,2)||substr(date,1,2)) >= ? AND (substr(date,7,4)||substr(date,4,2)||substr(date,1,2)) <= ?"
    params = [sd.strftime("%Y%m%d"), ed.strftime("%Y%m%d")]
    if sel_staff:
        query += " AND (q1 LIKE ? OR q2 LIKE ? OR q3 LIKE ? OR q4 LIKE ? OR q5 LIKE ? OR q6 LIKE ? OR q7 LIKE ?)"
        for _ in range(7):
            params.append(f"%{sel_staff}%")
    c.execute(query + " ORDER BY id DESC", params)
    records = [dict(r) for r in c.fetchall()]

    task_keys, task_labels, task_units = get_task_dicts()
    grid_data = {}
    grid_groups = {}
    grid_staff = sel_staff if view_mode == "grid" else ""

    # Build list_groups for list view (group records by year/month)
    mon_list = ['','January','February','March','April','May','June','July','August','September','October','November','December']
    list_yr_month = {}
    for rec in records:
        try:
            dt = datetime.strptime(rec["date"], "%d/%m/%Y")
        except:
            try:
                dt = datetime.strptime(rec["date"], "%Y-%m-%d")
            except:
                continue
        yr = str(dt.year)
        mn = dt.strftime("%m")
        list_yr_month.setdefault(yr, {}).setdefault(mn, []).append(rec)
    list_groups = []
    for yr in sorted(list_yr_month.keys(), reverse=True):
        months_data = []
        for mn in sorted(list_yr_month[yr].keys(), reverse=True):
            months_data.append({"mn": mn, "name": mon_list[int(mn)], "records": list_yr_month[yr][mn]})
        list_groups.append({"yr": yr, "months": months_data})

    list_show_year = len(list_groups) > 1
    list_show_month = sum(len(yg["months"]) for yg in list_groups) > 1

    if view_mode == "grid" and grid_staff:
        c.execute("SELECT * FROM responses ORDER BY date")
        all_rows = [dict(r) for r in c.fetchall()]
        filtered_rows = []
        for row in all_rows:
            try:
                dt = datetime.strptime(row["date"], "%d/%m/%Y")
            except:
                try:
                    dt = datetime.strptime(row["date"], "%Y-%m-%d")
                except:
                    continue
            if sd <= dt <= ed:
                filtered_rows.append((dt, row))
        date_map = {}
        for dt, row in filtered_rows:
            date_key = dt.strftime("%d/%m/%Y")
            if date_key not in date_map:
                date_map[date_key] = {"dt": dt, "row": dict(row)}
            else:
                existing = date_map[date_key]["row"]
                for qk in ["q1","q2","q3","q4","q5","q6","q7"]:
                    if row[qk] and str(row[qk]).strip():
                        existing_entries = {}
                        if existing[qk] and str(existing[qk]).strip():
                            for name, qty in parse_answer(existing[qk]):
                                existing_entries[name] = existing_entries.get(name, 0) + qty
                        for name, qty in parse_answer(row[qk]):
                            existing_entries[name] = existing_entries.get(name, 0) + qty
                        merged = ", ".join(f"{n} - {q}" for n, q in existing_entries.items())
                        existing[qk] = merged

        sorted_dates = sorted(date_map.keys(), key=lambda d: date_map[d]["dt"])
        mon_list = ['','January','February','March','April','May','June','July','August','September','October','November','December']
        yr_month_map = {}
        for date_key in sorted_dates:
            row = date_map[date_key]["row"]
            dt_obj = date_map[date_key]["dt"]
            yr = str(dt_obj.year)
            mn = dt_obj.strftime("%m")
            yr_month_map.setdefault(yr, {}).setdefault(mn, []).append(date_key)
            grid_data[date_key] = {}
            for qi, tk in enumerate(task_keys[:7]):
                qk = f"q{qi+1}"
                val = row.get(qk, "") or ""
                pairs = parse_answer(val)
                staff_val = ""
                for name, qty in pairs:
                    if name.upper() == grid_staff.upper():
                        staff_val = str(int(qty) if qty == int(qty) else qty)
                        break
                grid_data[date_key][tk] = staff_val

        grid_groups = []
        for yr in sorted(yr_month_map.keys(), reverse=True):
            months_data = []
            for mn in sorted(yr_month_map[yr].keys(), reverse=True):
                months_data.append({"mn": mn, "name": mon_list[int(mn)], "dates": yr_month_map[yr][mn]})
            grid_groups.append({"yr": yr, "months": months_data})

        if not grid_groups:
            mon_list_fb = ['','January','February','March','April','May','June','July','August','September','October','November','December']
            yr_month_fb = {}
            sd2 = sd
            while sd2 <= ed:
                date_key = sd2.strftime("%d/%m/%Y")
                yr = str(sd2.year)
                mn = sd2.strftime("%m")
                yr_month_fb.setdefault(yr, {}).setdefault(mn, []).append(date_key)
                grid_data[date_key] = {tk: "" for tk in task_keys[:7]}
                sd2 += timedelta(days=1)
            for yr in sorted(yr_month_fb.keys(), reverse=True):
                months_fb = []
                for mn in sorted(yr_month_fb[yr].keys(), reverse=True):
                    months_fb.append({"mn": mn, "name": mon_list_fb[int(mn)], "dates": yr_month_fb[yr][mn]})
                grid_groups.append({"yr": yr, "months": months_fb})

    show_year_heading = len(grid_groups) > 1
    show_month_heading = sum(len(yg["months"]) for yg in grid_groups) > 1

    # Get all staff from responses (current + removed)
    c.execute("SELECT * FROM responses WHERE (substr(date,7,4)||substr(date,4,2)||substr(date,1,2)) >= ? AND (substr(date,7,4)||substr(date,4,2)||substr(date,1,2)) <= ?", [sd.strftime("%Y%m%d"), ed.strftime("%Y%m%d")])
    all_staff_names = set()
    for row in c.fetchall():
        for qi in range(7):
            val = row[qi+3] if row[qi+3] else ""
            for name, _ in parse_answer(val):
                all_staff_names.add(name.upper())
    current_staff = get_staff_order()
    all_staff_list = current_staff + [n for n in all_staff_names if n not in current_staff]

    conn.close()

    content = render_template_string(RECORDS_TPL, records=records,
        start_date=start_date, end_date=end_date,
        sel_staff=sel_staff, staff_order=all_staff_list, now=now, datetime=datetime,
        view_mode=view_mode, grid_staff=grid_staff, grid_data=grid_data,
        grid_groups=grid_groups, task_keys=task_keys, task_labels=task_labels,
        show_year_heading=show_year_heading, show_month_heading=show_month_heading,
        list_groups=list_groups, list_show_year=list_show_year, list_show_month=list_show_month,
        record_count=len(records))
    return render(content, "admin_records")

@app.route("/admin/records/bulk-update", methods=["POST"])
@admin_required
def admin_records_bulk_update():
    grid_staff = request.form.get("grid_staff", "").strip().upper()
    start_date = request.form.get("start_date", "")
    end_date = request.form.get("end_date", "")
    if not grid_staff:
        flash("No staff selected", "error")
        return redirect(url_for("admin_records", view="grid"))

    task_keys, _, _ = get_task_dicts()
    try:
        sd = datetime.strptime(start_date, "%Y-%m-%d")
        ed = datetime.strptime(end_date, "%Y-%m-%d")
    except:
        flash("Invalid date range", "error")
        return redirect(url_for("admin_records", view="grid"))

    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    updated = 0
    current = sd
    while current <= ed:
        date_key_compact = current.strftime("%d%m%Y")
        day_db = current.strftime("%d/%m/%Y")
        new_vals = {}
        for tk in task_keys[:7]:
            raw = request.form.get(f"day_{date_key_compact}_{tk}", "").strip()
            if raw:
                try:
                    v = float(raw)
                    if v != 0:
                        new_vals[tk] = v
                except:
                    pass
        c.execute("SELECT * FROM responses WHERE date=?", (day_db,))
        existing = c.fetchone()
        if existing:
            existing_dict = dict(existing)
            for qi, tk in enumerate(task_keys[:7]):
                qk = f"q{qi+1}"
                old_val = existing_dict.get(qk, "") or ""
                pairs = parse_answer(old_val)
                pairs = [(n, q) for n, q in pairs if n.upper() != grid_staff]
                if tk in new_vals:
                    pairs.append((grid_staff, new_vals[tk]))
                new_str = " ".join(f"{n} - {int(q) if q == int(q) else q}" for n, q in pairs)
                c.execute(f"UPDATE responses SET {qk}=? WHERE id=?", (new_str, existing_dict["id"]))
            updated += 1
        elif new_vals:
            cols = ["date"]
            vals = [day_db]
            for qi, tk in enumerate(task_keys[:7]):
                qk = f"q{qi+1}"
                cols.append(qk)
                if tk in new_vals:
                    v = new_vals[tk]
                    vals.append(f"{grid_staff} - {int(v) if v == int(v) else v}")
                else:
                    vals.append("")
            placeholders = ",".join(["?"] * len(vals))
            c.execute(f"INSERT INTO responses ({','.join(cols)}) VALUES ({placeholders})", vals)
            updated += 1
        current += timedelta(days=1)

    conn.commit()
    conn.close()
    flash(f"Saved {updated} day(s) for {grid_staff}", "success")
    return redirect(url_for("admin_records", view="grid", staff=grid_staff.lower(), start_date=start_date, end_date=end_date))

@app.route("/admin/records/bulk-delete", methods=["POST"])
@admin_required
def admin_records_bulk_delete():
    record_ids = request.form.getlist("record_ids")
    if not record_ids:
        flash("No records selected", "error")
        return redirect(url_for("admin_records"))
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    placeholders = ",".join(["?"] * len(record_ids))
    c.execute(f"DELETE FROM responses WHERE id IN ({placeholders})", [int(x) for x in record_ids])
    deleted = c.rowcount
    conn.commit()
    conn.close()
    flash(f"Deleted {deleted} record(s)", "success")
    return redirect(url_for("admin_records"))

@app.route("/admin/records/bulk-delete-days", methods=["POST"])
@admin_required
def admin_records_bulk_delete_days():
    grid_staff = request.form.get("grid_staff", "").strip().upper()
    start_date = request.form.get("start_date", "")
    end_date = request.form.get("end_date", "")
    delete_days_raw = request.form.get("delete_days", "")
    if not grid_staff or not delete_days_raw:
        flash("No days selected", "error")
        return redirect(url_for("admin_records", view="grid"))

    delete_dates = [d.strip() for d in delete_days_raw.split(",") if d.strip()]
    try:
        sd = datetime.strptime(start_date, "%Y-%m-%d")
        ed = datetime.strptime(end_date, "%Y-%m-%d")
    except:
        flash("Invalid date range", "error")
        return redirect(url_for("admin_records", view="grid"))

    task_keys, _, _ = get_task_dicts()
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    cleared = 0
    current = sd
    while current <= ed:
        day_db = current.strftime("%d/%m/%Y")
        if day_db in delete_dates:
            day_db = current.strftime("%d/%m/%Y")
            c.execute("SELECT * FROM responses WHERE date=?", (day_db,))
            row = c.fetchone()
            if row:
                row_dict = dict(row)
                all_empty = True
                updates = {}
                for qi, tk in enumerate(task_keys[:7]):
                    qk = f"q{qi+1}"
                    old_val = row_dict.get(qk, "") or ""
                    pairs = parse_answer(old_val)
                    pairs = [(n, q) for n, q in pairs if n.upper() != grid_staff]
                    if pairs:
                        all_empty = False
                    new_str = " ".join(f"{n} - {int(q) if q == int(q) else q}" for n, q in pairs)
                    updates[qk] = new_str
                if all_empty:
                    c.execute("DELETE FROM responses WHERE id=?", (row_dict["id"],))
                else:
                    set_clause = ", ".join(f"{k}=?" for k in updates)
                    c.execute(f"UPDATE responses SET {set_clause} WHERE id=?", list(updates.values()) + [row_dict["id"]])
                cleared += 1
        current += timedelta(days=1)

    conn.commit()
    conn.close()
    flash(f"Cleared {grid_staff} data from {cleared} day(s)", "success")
    return redirect(url_for("admin_records", view="grid", staff=grid_staff.lower(), start_date=start_date, end_date=end_date))

@app.route("/admin/records/edit/<int:record_id>", methods=["GET","POST"])
@admin_required
def admin_record_edit(record_id):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    if request.method == "POST":
        date_str = request.form.get("date","")
        q1 = request.form.get("q1","").strip()
        q2 = request.form.get("q2","").strip()
        q3 = request.form.get("q3","").strip()
        q4 = request.form.get("q4","").strip()
        q5 = request.form.get("q5","").strip()
        q6 = request.form.get("q6","").strip()
        q7 = request.form.get("q7","").strip()
        c.execute("UPDATE responses SET date=?,q1=?,q2=?,q3=?,q4=?,q5=?,q6=?,q7=? WHERE id=?",
                  (date_str,q1,q2,q3,q4,q5,q6,q7,record_id))
        conn.commit()
        conn.close()
        flash(f"Record #{record_id} updated", "success")
        return redirect(url_for("admin_records"))
    c.execute("SELECT * FROM responses WHERE id=?", (record_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        flash(f"Record #{record_id} not found", "error")
        return redirect(url_for("admin_records"))
    rec = dict(row)
    task_keys, task_labels, task_units = get_task_dicts()
    content = render_template_string(RECORD_EDIT_TPL, rec=rec, staff_order=get_staff_order(),
        task_labels=task_labels, task_units=task_units, task_keys=task_keys)
    return render(content, "admin_record_edit")

@app.route("/admin/records/delete/<int:record_id>", methods=["POST"])
@admin_required
def admin_record_delete(record_id):
    conn = sqlite3.connect(DB_FILE)
    conn.cursor().execute("DELETE FROM responses WHERE id=?", (record_id,))
    conn.commit()
    conn.close()
    flash(f"Record #{record_id} deleted", "success")
    return redirect(url_for("admin_records"))

@app.route("/generate")
@admin_required
def generate():
    try:
        from_date = request.args.get("from", "")
        to_date = request.args.get("to", "")
        staff_filter = request.args.get("staff", "").strip().upper()
        result = generate_excel_new(from_date, to_date, staff_filter)
        if isinstance(result, dict):
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fname, fdata in result.items():
                    zf.writestr(fname, fdata.getvalue())
            mem.seek(0)
            return send_file(mem, mimetype="application/zip",
                             as_attachment=True,
                             download_name=f"OneTrack_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")
        else:
            fname = getattr(result, 'name', None) or f"OneTrack_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            if not fname.endswith('.xlsx'):
                fname += '.xlsx'
            return send_file(result, as_attachment=True, download_name=fname,
                             mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return "<h1 style='color:red;background:#1A1A1A;padding:20px'>Error</h1><p style='color:white;background:#262626;padding:20px'>Something went wrong. Please try again.</p><a href='/' style='color:#FFD700'>Back</a>"

@app.route("/clear", methods=["POST"])
@admin_required
def clear():
    conn = sqlite3.connect(DB_FILE)
    conn.cursor().execute("DELETE FROM responses")
    conn.commit()
    conn.close()
    return redirect(url_for("index"))

# ============== EXCEL GENERATION ==============
SITE_ID = "2365"
SITE_NAME = "SHELL BANDAR MAHKOTA CHERAS"
EXCEL_HEADER_FILL = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
EXCEL_HEADER_FONT = Font(name="Arial", size=10, bold=True)
EXCEL_DATA_FONT = Font(name="Arial", size=10)
EXCEL_THIN = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
EXCEL_MET_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
EXCEL_MISSED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FLAT_COLS = ["Site Id", "Site Name", "Date", "Day", "Staff", "Tasks", "Count", "Target", "Variance", "Status"]
DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

def generate_excel_new(from_date="", to_date="", staff_filter=""):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT * FROM responses ORDER BY id")
    rows = c.fetchall()
    conn.close()
    task_keys_db, task_labels, _ = get_task_dicts()
    answer_keys = task_keys_db[:7]
    while len(answer_keys) < 7:
        answer_keys.append("")
    records = []
    dt_cache = get_daily_targets()
    for row in rows:
        try:
            dt = datetime.strptime(row[2], "%d/%m/%Y")
        except:
            try:
                dt = datetime.strptime(row[2], "%Y-%m-%d")
            except:
                continue
        if from_date:
            try:
                fd = datetime.strptime(from_date, "%Y-%m-%d")
                if dt < fd: continue
            except: pass
        if to_date:
            try:
                td = datetime.strptime(to_date, "%Y-%m-%d")
                if dt > td: continue
            except: pass
        date_str = dt.strftime("%d/%m/%Y")
        day_name = DAY_NAMES[dt.weekday()]
        for i, answer in enumerate(row[3:10]):
            if not answer or not str(answer).strip(): continue
            task_key = answer_keys[i]
            if not task_key: continue
            for name, qty in parse_answer(answer):
                target = dt_cache.get(name, {}).get(task_key, 0)
                variance = qty - target
                status = "MET" if variance >= 0 else "MISSED"
                records.append({"date": dt, "date_str": date_str, "day": day_name,
                    "staff": name, "tasks": task_labels.get(task_key, task_key.upper()),
                    "count": qty, "target": target, "variance": variance, "status": status})
    if staff_filter:
        records = [r for r in records if r["staff"] == staff_filter]
    if not records:
        wb = openpyxl.Workbook()
        wb.active["A1"] = "No data found for selected date range"
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return buf
    by_year = {}
    for r in records:
        by_year.setdefault(r["date"].year, []).append(r)
    if len(by_year) == 1:
        year = list(by_year.keys())[0]
        wb = _build_year_workbook(year, by_year[year])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        buf.name = f"OneTrack_{year}.xlsx"
        return buf
    else:
        result = {}
        for year, yr_records in sorted(by_year.items()):
            wb = _build_year_workbook(year, yr_records)
            buf = io.BytesIO(); wb.save(buf); buf.seek(0)
            result[f"OneTrack_{year}.xlsx"] = buf
        return result

def _build_year_workbook(year, records):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    by_month = {}
    for r in records:
        by_month.setdefault(r["date"].month, []).append(r)
    for month_num in range(1, 13):
        if month_num not in by_month: continue
        month_records = by_month[month_num]
        month_name = datetime(year, month_num, 1).strftime("%b %Y")
        ws = wb.create_sheet(title=month_name)
        ws.merge_cells("A1:J1")
        ws["A1"] = f"SITE HERO SALES LEADERBOARD \u2014 {month_name}"
        ws["A1"].font = Font(name="Arial", size=12, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")
        ws["A2"] = "Station Name"; ws["B2"] = SITE_NAME
        ws["A3"] = "Site Id"; ws["B3"] = SITE_ID
        ws["A4"] = "Month"; ws["B4"] = month_name
        for cell in ["A2", "A3", "A4"]:
            ws[cell].font = Font(name="Arial", size=10, bold=True)
        for col_idx, col_name in enumerate(FLAT_COLS, 1):
            c = ws.cell(row=6, column=col_idx, value=col_name)
            c.font = EXCEL_HEADER_FONT; c.fill = EXCEL_HEADER_FILL
            c.border = EXCEL_THIN; c.alignment = Alignment(horizontal="center")
        row_idx = 7
        total_count = 0
        total_target = 0
        total_variance = 0
        for r in sorted(month_records, key=lambda x: (x["date"], x["staff"], x["tasks"])):
            vals = [SITE_ID, SITE_NAME, r["date_str"], r["day"], r["staff"],
                    r["tasks"], r["count"], r["target"], r["variance"], r["status"]]
            total_count += r["count"]
            total_target += r["target"]
            total_variance += r["variance"]
            for col_idx, val in enumerate(vals, 1):
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                c.font = EXCEL_DATA_FONT; c.border = EXCEL_THIN
                c.alignment = Alignment(horizontal="center")
                if col_idx == 10:
                    c.fill = EXCEL_MET_FILL if val == "MET" else EXCEL_MISSED_FILL
            row_idx += 1
        gt_row = row_idx + 1
        gt_status = "MET" if total_variance >= 0 else "MISSED"
        EXCEL_TOTAL_FONT = Font(name="Arial", size=10, bold=True)
        EXCEL_TOTAL_FILL = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
        for col_idx in range(1, 11):
            c = ws.cell(row=gt_row, column=col_idx)
            c.border = EXCEL_THIN
            c.alignment = Alignment(horizontal="center")
            c.font = EXCEL_TOTAL_FONT
            c.fill = EXCEL_TOTAL_FILL
        ws.cell(row=gt_row, column=5, value="GRAND TOTAL")
        ws.cell(row=gt_row, column=7, value=total_count)
        ws.cell(row=gt_row, column=8, value=total_target)
        ws.cell(row=gt_row, column=9, value=total_variance)
        status_cell = ws.cell(row=gt_row, column=10, value=gt_status)
        status_cell.fill = EXCEL_MET_FILL if gt_status == "MET" else EXCEL_MISSED_FILL
        status_cell.font = Font(name="Arial", size=10, bold=True)
        for i, w in enumerate([10, 32, 12, 6, 14, 16, 8, 8, 9, 8]):
            ws.column_dimensions[get_column_letter(i + 1)].width = w
    return wb

# ============== ADMIN TEMPLATES ==============
ADMIN_TASKS_TPL = '''
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-5xl mx-auto">
<h2 class="text-xl sm:text-2xl font-bold text-white">Task Editor</h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">Manage task categories, bulk assign, and per-staff targets</p>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-5xl mx-auto space-y-6">

<!-- Section 1: Task Categories -->
<div class="bg-dark-card border border-dark-border rounded-lg p-4">
<h3 class="text-white font-bold text-sm mb-3">Task Categories</h3>
<div class="space-y-2 mb-4">
{% for t in all_tasks_db %}
<div class="flex items-center justify-between py-2 border-b border-dark-border last:border-0">
<div class="flex items-center gap-3">
<span class="font-mono text-ny text-xs font-bold">{{ t.task_key }}</span>
<span class="text-white text-sm font-semibold">{{ t.label }}</span>
<span class="text-on-dark-dim text-xs">/ {{ t.unit }}</span>
{% if not t.active %}<span class="text-red text-[10px] ml-1">INACTIVE</span>{% endif %}
</div>
<div class="flex gap-3">
<button onclick="toggleEdit({{ t.id }},'{{ t.label }}','{{ t.unit }}')" class="text-ny text-xs hover:underline">Edit</button>
<a href="?action=delete&task_id={{ t.id }}" class="text-red text-xs hover:underline"
onclick="event.preventDefault();var link=this;showDeleteModal('Remove task {{ t.label }}? This will not delete existing data.',function(yes){if(yes)location.href=link.href})">Remove</a>
</div>
</div>
{% endfor %}
</div>
<div id="edit-task-form" class="hidden border border-ny/30 rounded p-3 mb-4 bg-dark">
<p class="text-ny text-xs font-bold mb-2">Edit Task</p>
<form method="POST" class="flex flex-wrap items-end gap-2">
<input type="hidden" name="form_type" value="edit_task">
<input type="hidden" name="edit_task_id" id="edit_task_id">
<div><label class="block text-[10px] text-on-dark-dim uppercase">Label</label>
<input type="text" name="edit_label" id="edit_label" class="bg-dark border border-dark-border rounded px-2 py-1 text-white text-sm w-32"></div>
<div><label class="block text-[10px] text-on-dark-dim uppercase">Unit</label>
<input type="text" name="edit_unit" id="edit_unit" class="bg-dark border border-dark-border rounded px-2 py-1 text-white text-sm w-24"></div>
<button type="submit" class="bg-ny text-dark px-3 py-1 rounded text-xs font-bold">Save</button>
<button type="button" onclick="document.getElementById('edit-task-form').classList.add('hidden')" class="text-on-dark-dim text-xs">Cancel</button>
</form>
</div>
<form method="POST" class="flex flex-wrap items-end gap-2">
<input type="hidden" name="form_type" value="add_task">
<div><label class="block text-[10px] text-on-dark-dim uppercase">Key</label>
<input type="text" name="task_key" placeholder="e.g. fuel" required class="bg-dark border border-dark-border rounded px-2 py-1 text-white text-sm w-28"></div>
<div><label class="block text-[10px] text-on-dark-dim uppercase">Label</label>
<input type="text" name="task_label" placeholder="e.g. FUEL" required class="bg-dark border border-dark-border rounded px-2 py-1 text-white text-sm w-32"></div>
<div><label class="block text-[10px] text-on-dark-dim uppercase">Unit</label>
<input type="text" name="task_unit" placeholder="e.g. LITRE" required class="bg-dark border border-dark-border rounded px-2 py-1 text-white text-sm w-24"></div>
<button type="submit" class="bg-ny text-dark px-3 py-1 rounded text-xs font-bold">Add Task</button>
</form>
</div>

<!-- Section 2: Bulk Manage -->
<div class="bg-dark-card border border-dark-border rounded-lg p-4">
<h3 class="text-white font-bold text-sm mb-1">Bulk Manage</h3>
<p class="text-on-dark-dim text-[10px] mb-3">Select staff, then add/remove tasks with targets. Click "Apply" to update selected staff.</p>
<form method="POST" id="bulkForm">
<input type="hidden" name="form_type" value="bulk_manage">
<div class="flex flex-col lg:flex-row gap-4">
<!-- Staff Picker -->
<div class="lg:w-1/3">
<label class="dp-label block mb-2">Select Staff</label>
<div class="flex items-center gap-2 mb-2">
<input type="text" id="bulkStaffSearch" placeholder="Search staff..." class="dp-input flex-1" style="width:100%" oninput="filterBulkStaff()">
<button type="button" onclick="toggleAllBulkStaff()" class="dp-quick-btn text-[10px]" style="background:#111;border:1px solid #333;border-radius:6px;padding:4px 8px;color:#aaa;cursor:pointer">All</button>
</div>
<div id="bulkStaffList" class="space-y-1 max-h-48 overflow-y-auto" style="border:1px solid #333;border-radius:6px;padding:6px">
{% for m in members %}
<label class="flex items-center gap-2 px-2 py-1.5 rounded cursor-pointer hover:bg-white/5 transition bulk-staff-item" data-name="{{ m.name|lower }}">
<input type="checkbox" name="bulk_staff" value="{{ m.staff_id }}" class="bulk-staff-cb w-3.5 h-3.5 rounded accent-ny" onchange="updateBulkCount()">
<span class="text-xs text-on-dark-muted">{{ m.name }}</span>
</label>
{% endfor %}
</div>
<span class="text-on-dark-dim text-[10px] mt-1 block" id="bulkCount">0 selected</span>
</div>
<!-- Task Picker -->
<div class="lg:w-2/3">
<label class="dp-label block mb-2">Tasks & Targets</label>
<div class="space-y-2" id="bulkTaskList">
{% for t in all_tasks_db %}
<div class="flex items-center gap-3 bg-dark border border-dark-border rounded-lg px-3 py-2">
<label class="flex items-center gap-2 cursor-pointer flex-1">
<input type="checkbox" name="bulk_tasks" value="{{ t.task_key }}" class="bulk-task-cb w-3.5 h-3.5 rounded accent-ny">
<span class="text-xs text-on-dark-muted font-semibold">{{ t.label }}</span>
</label>
<div class="flex items-center gap-1">
<input type="number" name="bulk_target_{{ t.task_key }}" value="0" min="0" class="w-16 bg-dark border border-dark-border rounded px-2 py-1 text-[11px] font-mono text-white text-center">
<span class="text-on-dark-dim text-[9px]">{{ t.unit }}/day</span>
</div>
</div>
{% endfor %}
</div>
</div>
</div>
<div class="flex gap-2 mt-3">
<button type="submit" name="bulk_action" value="apply" class="bg-ny text-dark px-4 py-2 rounded-lg font-bold text-xs hover:bg-ny-bright transition">Apply (Add Tasks)</button>
<button type="submit" name="bulk_action" value="remove" class="bg-red/20 border border-red/40 text-red px-4 py-2 rounded-lg font-bold text-xs hover:bg-red/30 transition">Remove Tasks</button>
</div>
</form>
</div>

<!-- Section 3: Per-Staff Assignment -->
<div class="bg-dark-card border border-dark-border rounded-lg p-4">
<h3 class="text-white font-bold text-sm mb-1">Per-Staff Assignment</h3>
<p class="text-on-dark-dim text-[10px] mb-3">Check the tasks each staff performs, set their per-day targets, then click Save at the bottom.</p>
<div class="space-y-3" id="staffAssignment">
{% for m in members %}
<div class="border border-dark-border rounded p-3 staff-card" data-staff="{{ m.staff_id }}">
<div class="flex items-center gap-2 mb-2">
<div class="w-8 h-8 rounded bg-dark-border flex items-center justify-center overflow-hidden">
{% if m.picture %}<img src="/static/uploads/{{ m.picture }}" class="w-full h-full object-cover">{% else %}<span class="text-on-dark-muted font-mono text-xs font-bold">{{ m.name[:2] }}</span>{% endif %}
</div>
<div>
<p class="text-white font-semibold text-xs">{{ m.name }}</p>
<p class="text-on-dark-dim text-[10px]">{{ m.staff_id }}</p>
</div>
</div>
<div class="flex flex-wrap gap-x-4 gap-y-1 mt-2">
{% for t in all_tasks_db %}
<label class="flex items-center gap-1.5 cursor-pointer">
<input type="checkbox" class="staff-task-cb w-3.5 h-3.5 rounded border-dark-border bg-dark accent-ny"
data-staff="{{ m.staff_id }}" data-task="{{ t.task_key }}"
{% if t.task_key in m.tasks_list %}checked{% endif %}>
<span class="text-xs text-on-dark-muted">{{ t.label }}</span>
<input type="number" class="target-input w-14 bg-dark border border-dark-border rounded px-1.5 py-0.5 text-[10px] font-mono text-white"
data-staff="{{ m.staff_id }}" data-task="{{ t.task_key }}" value="{{ m.targets.get(t.task_key, 0) }}" min="0">
<span class="text-on-dark-dim text-[9px]">{{ t.unit }}/day</span>
</label>
{% endfor %}
</div>
</div>
{% endfor %}
</div>
</div>

<!-- Save All Button (bottom of per-staff) -->
<div class="mt-4">
<button type="button" onclick="saveAllStaff()" class="bg-ny text-dark px-6 py-3 rounded-lg font-bold text-sm shadow-lg hover:bg-ny-bright transition flex items-center gap-2">
<span class="material-symbols-outlined text-base">save</span>Save All Changes
</button>
</div>

</div>
<script>
function toggleEdit(id,label,unit){
document.getElementById('edit-task-form').classList.remove('hidden');
document.getElementById('edit_task_id').value=id;
document.getElementById('edit_label').value=label;
document.getElementById('edit_unit').value=unit;
}
function filterBulkStaff(){
var q=document.getElementById('bulkStaffSearch').value.toLowerCase();
document.querySelectorAll('.bulk-staff-item').forEach(function(el){
el.style.display=el.dataset.name.indexOf(q)!==-1?'':'none';
});
}
function toggleAllBulkStaff(){
var cbs=document.querySelectorAll('.bulk-staff-cb');
var allChecked=true;
cbs.forEach(function(cb){if(!cb.checked)allChecked=false});
cbs.forEach(function(cb){cb.checked=!allChecked});
updateBulkCount();
}
function updateBulkCount(){
var n=document.querySelectorAll('.bulk-staff-cb:checked').length;
document.getElementById('bulkCount').textContent=n+' selected';
}
function saveAllStaff(){
var data={form_type:'save_all_staff'};
var cards=document.querySelectorAll('.staff-card');
cards.forEach(function(card){
var sid=card.dataset.staff;
data['tasks_'+sid]=[];
var cbs=card.querySelectorAll('.staff-task-cb');
cbs.forEach(function(cb){
var tk=cb.dataset.task;
var target=card.querySelector('.target-input[data-staff="'+sid+'"][data-task="'+tk+'"]');
if(cb.checked){
data['tasks_'+sid].push(tk);
data['target_'+sid+'_'+tk]=target?target.value:'0';
}
});
});
var form=document.createElement('form');
form.method='POST';
form.action='/admin/tasks';
Object.keys(data).forEach(function(key){
if(Array.isArray(data[key])){
if(data[key].length===0){
var inp=document.createElement('input');
inp.type='hidden';inp.name=key;inp.value='';
form.appendChild(inp);
}else{
data[key].forEach(function(v){
var inp=document.createElement('input');
inp.type='hidden';inp.name=key;inp.value=v;
form.appendChild(inp);
});
}
}else{
var inp=document.createElement('input');
inp.type='hidden';inp.name=key;inp.value=data[key];
form.appendChild(inp);
}
});
document.body.appendChild(form);
form.submit();
}
</script>
'''

RECORDS_TPL = '''
<style>
.dp-wrap{display:flex;align-items:flex-end;gap:8px;flex-wrap:wrap}
.dp-group{display:flex;flex-direction:column;gap:3px}
.dp-label{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px}
.dp-input{background:#111;border:1px solid #333;border-radius:6px;padding:6px 10px;color:#fff;font-family:monospace;font-size:12px;outline:none;transition:border-color 0.2s;width:140px}
.dp-input:focus{border-color:#FFD700}
.dp-input::-webkit-calendar-picker-indicator{filter:invert(0.7);cursor:pointer}
.dp-select{background:#111;border:1px solid #333;border-radius:6px;padding:6px 10px;color:#fff;font-family:monospace;font-size:12px;outline:none;transition:border-color 0.2s}
.dp-select:focus{border-color:#FFD700}
.dp-go{background:#FFD700;color:#000;border:none;border-radius:6px;padding:6px 16px;font-size:11px;font-weight:700;cursor:pointer;transition:background 0.2s}
.dp-go:hover{background:#FFE44D}
.tab-btn{padding:8px 16px;font-size:12px;font-weight:600;border-bottom:2px solid transparent;color:#666;cursor:pointer;transition:all 0.2s;text-decoration:none}
.tab-btn.active{color:#FFD700;border-bottom-color:#FFD700}
.tab-btn:hover{color:#aaa}
.grid-cell{background:#111;border:1px solid #333;border-radius:4px;padding:4px 6px;color:#fff;font-family:monospace;font-size:11px;text-align:center;width:70px;outline:none;transition:border-color 0.2s}
.grid-cell:focus{border-color:#FFD700}
.grid-cell.has-data{background:#1A2A1A;border-color:#444}
.rec-check{accent-color:#FFD700;width:14px;height:14px}
.btn-del{background:rgba(255,0,0,0.15);border:1px solid rgba(255,0,0,0.3);color:#FF4444;padding:6px 14px;border-radius:6px;font-size:11px;font-weight:700;cursor:pointer;transition:background 0.2s}
.btn-del:hover{background:rgba(255,0,0,0.25)}
</style>
<div id="noDataModal" class="hidden fixed inset-0 z-50 flex items-center justify-center bg-black/60">
<div class="bg-dark-card border border-dark-border rounded-xl p-6 w-full max-w-xs mx-4 text-center">
<span class="material-symbols-outlined text-5xl text-on-dark-dim mb-3">info</span>
<h3 class="text-white font-bold text-base mb-1">No Data Found</h3>
<p class="text-on-dark-muted text-xs mb-4">There are no records for the selected date range to export.</p>
<button onclick="document.getElementById('noDataModal').classList.add('hidden')" class="bg-ny text-dark font-bold py-2 px-6 rounded text-sm hover:bg-ny-bright transition w-full">OK</button>
</div>
</div>
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-6xl mx-auto flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
<div>
<h2 class="text-xl sm:text-2xl font-bold text-white">Records <span class="text-red text-xs align-top ml-1">ADMIN</span></h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">View, edit, and delete form submissions</p>
</div>
{% if record_count > 0 %}
<a href="/generate?from={{ start_date }}&to={{ end_date }}{% if view_mode == 'grid' and grid_staff %}&staff={{ grid_staff }}{% endif %}" class="bg-dark-border text-on-dark px-3 py-2 rounded-lg text-xs sm:text-sm font-medium hover:bg-dark-hover transition flex items-center gap-1.5">
<span class="material-symbols-outlined text-base">download</span>Excel</a>
{% else %}
<button onclick="document.getElementById('noDataModal').classList.remove('hidden')" class="bg-dark-border text-on-dark px-3 py-2 rounded-lg text-xs sm:text-sm font-medium hover:bg-dark-hover transition flex items-center gap-1.5 opacity-50 cursor-not-allowed">
<span class="material-symbols-outlined text-base">download</span>Excel</button>
{% endif %}
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-6xl mx-auto">
<div class="flex gap-0 border-b border-dark-border mb-4">
<a href="/admin/records?view=list&start_date={{ start_date }}&end_date={{ end_date }}&staff={{ sel_staff }}" class="tab-btn {% if view_mode != 'grid' %}active{% endif %}">List View</a>
<a href="/admin/records?view=grid&start_date={{ start_date }}&end_date={{ end_date }}&staff={{ sel_staff }}" class="tab-btn {% if view_mode == 'grid' %}active{% endif %}">Staff Grid</a>
</div>
{% if view_mode == 'grid' %}
<form method="GET" class="dp-wrap mb-4">
<input type="hidden" name="view" value="grid">
<div class="dp-group">
<span class="dp-label">Staff</span>
<select name="staff" class="dp-select" required>
<option value="">Select Staff...</option>
{% for s in staff_order %}
<option value="{{ s }}" {% if s==grid_staff %}selected{% endif %}>{{ s }}</option>
{% endfor %}
</select>
</div>
<div class="dp-group">
<span class="dp-label">From</span>
<input type="date" name="start_date" value="{{ start_date }}" class="dp-input">
</div>
<div class="dp-group">
<span class="dp-label">To</span>
<input type="date" name="end_date" value="{{ end_date }}" class="dp-input">
</div>
<button class="dp-go">Load</button>
</form>
{% if grid_staff %}
<form method="POST" action="/admin/records/bulk-update" id="gridForm">
<input type="hidden" name="grid_staff" value="{{ grid_staff }}">
<input type="hidden" name="start_date" value="{{ start_date }}">
<input type="hidden" name="end_date" value="{{ end_date }}">
{% for yg in grid_groups %}
{% if show_year_heading %}
<div class="mb-4 border border-dark-border rounded-lg overflow-hidden">
<div class="grid-toggle flex items-center gap-2 px-4 py-3 bg-dark-card cursor-pointer hover:bg-dark-hover transition select-none" onclick="toggleSection('yr-{{ yg.yr }}')">
<span class="material-symbols-outlined text-ny text-base transition-transform" id="yr-{{ yg.yr }}-icon">expand_more</span>
<span class="text-white font-bold text-sm">{{ yg.yr }}</span>
<span class="text-on-dark-dim text-[10px] ml-1">({{ yg.months|length }} month{{ 's' if yg.months|length > 1 }})</span>
</div>
<div id="yr-{{ yg.yr }}">
{% endif %}
{% for mg in yg.months %}
{% if show_month_heading %}
<div class="{% if show_year_heading %}ml-4 mb-2 border-l border-dark-border{% else %}mb-4 border border-dark-border rounded-lg overflow-hidden{% endif %}">
{% if not show_year_heading %}
<div class="grid-toggle flex items-center gap-2 px-4 py-3 bg-dark-card cursor-pointer hover:bg-dark-hover transition select-none" onclick="toggleSection('mn-{{ yg.yr }}{{ mg.mn }}')">
<span class="material-symbols-outlined text-ny text-base transition-transform" id="mn-{{ yg.yr }}{{ mg.mn }}-icon">expand_more</span>
<span class="text-white font-bold text-sm">{{ mg.name }} {{ yg.yr }}</span>
<span class="text-on-dark-dim text-[10px] ml-1">({{ mg.dates|length }} day{{ 's' if mg.dates|length > 1 }})</span>
</div>
{% else %}
<div class="grid-toggle flex items-center gap-2 px-4 py-2 cursor-pointer hover:bg-white/5 transition select-none" onclick="toggleSection('mn-{{ yg.yr }}{{ mg.mn }}')">
<span class="material-symbols-outlined text-ny text-sm transition-transform" id="mn-{{ yg.yr }}{{ mg.mn }}-icon">expand_more</span>
<span class="text-on-dark font-semibold text-xs">{{ mg.name }}</span>
<span class="text-on-dark-dim text-[10px] ml-1">({{ mg.dates|length }} day{{ 's' if mg.dates|length > 1 }})</span>
</div>
{% endif %}
<div id="mn-{{ yg.yr }}{{ mg.mn }}">
{% endif %}
<table class="w-full text-xs border-collapse">
<thead><tr class="border-b border-dark-border">
<th class="py-2 px-2 text-left text-on-dark-dim w-8"><input type="checkbox" class="rec-check grid-month-cb" data-yr="{{ yg.yr }}" data-mn="{{ mg.mn }}" onclick="toggleMonthSelect(this)"></th>
<th class="py-2 px-2 text-left text-on-dark-dim w-28">Date</th>
{% for tk in task_keys %}
<th class="py-2 px-2 text-center text-on-dark-dim">{{ task_labels.get(tk, tk|upper) }}</th>
{% endfor %}
</tr></thead>
<tbody>
{% for date_key in mg.dates %}
<tr class="border-b border-dark-border hover:bg-dark-hover">
<td class="py-1 px-2 text-center"><input type="checkbox" name="delete_days" value="{{ date_key }}" class="rec-check grid-del-cb"></td>
<td class="py-1 px-2 font-mono font-bold text-on-dark">{{ date_key }}</td>
{% for tk in task_keys %}
{% set val = grid_data.get(date_key, {}).get(tk, '') %}
{% set compact = date_key.replace('/','') %}
<td class="py-1 px-2 text-center">
<input type="text" name="day_{{ compact }}_{{ tk }}" value="{{ val }}" class="grid-cell {% if val %}has-data{% endif %}" placeholder="-">
</td>
{% endfor %}
</tr>
{% endfor %}
</tbody>
</table>
{% if show_month_heading %}
</div>
</div>
{% endif %}
{% endfor %}
{% if show_year_heading %}
</div>
</div>
{% endif %}
{% endfor %}
<div class="flex items-center gap-3">
<button type="submit" class="bg-ny text-dark px-5 py-2 rounded-lg font-bold text-xs hover:bg-ny-bright transition">Save All Changes</button>
</div>
</form>
<form method="POST" action="/admin/records/bulk-delete-days" id="gridDeleteForm" class="mt-3">
<input type="hidden" name="grid_staff" value="{{ grid_staff }}">
<input type="hidden" name="start_date" value="{{ start_date }}">
<input type="hidden" name="end_date" value="{{ end_date }}">
<input type="hidden" name="delete_days" id="gridDeleteDays" value="">
<button type="button" onclick="submitGridDelete()" class="btn-del">Delete Selected Days</button>
</form>
{% elif not grid_staff %}
<p class="text-on-dark-dim text-sm py-8 text-center">Select a staff member and click Load.</p>
{% endif %}
{% else %}
<form method="GET" class="dp-wrap mb-4">
<input type="hidden" name="view" value="list">
<div class="dp-group">
<span class="dp-label">From</span>
<input type="date" name="start_date" value="{{ start_date }}" class="dp-input">
</div>
<div class="dp-group">
<span class="dp-label">To</span>
<input type="date" name="end_date" value="{{ end_date }}" class="dp-input">
</div>
<div class="dp-group">
<span class="dp-label">Staff</span>
<select name="staff" class="dp-select">
<option value="">All Staff</option>
{% for s in staff_order %}
<option value="{{ s }}" {% if s==sel_staff %}selected{% endif %}>{{ s }}</option>
{% endfor %}
</select>
</div>
<button class="dp-go">Filter</button>
</form>
<form method="POST" action="/admin/records/bulk-delete">
<div class="flex items-center gap-3 mb-3">
<label class="flex items-center gap-2 cursor-pointer"><input type="checkbox" id="listSelectAll" class="rec-check" onclick="toggleListSelectAll()"><span class="text-xs text-on-dark-dim">Select All</span></label>
<button type="submit" class="btn-del" id="listDelBtn" style="display:none" onclick="event.preventDefault();var f=this.form;showDeleteModal('Delete selected records?',function(yes){if(yes)f.submit()})">Delete Selected (<span id="listDelCount">0</span>)</button>
</div>
{% if not list_groups %}
<p class="text-on-dark-dim text-sm py-8 text-center">No records found</p>
{% endif %}
{% for yg in list_groups %}
{% if list_show_year %}
<div class="mb-4 border border-dark-border rounded-lg overflow-hidden">
<div class="grid-toggle flex items-center gap-2 px-4 py-3 bg-dark-card cursor-pointer hover:bg-dark-hover transition select-none" onclick="toggleSection('lyr-{{ yg.yr }}')">
<span class="material-symbols-outlined text-ny text-base transition-transform" id="lyr-{{ yg.yr }}-icon">expand_more</span>
<span class="text-white font-bold text-sm">{{ yg.yr }}</span>
<span class="text-on-dark-dim text-[10px] ml-1">({{ yg.months|length }} month{{ 's' if yg.months|length > 1 }})</span>
</div>
<div id="lyr-{{ yg.yr }}">
{% endif %}
{% for mg in yg.months %}
{% if list_show_month %}
<div class="{% if list_show_year %}ml-4 mb-2 border-l border-dark-border{% else %}mb-4 border border-dark-border rounded-lg overflow-hidden{% endif %}">
{% if not list_show_year %}
<div class="grid-toggle flex items-center gap-2 px-4 py-3 bg-dark-card cursor-pointer hover:bg-dark-hover transition select-none" onclick="toggleSection('lmn-{{ yg.yr }}{{ mg.mn }}')">
<span class="material-symbols-outlined text-ny text-base transition-transform" id="lmn-{{ yg.yr }}{{ mg.mn }}-icon">expand_more</span>
<span class="text-white font-bold text-sm">{{ mg.name }} {{ yg.yr }}</span>
<span class="text-on-dark-dim text-[10px] ml-1">({{ mg.records|length }} record{{ 's' if mg.records|length > 1 }})</span>
</div>
{% else %}
<div class="grid-toggle flex items-center gap-2 px-4 py-2 cursor-pointer hover:bg-white/5 transition select-none" onclick="toggleSection('lmn-{{ yg.yr }}{{ mg.mn }}')">
<span class="material-symbols-outlined text-ny text-sm transition-transform" id="lmn-{{ yg.yr }}{{ mg.mn }}-icon">expand_more</span>
<span class="text-on-dark font-semibold text-xs">{{ mg.name }}</span>
<span class="text-on-dark-dim text-[10px] ml-1">({{ mg.records|length }} record{{ 's' if mg.records|length > 1 }})</span>
</div>
{% endif %}
<div id="lmn-{{ yg.yr }}{{ mg.mn }}">
<div class="space-y-2 p-2">
{% else %}
<div class="space-y-2">
{% endif %}
{% for r in mg.records %}
<div class="bg-dark-card border border-dark-border rounded-lg p-3 flex items-center justify-between gap-3">
<div class="flex items-center gap-3 min-w-0">
<input type="checkbox" name="record_ids" value="{{ r.id }}" class="rec-check list-rec-cb" onchange="updateListDelCount()">
<span class="font-mono text-ny text-sm font-bold shrink-0">#{{ r.id }}</span>
<div class="min-w-0">
<p class="text-white text-sm truncate">{{ r.date }}</p>
<p class="text-on-dark-dim text-xs truncate">{{ r.q1 or '-' }} | {{ r.q2 or '-' }} | {{ r.q3 or '-' }} | {{ r.q4 or '-' }}</p>
</div>
</div>
<div class="flex gap-2 shrink-0 items-center">
<a href="/admin/records/edit/{{ r.id }}" class="inline-flex items-center justify-center px-3 py-1.5 text-ny text-xs font-semibold border border-ny/30 rounded hover:bg-ny/10 transition">Edit</a>
<form method="POST" action="/admin/records/delete/{{ r.id }}" style="display:inline" onsubmit="event.preventDefault();var form=this;showDeleteModal('Delete record #{{ r.id }}?',function(yes){if(yes)form.submit()})">
<button type="submit" class="inline-flex items-center justify-center px-3 py-1.5 text-red text-xs font-semibold border border-red/30 rounded hover:bg-red/10 transition cursor-pointer">Delete</button>
</form>
</div>
</div>
{% endfor %}
</div>
{% if list_show_month %}
</div>
</div>
{% endif %}
{% endfor %}
{% if list_show_year %}
</div>
</div>
{% endif %}
{% endfor %}
</form>
{% endif %}
</div>
<script>
function toggleListSelectAll(){var cbs=document.querySelectorAll('.list-rec-cb');var all=true;cbs.forEach(function(cb){if(!cb.checked)all=false});cbs.forEach(function(cb){cb.checked=!all});updateListDelCount()}
function updateListDelCount(){var n=document.querySelectorAll('.list-rec-cb:checked').length;document.getElementById('listDelCount').textContent=n;document.getElementById('listDelBtn').style.display=n>0?'inline-block':'none'}
function toggleSection(id){var el=document.getElementById(id);var icon=document.getElementById(id+'-icon');if(!el)return;var hidden=el.style.display==='none';el.style.display=hidden?'':'none';if(icon)icon.style.transform=hidden?'rotate(0deg)':'rotate(-90deg)'}
function toggleMonthSelect(cb){var yr=cb.getAttribute('data-yr');var mn=cb.getAttribute('data-mn');var cbs=document.querySelectorAll('.grid-del-cb');cbs.forEach(function(c){var parentTable=c.closest('table');if(parentTable){var monthCb=parentTable.querySelector('.grid-month-cb');if(monthCb&&monthCb.getAttribute('data-yr')===yr&&monthCb.getAttribute('data-mn')===mn){c.checked=cb.checked}}})}
function submitGridDelete(){var checked=document.querySelectorAll('.grid-del-cb:checked');if(checked.length===0){showDeleteModal('Select at least one day first.',function(){});return}showDeleteModal('Delete data for '+checked.length+' day(s)?',function(yes){if(!yes)return;var ids=[];checked.forEach(function(cb){ids.push(cb.value)});document.getElementById('gridDeleteDays').value=ids.join(',');document.getElementById('gridDeleteForm').submit()})}
</script>
'''

RECORD_EDIT_TPL = '''
<header class="bg-dark-card px-4 lg:px-8 py-4 sm:py-6 border-b border-dark-border">
<div class="max-w-3xl mx-auto">
<h2 class="text-xl sm:text-2xl font-bold text-white">Edit Record #{{ rec.id }}</h2>
<p class="text-on-dark-muted mt-1 text-xs sm:text-sm">Modify form submission data</p>
</div>
</header>
<div class="p-3 sm:p-4 lg:p-8 max-w-3xl mx-auto">
{% with messages = get_flashed_messages(with_categories=true) %}
{% if messages %}
{% for cat, msg in messages %}
<div class="px-4 py-3 rounded text-sm mb-4 {% if cat=='error' %}bg-red/10 border border-red/30 text-red{% else %}bg-ny/10 border border-ny/30 text-ny{% endif %}">{{ msg }}</div>
{% endfor %}
{% endif %}
{% endwith %}
<form method="POST" class="space-y-4">
<div>
<label class="block text-xs text-on-dark-muted uppercase mb-1">Date</label>
<input type="text" name="date" value="{{ rec.date }}" class="w-full bg-dark border border-dark-border rounded px-3 py-2 text-white font-mono text-sm">
</div>
{% set q_fields = ['q1','q2','q3','q4','q5','q6','q7'] %}
{% for i in range(task_keys|length) %}
{% if i < 7 %}
<div>
<label class="block text-xs text-on-dark-muted uppercase mb-1">{{ task_labels.get(task_keys[i], task_keys[i]|upper) }} ({{ task_units.get(task_keys[i],'') }})</label>
<input type="text" name="{{ q_fields[i] }}" value="{{ rec[q_fields[i]] or '' }}" placeholder="Staff - quantity pairs"
class="w-full bg-dark border border-dark-border rounded px-3 py-2 text-white font-mono text-sm">
<p class="text-on-dark-dim text-[10px] mt-0.5">Format: "Name - qty" separated by spaces</p>
</div>
{% endif %}
{% endfor %}
<div class="flex gap-3 pt-2">
<button type="submit" class="bg-ny text-dark px-6 py-2 rounded text-sm font-bold">Save Changes</button>
<a href="/admin/records" class="text-on-dark-muted text-sm py-2 hover:text-white">Cancel</a>
</div>
</form>
</div>
'''

# ============== 404 ERROR HANDLER ==============
NOT_FOUND_TPL = '''
<div class="text-center mb-6">
<img src="/static/images/shell-logo.png" alt="Shell" class="w-14 h-14 rounded-full object-cover mx-auto mb-3 border-2 border-ny/30">
<h1 class="text-2xl font-black text-white">OneTrack</h1>
<p class="text-on-dark-muted text-xs mt-0.5">Shell Bandar Mahkota Cheras</p>
</div>
<div class="bg-dark-card/90 backdrop-blur-sm border border-dark-border rounded-lg p-5 text-center">
<h1 class="text-5xl font-black text-ny mb-2">404</h1>
<h2 class="text-base font-bold text-white mb-2">Page Not Found</h2>
<p class="text-on-dark-muted text-xs mb-5">The page you're looking for doesn't exist or has been moved.</p>
<a href="/" class="block w-full bg-ny text-dark font-bold py-2.5 rounded text-sm hover:bg-ny-bright transition text-center">
GO TO DASHBOARD
</a>
<div class="text-center mt-3">
<button onclick="history.back()" class="text-ny text-xs hover:underline">Go Back</button>
</div>
</div>
<p class="text-center text-on-dark-dim text-[10px] mt-4">&copy; 2026 Shell Prosper Niaga</p>
'''

@app.errorhandler(404)
def page_not_found(e):
    return render_template_string(AUTH_LAYOUT, content=render_template_string(NOT_FOUND_TPL)), 404

# ============== NO INTERNET OVERLAY ==============
NO_INTERNET_TPL = '''
<div id="offline-overlay" class="hidden fixed inset-0 z-[9999] bg-dark/95 backdrop-blur-sm flex items-center justify-center p-4">
  <div class="w-full max-w-sm text-center">
    <div class="mb-6">
      <div class="relative inline-block">
        <span class="material-symbols-outlined text-7xl text-red">wifi_off</span>
        <span class="absolute -top-1 -right-1 w-4 h-4 bg-red rounded-full animate-ping"></span>
        <span class="absolute -top-1 -right-1 w-4 h-4 bg-red rounded-full"></span>
      </div>
    </div>
    <h2 class="text-xl font-bold text-white mb-2">No Internet Connection</h2>
    <p class="text-on-dark-muted text-sm mb-6">Please check your network settings and try again.</p>
    <div class="bg-dark-card border border-dark-border rounded-lg p-4 mb-6 text-left">
      <div class="space-y-2 text-xs font-mono">
        <div class="flex justify-between"><span class="text-on-dark-dim">WiFi / Cellular</span><span class="text-red">No Link</span></div>
        <div class="flex justify-between"><span class="text-on-dark-dim">Gateway</span><span class="text-red">Unreachable</span></div>
        <div class="flex justify-between"><span class="text-on-dark-dim">Server</span><span class="text-red">Offline</span></div>
      </div>
    </div>
    <button onclick="location.reload()" class="w-full bg-ny text-dark py-3 rounded-lg font-bold text-sm hover:bg-ny-bright transition flex items-center justify-center gap-2 mb-3">
      <span class="material-symbols-outlined text-base">refresh</span>
      Try Again
    </button>
    <a href="/" class="text-ny text-sm hover:underline inline-flex items-center gap-1">
      <span class="material-symbols-outlined text-base">dashboard</span>Go to Dashboard (Offline Mode)
    </a>
  </div>
</div>
<script>
(function(){
  var overlay=document.getElementById('offline-overlay');
  if(!overlay){overlay=document.createElement('div');overlay.id='offline-overlay';overlay.className='hidden fixed inset-0 z-[9999] bg-dark/95 backdrop-blur-sm flex items-center justify-center p-4';overlay.innerHTML='<div class="w-full max-w-sm text-center"><div class="mb-6"><span class="material-symbols-outlined text-7xl text-red">wifi_off</span></div><h2 class="text-xl font-bold text-white mb-2">No Internet Connection</h2><p class="text-on-dark-muted text-sm mb-6">Please check your network settings and try again.</p><button onclick="location.reload()" class="w-full bg-ny text-dark py-3 rounded-lg font-bold text-sm hover:bg-ny-bright transition flex items-center justify-center gap-2"><span class="material-symbols-outlined text-base">refresh</span>Try Again</button></div>';document.body.appendChild(overlay);}
  function update(){if(!navigator.onLine){overlay.classList.remove('hidden');}else{overlay.classList.add('hidden');}}
  window.addEventListener('online',update);window.addEventListener('offline',update);update();
})();
</script>
'''

if __name__ == "__main__":
    init_db()
    print("\n" + "="*60)
    print(" OneTrack - Shell Prosper Niaga Sales Leaderboard")
    print("="*60)
    print(" URL: http://localhost:5000")
    print(" LAN: http://192.168.0.48:5000")
    print("="*60 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)
