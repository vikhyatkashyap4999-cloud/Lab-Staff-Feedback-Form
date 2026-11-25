"""Lab Staff Quality Feedback Form.

Run:            python app.py                 (http://127.0.0.1:5000)
Import Excel:   python app.py import "Lab_Feedback_Demo_Data.xlsx"
Reset admin:    python app.py reset-admin
"""
import hashlib
import hmac
import io
import os
import re
import secrets
import sqlite3
import sys
import time
from datetime import datetime, timedelta
from functools import wraps

from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   send_file, session, url_for)
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash

from icons import icon
from workbook_utils import build_upload_workbook, write_sheet

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INSTANCE_DIR = os.path.join(BASE_DIR, "instance")
DB_PATH = os.path.join(INSTANCE_DIR, "feedback.db")
os.makedirs(INSTANCE_DIR, exist_ok=True)

PORTAL = {
    "title": "Lab Staff Quality Feedback Form",
    "subtitle": "Online Feedback Submission · Session Jul – Dec 2026",
    "org": "UPES · Laboratory Quality Cell",
    "short": "LQ",
}

STUDENT_QUESTIONS = [
    "Availability of lab staff during the session",
    "Helpfulness and support provided",
    "Technical knowledge and problem-solving assistance",
    "Professionalism and communication",
    "Readiness and organization of the laboratory",
    "Overall satisfaction with lab staff support",
]
FACULTY_QUESTIONS = [
    "Timely preparation of laboratory facilities and equipment",
    "Availability and responsiveness during lab sessions",
    "Technical competence in handling laboratory equipment",
    "Support provided for conducting experiments and practical sessions",
    "Maintenance and upkeep of laboratory resources",
    "Professionalism, communication, and cooperation",
    "Adherence to laboratory safety procedures",
]
QUESTIONS = {"student": STUDENT_QUESTIONS, "faculty": FACULTY_QUESTIONS}
MAX_Q = 7
RATING_SCALE = [
    (1, "Poor", "Staff/service did not meet expectations. Significant improvement needed."),
    (2, "Fair", "Below average performance. Some areas require notable improvement."),
    (3, "Average", "Meets basic expectations. Acceptable but room for improvement exists."),
    (4, "Good", "Above average. Demonstrates competence and reliability consistently."),
    (5, "Excellent", "Exceptional performance. Exceeds all expectations in every aspect."),
]
RATING_LABEL = {v: label for v, label, _ in RATING_SCALE}
ROLE_LABEL = {"student": "Student", "faculty": "Faculty / Lab Head / PL / Cluster Head"}
STEPS = ["Login", "Select Lab Staff", "Fill Feedback", "Preview & Declaration", "Submitted"]
SESSION_MINUTES = 30
STRONG_HASH = "scrypt"              # admin password
PW_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789"


def _secret_key():
    path = os.path.join(INSTANCE_DIR, "secret.key")
    if not os.path.exists(path):
        with open(path, "w") as f:
            f.write(secrets.token_hex(32))
    with open(path) as f:
        return f.read().strip()


app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.environ.get("LAB_SECRET_KEY") or _secret_key(),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("LAB_HTTPS") == "1",
    SESSION_COOKIE_NAME="lqf_session",
    PERMANENT_SESSION_LIFETIME=timedelta(minutes=SESSION_MINUTES),
    MAX_CONTENT_LENGTH=10 * 1024 * 1024,
)

# ------------------------------------------------------------------ database
# column "lab" holds the Lab Room No.
RESPONSES_SQL = f"""
CREATE TABLE IF NOT EXISTS responses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ref_no TEXT UNIQUE NOT NULL,
    rater_role TEXT NOT NULL, rater_id TEXT NOT NULL, rater_name TEXT, designation TEXT,
    staff_id TEXT NOT NULL, staff_name TEXT, lab TEXT NOT NULL,
    course_code TEXT NOT NULL DEFAULT '', course_name TEXT, cohort TEXT NOT NULL,
    faculty_name TEXT,
    {", ".join(f"q{i} INTEGER" for i in range(1, MAX_Q + 1))},
    average REAL, comments TEXT, submitted_at TEXT NOT NULL, client_ip TEXT,
    UNIQUE (rater_role, rater_id, staff_id, lab, course_code, cohort));"""
SCHEMA = f"""
CREATE TABLE IF NOT EXISTS assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id TEXT NOT NULL, student_name TEXT, lab TEXT NOT NULL,
    course_code TEXT NOT NULL DEFAULT '', course_name TEXT, cohort TEXT NOT NULL,
    faculty_id TEXT, faculty_name TEXT, staff_id TEXT NOT NULL, staff_name TEXT);
CREATE INDEX IF NOT EXISTS ix_asg_student ON assignments(student_id);
CREATE INDEX IF NOT EXISTS ix_asg_faculty ON assignments(faculty_id);
CREATE TABLE IF NOT EXISTS reviewer_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    reviewer_id TEXT NOT NULL, reviewer_name TEXT, designation TEXT,
    lab TEXT NOT NULL, staff_id TEXT NOT NULL, staff_name TEXT);
{RESPONSES_SQL}
CREATE TABLE IF NOT EXISTS admins (username TEXT PRIMARY KEY, password_hash TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT, actor TEXT, action TEXT, detail TEXT, ip TEXT);
"""


def connect():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def get_db():
    if "db" not in g:
        g.db = connect()
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def gen_password(n=10):
    return "".join(secrets.choice(PW_ALPHABET) for _ in range(n))


def set_admin_password(conn, password):
    conn.execute("INSERT OR REPLACE INTO admins VALUES ('admin', ?)",
                 (generate_password_hash(password, method=STRONG_HASH),))
    conn.commit()


def _columns(conn, table):
    return {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}


def migrate(conn):
    """Upgrade databases created before Course Code / Course Name existed."""
    if "course_code" not in _columns(conn, "assignments"):
        conn.execute("ALTER TABLE assignments ADD COLUMN course_code TEXT NOT NULL DEFAULT ''")
        conn.execute("ALTER TABLE assignments ADD COLUMN course_name TEXT")
    if "course_code" not in _columns(conn, "responses"):
        # the duplicate rule must include the course, so the table is rebuilt (rows kept)
        old = [c for c in _columns(conn, "responses")]
        conn.executescript(f"""
            BEGIN;
            ALTER TABLE responses RENAME TO responses_old;
            {RESPONSES_SQL}
            INSERT INTO responses ({", ".join(old)}) SELECT {", ".join(old)} FROM responses_old;
            DROP TABLE responses_old;
            COMMIT;""")


def init_db():
    conn = connect()
    conn.executescript(SCHEMA)
    migrate(conn)
    if not conn.execute("SELECT 1 FROM admins").fetchone():
        pwd = os.environ.get("LAB_ADMIN_PASSWORD") or gen_password(12)
        set_admin_password(conn, pwd)
        print("\n" + "=" * 60)
        print("  ADMIN ACCOUNT CREATED")
        print("  Username : admin")
        print(f"  Password : {pwd}")
        print("  Note it down now - it is not shown again.")
        print("  (Change it in Admin Panel, or run: python app.py reset-admin)")
        print("=" * 60 + "\n")
    conn.close()


def now_str():
    return datetime.now().strftime("%d-%m-%Y %H:%M:%S")


def client_ip():
    return request.remote_addr or ""


def audit(actor, action, detail=""):
    db = get_db()
    db.execute("INSERT INTO audit_log (at, actor, action, detail, ip) VALUES (?,?,?,?,?)",
               (now_str(), actor, action, detail[:500], client_ip()))
    db.commit()


# ------------------------------------------------------------- security bits
def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    return session["csrf"]


@app.before_request
def before():
    session.permanent = True
    if request.method == "POST":
        sent = request.form.get("csrf_token", "")
        if not sent or not hmac.compare_digest(sent, session.get("csrf", "")):
            # usually a page left open on a phone until the session expired
            is_admin = request.path.startswith("/admin")
            flash("Your session expired. Please " + ("log in again." if is_admin or not session.get("uid")
                                                   else "try again."), "error")
            if is_admin:
                return redirect(url_for("admin_login"))
            return redirect(url_for("dashboard") if session.get("uid") else url_for("login"))


@app.after_request
def security_headers(resp):
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; script-src 'self'; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com; img-src 'self' data:; "
        "form-action 'self'; frame-ancestors 'none'; base-uri 'self'")
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "same-origin"
    if not request.path.startswith("/static/"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


_FAILS = {}
LOCK_AFTER, LOCK_SECONDS = 5, 15 * 60


def locked_for(*keys):
    """Seconds left on the longest lock among keys (0 = not locked)."""
    now = time.time()
    return max([int(_FAILS.get(k, {}).get("until", 0) - now) for k in keys] + [0])


def note_fail(*keys):
    for k in keys:
        limit = LOCK_AFTER * 4 if k.startswith("ip:") else LOCK_AFTER
        rec = _FAILS.setdefault(k, {"n": 0, "until": 0})
        rec["n"] += 1
        if rec["n"] >= limit:
            rec["n"], rec["until"] = 0, time.time() + LOCK_SECONDS


def clear_fail(key):
    _FAILS.pop(key, None)


_DUMMY_HASH = generate_password_hash("not-a-real-password", method=STRONG_HASH)


def new_captcha(ns):
    a, b = secrets.randbelow(9) + 1, secrets.randbelow(9) + 1
    session[f"cap_{ns}"] = str(a + b)
    return [(ch, f"rot{secrets.randbelow(4) + 1}") for ch in f"{a}+{b}=?"]


def captcha_ok(ns, answer):
    expected = session.pop(f"cap_{ns}", None)
    return expected is not None and answer.strip() == expected


@app.context_processor
def inject_globals():
    return {"portal": PORTAL, "csrf_token": csrf_token, "now_str": now_str(),
            "steps": STEPS, "rating_scale": RATING_SCALE, "role_label": ROLE_LABEL,
            "session_seconds": SESSION_MINUTES * 60, "icon": icon}


@app.template_filter("initials")
def initials_filter(name):
    parts = [p for p in re.split(r"[\s.]+", str(name or "")) if p and p.lower() not in ("dr", "prof", "mr", "ms", "mrs")]
    return ("".join(p[0] for p in parts[:2]) or "?").upper()


@app.template_filter("hue")
def hue_filter(value):
    return int(hashlib.md5(str(value).encode()).hexdigest()[:4], 16) % 360


def login_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if not session.get("uid"):
            flash("Your session has expired or you are not logged in. Please log in again.", "error")
            return redirect(url_for("login"))
        return fn(*a, **kw)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        if not session.get("admin"):
            return redirect(url_for("admin_login"))
        return fn(*a, **kw)
    return wrapper


# ------------------------------------------------------------------ targets
def target_key(staff_id, lab, course_code, cohort):
    return hashlib.sha256(f"{staff_id}\x1f{lab}\x1f{course_code}\x1f{cohort}".encode()).hexdigest()[:20]


def get_targets(role, uid):
    """One feedback target per lab staff + lab room + course + cohort assigned to this person."""
    db = get_db()
    if role == "student":
        rows = db.execute(
            """SELECT staff_id, MAX(staff_name) AS staff_name, lab, course_code, MAX(course_name) AS course_name,
                      cohort, GROUP_CONCAT(DISTINCT faculty_name) AS faculty_name,
                      '' AS designation, 0 AS students
               FROM assignments WHERE student_id = ?
               GROUP BY staff_id, lab, course_code, cohort
               ORDER BY course_code, lab, cohort, staff_name""", (uid,)).fetchall()
    else:
        rows = db.execute(
            """SELECT staff_id, MAX(staff_name) AS staff_name, lab, course_code, MAX(course_name) AS course_name,
                      cohort, MAX(faculty_name) AS faculty_name, 'Faculty' AS designation,
                      COUNT(DISTINCT student_id) AS students
               FROM assignments WHERE faculty_id = ?
               GROUP BY staff_id, lab, course_code, cohort
               UNION ALL
               SELECT staff_id, MAX(staff_name), lab, '' AS course_code, '' AS course_name, '' AS cohort,
                      MAX(reviewer_name), MAX(designation), 0
               FROM reviewer_assignments WHERE reviewer_id = ?
               GROUP BY staff_id, lab
               ORDER BY 4, 3, 6, 2""", (uid, uid)).fetchall()
    done = {(r["staff_id"], r["lab"], r["course_code"], r["cohort"]): r for r in db.execute(
        "SELECT staff_id, lab, course_code, cohort, ref_no, submitted_at FROM responses "
        "WHERE rater_role = ? AND rater_id = ?", (role, uid))}
    targets = []
    for r in rows:
        t = dict(r)
        t["key"] = target_key(t["staff_id"], t["lab"], t["course_code"], t["cohort"])
        sub = done.get((t["staff_id"], t["lab"], t["course_code"], t["cohort"]))
        t["ref_no"] = sub["ref_no"] if sub else None
        t["submitted_at"] = sub["submitted_at"] if sub else None
        targets.append(t)
    return targets


def find_target(key):
    for t in get_targets(session["role"], session["uid"]):
        if hmac.compare_digest(t["key"], key):
            return t
    abort(404)


def display_name(role, uid):
    db = get_db()
    if role == "student":
        r = db.execute("SELECT student_name FROM assignments WHERE student_id=? LIMIT 1", (uid,)).fetchone()
    else:
        r = db.execute("SELECT faculty_name FROM assignments WHERE faculty_id=? LIMIT 1", (uid,)).fetchone() \
            or db.execute("SELECT reviewer_name FROM reviewer_assignments WHERE reviewer_id=? LIMIT 1", (uid,)).fetchone()
    return (r and r[0]) or uid


def id_exists(role, uid):
    db = get_db()
    if role == "student":
        return db.execute("SELECT 1 FROM assignments WHERE student_id=? LIMIT 1", (uid,)).fetchone() is not None
    return (db.execute("SELECT 1 FROM assignments WHERE faculty_id=? LIMIT 1", (uid,)).fetchone()
            or db.execute("SELECT 1 FROM reviewer_assignments WHERE reviewer_id=? LIMIT 1", (uid,)).fetchone()) is not None


# -------------------------------------------------------------- user routes
@app.route("/")
def index():
    if session.get("uid"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET" and session.get("uid"):
        return redirect(url_for("dashboard"))
    form = {"role": "", "user_id": "", "ask_role": False}
    if request.method == "POST":
        uid = request.form.get("user_id", "").strip()[:50]
        chosen = request.form.get("role", "")
        form = {"role": chosen, "user_id": uid, "ask_role": False}
        ip_key = f"ip:{client_ip()}"
        wait = locked_for(ip_key)
        if wait:
            flash(f"Too many failed attempts. Please try again after {wait // 60 + 1} minute(s).", "error")
        elif not captcha_ok("login", request.form.get("captcha", "")):
            flash("Security code is incorrect. Please solve the new code shown.", "error")
        else:
            # detect login type from the uploaded lists
            roles = [r for r in ("student", "faculty") if uid and id_exists(r, uid)]
            if len(roles) == 2 and chosen in roles:
                roles = [chosen]
            if len(roles) == 1:
                role = roles[0]
                session.clear()
                session.permanent = True
                session.update(uid=uid, role=role, name=display_name(role, uid), login_at=now_str())
                audit(f"{role}:{uid}", "login")
                return redirect(url_for("dashboard"))
            if len(roles) == 2:
                form["ask_role"] = True
                flash("This ID is registered both as a Student and as Faculty. Please choose how you want to "
                      "log in, solve the new security code and click Login again.", "info")
            else:
                note_fail(ip_key)
                flash("This ID is not registered. Please check your Student ID / Faculty ID.", "error")
    return render_template("login.html", step=1, captcha=new_captcha("login"), form=form)


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    flash("You have been logged out successfully.", "success")
    return redirect(url_for("login"))


@app.route("/instructions")
def instructions():
    return render_template("instructions.html", student_q=STUDENT_QUESTIONS, faculty_q=FACULTY_QUESTIONS)


@app.route("/dashboard")
@login_required
def dashboard():
    targets = get_targets(session["role"], session["uid"])
    done = sum(1 for t in targets if t["ref_no"])
    return render_template("dashboard.html", step=2, targets=targets, done=done)


@app.route("/feedback/<key>", methods=["GET", "POST"])
@login_required
def feedback_form(key):
    t = find_target(key)
    if t["ref_no"]:
        flash("You have already submitted feedback for this lab staff. It cannot be submitted again.", "info")
        return redirect(url_for("feedback_summary", ref_no=t["ref_no"]))
    questions = QUESTIONS[session["role"]]
    draft = session.get("draft") if (session.get("draft") or {}).get("key") == key else None
    answers = draft["answers"] if draft else [None] * len(questions)
    comments = draft["comments"] if draft else ""
    missing = []
    if request.method == "POST":
        answers, missing = [], []
        for i in range(1, len(questions) + 1):
            v = request.form.get(f"q{i}", "")
            if v in ("1", "2", "3", "4", "5"):
                answers.append(int(v))
            else:
                answers.append(None)
                missing.append(i)
        comments = request.form.get("comments", "").strip()[:1000]
        if not missing:
            session["draft"] = {"key": key, "answers": answers, "comments": comments}
            return redirect(url_for("feedback_preview", key=key))
        flash(f"Please give a rating for question(s): {', '.join(map(str, missing))}.", "error")
    return render_template("form.html", step=3, t=t, questions=questions, answers=answers,
                           comments=comments, missing=missing)


@app.route("/feedback/<key>/preview")
@login_required
def feedback_preview(key):
    t = find_target(key)
    if t["ref_no"]:
        return redirect(url_for("feedback_summary", ref_no=t["ref_no"]))
    draft = session.get("draft")
    if not draft or draft.get("key") != key:
        return redirect(url_for("feedback_form", key=key))
    questions = QUESTIONS[session["role"]]
    avg = round(sum(draft["answers"]) / len(draft["answers"]), 2)
    rows = [(i + 1, q, a, RATING_LABEL[a]) for i, (q, a) in enumerate(zip(questions, draft["answers"]))]
    return render_template("preview.html", step=4, t=t, rows=rows, avg=avg, comments=draft["comments"])


@app.route("/feedback/<key>/submit", methods=["POST"])
@login_required
def feedback_submit(key):
    t = find_target(key)
    draft = session.get("draft")
    if t["ref_no"]:
        flash("You have already submitted feedback for this lab staff.", "info")
        return redirect(url_for("feedback_summary", ref_no=t["ref_no"]))
    if not draft or draft.get("key") != key:
        flash("Please fill the feedback form first.", "error")
        return redirect(url_for("feedback_form", key=key))
    if request.form.get("declaration") != "yes":
        flash("Please tick the declaration before final submission.", "error")
        return redirect(url_for("feedback_preview", key=key))
    answers = draft["answers"]
    qs = answers + [None] * (MAX_Q - len(answers))
    ref_no = f"LSF{datetime.now():%y%m%d}{''.join(secrets.choice('ABCDEFGHJKLMNPQRSTUVWXYZ23456789') for _ in range(6))}"
    role, uid = session["role"], session["uid"]
    db = get_db()
    try:
        db.execute(
            f"""INSERT INTO responses (ref_no, rater_role, rater_id, rater_name, designation, staff_id,
                    staff_name, lab, course_code, course_name, cohort, faculty_name,
                    {", ".join(f"q{i}" for i in range(1, MAX_Q + 1))}, average, comments, submitted_at, client_ip)
                VALUES ({", ".join("?" * (16 + MAX_Q))})""",
            (ref_no, role, uid, session["name"], t["designation"] if role == "faculty" else "",
             t["staff_id"], t["staff_name"], t["lab"], t["course_code"], t["course_name"], t["cohort"],
             t["faculty_name"], *qs,
             round(sum(answers) / len(answers), 2), draft["comments"], now_str(), client_ip()))
        db.commit()
    except sqlite3.IntegrityError:
        db.rollback()
        flash("Feedback for this lab staff has already been recorded. Duplicate submission blocked.", "info")
        return redirect(url_for("dashboard"))
    session.pop("draft", None)
    audit(f"{role}:{uid}", "submit",
          f"{ref_no} staff={t['staff_id']} room={t['lab']} course={t['course_code']} cohort={t['cohort']}")
    flash("Your feedback has been submitted successfully.", "success")
    return redirect(url_for("feedback_summary", ref_no=ref_no))


@app.route("/summary/<ref_no>")
@login_required
def feedback_summary(ref_no):
    r = get_db().execute("SELECT * FROM responses WHERE ref_no=? AND rater_role=? AND rater_id=?",
                         (ref_no, session["role"], session["uid"])).fetchone()
    if not r:
        abort(404)
    questions = QUESTIONS[r["rater_role"]]
    rows = [(i + 1, q, r[f"q{i + 1}"], RATING_LABEL.get(r[f"q{i + 1}"], "")) for i, q in enumerate(questions)]
    pending = [t for t in get_targets(session["role"], session["uid"]) if not t["ref_no"]]
    return render_template("summary.html", step=5, r=r, rows=rows, pending=pending)


# ------------------------------------------------------------------- import
def norm(h):
    return re.sub(r"[^a-z0-9]", "", str(h or "").lower())


def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()[:200]


ROOM_ALIASES = {"labroomno": "lab", "labroomnumber": "lab", "labroom": "lab", "roomno": "lab", "lab": "lab"}
ASSIGN_MAP = {"studentid": "student_id", "studentname": "student_name", **ROOM_ALIASES,
              "coursecode": "course_code", "coursename": "course_name",
              "cohort": "cohort", "facultyid": "faculty_id", "facultyname": "faculty_name",
              "labstaffid": "staff_id", "labstaffname": "staff_name"}
REVIEWER_MAP = {"reviewerid": "reviewer_id", "reviewername": "reviewer_name", "designation": "designation",
                **ROOM_ALIASES, "labstaffid": "staff_id", "labstaffname": "staff_name"}
FIELD_LABEL = {"student_id": "Student ID", "student_name": "Student Name", "lab": "Lab Room No.",
               "course_code": "Course Code", "course_name": "Course Name", "cohort": "Cohort",
               "faculty_id": "Faculty ID", "faculty_name": "Faculty Name", "staff_id": "Lab Staff ID",
               "staff_name": "Lab Staff Name", "reviewer_id": "Reviewer ID", "reviewer_name": "Reviewer Name",
               "designation": "Designation"}


class ImportProblem(Exception):
    pass


def read_table(ws, header_map, needed_cols, required_fields, label):
    out, errors, idx = [], [], None
    for rnum, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if idx is None:
            normed = [norm(c) for c in row]
            found = {header_map[n]: i for i, n in enumerate(normed) if n in header_map}
            if all(c in found for c in needed_cols):
                idx = found
            elif rnum >= 10:
                break
            continue
        rec = {f: clean(row[i]) if i < len(row) else "" for f, i in idx.items()}
        if not any(rec.values()):
            continue
        missing = [FIELD_LABEL[f] for f in required_fields if not rec.get(f)]
        if missing:
            errors.append(f"{label} row {rnum}: missing {', '.join(missing)}")
        else:
            out.append(rec)
    if idx is None:
        errors.append(f"Sheet '{ws.title}': header row not found. Expected columns: "
                      + ", ".join(FIELD_LABEL[c] for c in needed_cols))
    return out, errors


def import_workbook(fileobj, conn):
    try:
        wb = load_workbook(fileobj, read_only=True, data_only=True)
    except Exception:
        raise ImportProblem(["The file could not be read. Please upload a valid .xlsx Excel file."])
    sheets = {norm(n): wb[n] for n in wb.sheetnames}
    ws = sheets.get("assignments") or wb[wb.sheetnames[0]]
    assigns, errors = read_table(ws, ASSIGN_MAP, ["student_id", "student_name", "lab", "course_code", "course_name",
                                                  "cohort", "faculty_id", "faculty_name", "staff_id", "staff_name"],
                                 ["student_id", "student_name", "lab", "course_code", "cohort", "staff_id",
                                  "staff_name"], "Assignments")
    reviewers = []
    if "reviewers" in sheets:
        reviewers, e = read_table(sheets["reviewers"], REVIEWER_MAP,
                                  ["reviewer_id", "reviewer_name", "lab", "staff_id", "staff_name"],
                                  ["reviewer_id", "reviewer_name", "lab", "staff_id", "staff_name"], "Reviewers")
        errors += e
    if not assigns and not errors:
        errors.append("No assignment rows found in the Assignments sheet.")
    if errors:
        raise ImportProblem(errors)

    try:
        conn.execute("BEGIN")
        conn.execute("DELETE FROM assignments")
        conn.executemany(
            "INSERT INTO assignments (student_id, student_name, lab, course_code, course_name, cohort, faculty_id, "
            "faculty_name, staff_id, staff_name) VALUES (:student_id, :student_name, :lab, :course_code, :course_name, "
            ":cohort, :faculty_id, :faculty_name, :staff_id, :staff_name)", assigns)
        conn.execute("DELETE FROM reviewer_assignments")
        conn.executemany(
            "INSERT INTO reviewer_assignments (reviewer_id, reviewer_name, designation, lab, staff_id, staff_name) "
            "VALUES (?,?,?,?,?,?)",
            [(r["reviewer_id"], r["reviewer_name"], r.get("designation") or "Lab Head", r["lab"],
              r["staff_id"], r["staff_name"]) for r in reviewers])
        conn.commit()
    except Exception:
        conn.rollback()
        raise

    return {"assignments": len(assigns), "reviewers": len(reviewers),
            "students": len({a["student_id"] for a in assigns}),
            "faculty": len({a["faculty_id"] for a in assigns if a["faculty_id"]} | {r["reviewer_id"] for r in reviewers})}


# -------------------------------------------------------------- admin routes
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()[:50]
        pwd = request.form.get("password", "")
        acct_key, ip_key = f"admin:{username.lower()}", f"ip:{client_ip()}"
        wait = locked_for(acct_key, ip_key)
        if wait:
            flash(f"Too many failed attempts. Try again after {wait // 60 + 1} minute(s).", "error")
        elif not captcha_ok("admin", request.form.get("captcha", "")):
            flash("Security code is incorrect.", "error")
        else:
            row = get_db().execute("SELECT * FROM admins WHERE username=?", (username,)).fetchone()
            if check_password_hash(row["password_hash"] if row else _DUMMY_HASH, pwd) and row:
                clear_fail(acct_key)
                session.clear()
                session.permanent = True
                session["admin"] = username
                audit(f"admin:{username}", "login")
                return redirect(url_for("admin_panel"))
            note_fail(acct_key, ip_key)
            flash("Invalid admin credentials.", "error")
    return render_template("admin_login.html", captcha=new_captcha("admin"))


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    session.clear()
    flash("Admin logged out.", "success")
    return redirect(url_for("admin_login"))


def staff_summary(db):
    staff = {}
    new = lambda sid, name, labs="", courses="": {"staff_id": sid, "name": name, "labs": labs, "courses": courses,
                                                  "exp_s": 0, "exp_f": 0, "got_s": 0, "got_f": 0,
                                                  "avg_s": None, "avg_f": None}
    for r in db.execute("""SELECT staff_id, MAX(staff_name) AS name, GROUP_CONCAT(DISTINCT lab) AS labs,
                                  GROUP_CONCAT(DISTINCT course_code) AS courses,
                                  COUNT(DISTINCT student_id || '|' || lab || '|' || course_code || '|' || cohort) AS exp_s
                           FROM assignments GROUP BY staff_id"""):
        staff[r["staff_id"]] = new(r["staff_id"], r["name"], r["labs"], r["courses"] or "")
        staff[r["staff_id"]]["exp_s"] = r["exp_s"]
    fac_targets = db.execute("""SELECT staff_id, MAX(staff_name) AS name, MAX(lab) AS lab, COUNT(*) AS n FROM (
                                   SELECT DISTINCT faculty_id AS rid, staff_id, staff_name, lab, course_code, cohort
                                   FROM assignments WHERE faculty_id <> ''
                                   UNION
                                   SELECT DISTINCT reviewer_id, staff_id, staff_name, lab, '', ''
                                   FROM reviewer_assignments) GROUP BY staff_id""").fetchall()
    for r in fac_targets:
        staff.setdefault(r["staff_id"], new(r["staff_id"], r["name"], r["lab"]))["exp_f"] = r["n"]
    for r in db.execute("""SELECT staff_id, MAX(staff_name) AS name, rater_role, COUNT(*) AS n,
                                  ROUND(AVG(average), 2) AS avg
                           FROM responses GROUP BY staff_id, rater_role"""):
        s = staff.setdefault(r["staff_id"], new(r["staff_id"], r["name"]))
        sfx = "s" if r["rater_role"] == "student" else "f"
        s[f"got_{sfx}"], s[f"avg_{sfx}"] = r["n"], r["avg"]
    return sorted(staff.values(), key=lambda s: str(s["staff_id"]))


def pending_rows(db):
    return db.execute("""
        SELECT 'Student' AS role, a.student_id AS rid, MAX(a.student_name) AS rname, a.lab, a.course_code,
               MAX(a.course_name) AS course_name, a.cohort, a.staff_id, MAX(a.staff_name) AS staff_name
        FROM assignments a
        WHERE NOT EXISTS (SELECT 1 FROM responses r WHERE r.rater_role='student' AND r.rater_id=a.student_id
                          AND r.staff_id=a.staff_id AND r.lab=a.lab AND r.course_code=a.course_code
                          AND r.cohort=a.cohort)
        GROUP BY a.student_id, a.staff_id, a.lab, a.course_code, a.cohort
        UNION ALL
        SELECT 'Faculty', a.faculty_id, MAX(a.faculty_name), a.lab, a.course_code, MAX(a.course_name), a.cohort,
               a.staff_id, MAX(a.staff_name)
        FROM assignments a
        WHERE a.faculty_id <> '' AND NOT EXISTS (SELECT 1 FROM responses r WHERE r.rater_role='faculty'
                          AND r.rater_id=a.faculty_id AND r.staff_id=a.staff_id AND r.lab=a.lab
                          AND r.course_code=a.course_code AND r.cohort=a.cohort)
        GROUP BY a.faculty_id, a.staff_id, a.lab, a.course_code, a.cohort
        UNION ALL
        SELECT MAX(v.designation), v.reviewer_id, MAX(v.reviewer_name), v.lab, '', '', '', v.staff_id,
               MAX(v.staff_name)
        FROM reviewer_assignments v
        WHERE NOT EXISTS (SELECT 1 FROM responses r WHERE r.rater_role='faculty' AND r.rater_id=v.reviewer_id
                          AND r.staff_id=v.staff_id AND r.lab=v.lab AND r.course_code='' AND r.cohort='')
        GROUP BY v.reviewer_id, v.staff_id, v.lab
        ORDER BY 1, 2""").fetchall()


@app.route("/admin")
@admin_required
def admin_panel():
    db = get_db()
    one = lambda sql: db.execute(sql).fetchone()[0]
    stats = {
        "students": one("SELECT COUNT(DISTINCT student_id) FROM assignments"),
        "faculty": one("SELECT COUNT(*) FROM (SELECT faculty_id FROM assignments WHERE faculty_id<>'' "
                       "UNION SELECT reviewer_id FROM reviewer_assignments)"),
        "staff": one("SELECT COUNT(DISTINCT staff_id) FROM assignments"),
        "rows": one("SELECT COUNT(*) FROM assignments"),
        "got_s": one("SELECT COUNT(*) FROM responses WHERE rater_role='student'"),
        "got_f": one("SELECT COUNT(*) FROM responses WHERE rater_role='faculty'"),
        "exp_s": one("SELECT COUNT(*) FROM (SELECT DISTINCT student_id, staff_id, lab, course_code, cohort "
                     "FROM assignments)"),
        "exp_f": one("SELECT COUNT(*) FROM (SELECT DISTINCT faculty_id, staff_id, lab, course_code, cohort "
                     "FROM assignments WHERE faculty_id<>'' UNION SELECT DISTINCT reviewer_id, staff_id, lab, '', '' "
                     "FROM reviewer_assignments)"),
    }
    recent = db.execute("SELECT ref_no, submitted_at, rater_role, rater_id, rater_name, staff_name, lab, course_code, "
                        "cohort, average FROM responses ORDER BY id DESC LIMIT 10").fetchall()
    return render_template("admin.html", stats=stats, staff=staff_summary(db), recent=recent)


@app.route("/admin/responses")
@admin_required
def admin_responses():
    q = request.args.get("q", "").strip()[:100]
    sql = ("SELECT id, ref_no, submitted_at, rater_role, rater_id, rater_name, designation, staff_id, staff_name, "
           "lab, course_code, course_name, cohort, average FROM responses")
    args = ()
    if q:
        like = f"%{q}%"
        sql += (" WHERE rater_id LIKE ? OR rater_name LIKE ? OR staff_id LIKE ? OR staff_name LIKE ? "
                "OR ref_no LIKE ? OR lab LIKE ? OR cohort LIKE ? OR course_code LIKE ? OR course_name LIKE ?")
        args = (like,) * 9
    rows = get_db().execute(sql + " ORDER BY id DESC LIMIT 1000", args).fetchall()
    total = get_db().execute("SELECT COUNT(*) FROM responses").fetchone()[0]
    return render_template("admin_responses.html", rows=rows, q=q, total=total)


@app.route("/admin/responses/<int:rid>/delete", methods=["POST"])
@admin_required
def admin_delete_response(rid):
    db = get_db()
    r = db.execute("SELECT ref_no, rater_role, rater_id, staff_id, lab, course_code, cohort FROM responses WHERE id=?",
                   (rid,)).fetchone()
    if r:
        db.execute("DELETE FROM responses WHERE id=?", (rid,))
        db.commit()
        audit(f"admin:{session['admin']}", "delete response",
              f"{r['ref_no']} {r['rater_role']}:{r['rater_id']} staff={r['staff_id']} room={r['lab']} "
              f"course={r['course_code']} cohort={r['cohort']}")
        flash(f"Response {r['ref_no']} deleted. {r['rater_id']} can now submit feedback for this lab staff again.",
              "success")
    return redirect(url_for("admin_responses", q=request.form.get("q", "")))


@app.route("/admin/responses/delete-all", methods=["POST"])
@admin_required
def admin_delete_all():
    if request.form.get("confirm_text", "").strip() != "DELETE":
        flash("Type DELETE (in capital letters) to confirm deleting all responses.", "error")
        return redirect(url_for("admin_responses"))
    db = get_db()
    n = db.execute("SELECT COUNT(*) FROM responses").fetchone()[0]
    db.execute("DELETE FROM responses")
    db.commit()
    audit(f"admin:{session['admin']}", "delete all responses", f"{n} rows")
    flash(f"All {n} responses deleted. Everyone can submit feedback again.", "success")
    return redirect(url_for("admin_responses"))


@app.route("/admin/upload", methods=["POST"])
@admin_required
def admin_upload():
    f = request.files.get("datafile")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Please choose an .xlsx Excel file to upload.", "error")
        return redirect(url_for("admin_panel"))
    try:
        res = import_workbook(io.BytesIO(f.read()), get_db())
    except ImportProblem as e:
        problems = e.args[0]
        flash("Upload rejected - nothing was changed. Please correct these problems:", "error")
        for p in problems[:25]:
            flash(p, "error")
        if len(problems) > 25:
            flash(f"... and {len(problems) - 25} more.", "error")
        return redirect(url_for("admin_panel"))
    audit(f"admin:{session['admin']}", "upload", f"{f.filename}: {res}")
    msg = (f"Upload successful: {res['assignments']} assignment rows, {res['reviewers']} reviewer rows - "
           f"{res['students']} students and {res['faculty']} faculty/reviewers can now log in with their ID.")
    flash(msg, "success")
    return redirect(url_for("admin_panel"))


def _xlsx_response(wb, name):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/admin/download/responses")
@admin_required
def download_responses():
    db = get_db()
    wb = Workbook()
    stamp = f"{datetime.now():%Y%m%d_%H%M}"
    for idx, (role, qs, title) in enumerate([("student", STUDENT_QUESTIONS, "Student Responses"),
                                             ("faculty", FACULTY_QUESTIONS, "Faculty Responses")]):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = title
        who = "Student" if role == "student" else "Reviewer"
        headers = ["Ref No", "Submitted At", f"{who} ID", f"{who} Name"]
        headers += ["Designation"] if role == "faculty" else []
        headers += ["Lab Room No.", "Course Code", "Course Name", "Cohort", "Faculty Name", "Lab Staff ID",
                    "Lab Staff Name"]
        headers += [f"Q{i + 1}. {q}" for i, q in enumerate(qs)] + ["Average", "Remarks"]
        rows = []
        for r in db.execute("SELECT * FROM responses WHERE rater_role=? ORDER BY id", (role,)):
            row = [r["ref_no"], r["submitted_at"], r["rater_id"], r["rater_name"]]
            row += [r["designation"]] if role == "faculty" else []
            row += [r["lab"], r["course_code"] or "All courses", r["course_name"], r["cohort"] or "All cohorts",
                    r["faculty_name"], r["staff_id"], r["staff_name"]]
            row += [r[f"q{i + 1}"] for i in range(len(qs))] + [r["average"], r["comments"]]
            rows.append(row)
        write_sheet(ws, headers, rows, max_width=40)

    ws = wb.create_sheet("Staff Summary")
    write_sheet(ws, ["Lab Staff ID", "Lab Staff Name", "Lab Room Nos.", "Course Codes", "Student Feedback Expected",
                     "Student Feedback Received", "Student Avg (out of 5)", "Faculty/Lab Head Feedback Expected",
                     "Faculty/Lab Head Feedback Received", "Faculty/Lab Head Avg (out of 5)"],
                [[s["staff_id"], s["name"], s["labs"], s["courses"], s["exp_s"], s["got_s"], s["avg_s"], s["exp_f"],
                  s["got_f"], s["avg_f"]] for s in staff_summary(db)])
    ws = wb.create_sheet("Pending")
    write_sheet(ws, ["Rater Type", "Rater ID", "Rater Name", "Lab Room No.", "Course Code", "Course Name", "Cohort",
                     "Lab Staff ID", "Lab Staff Name"],
                [[r["role"], r["rid"], r["rname"], r["lab"], r["course_code"] or "All courses", r["course_name"],
                  r["cohort"] or "All cohorts", r["staff_id"], r["staff_name"]] for r in pending_rows(db)])
    audit(f"admin:{session['admin']}", "download", "responses")
    return _xlsx_response(wb, f"Lab_Feedback_Responses_{stamp}.xlsx")


@app.route("/admin/download/template")
@admin_required
def download_template():
    return _xlsx_response(build_upload_workbook([], []), "Lab_Feedback_Upload_Template.xlsx")


@app.route("/admin/password", methods=["POST"])
@admin_required
def admin_password():
    db = get_db()
    row = db.execute("SELECT * FROM admins WHERE username=?", (session["admin"],)).fetchone()
    cur, new, conf = (request.form.get(k, "") for k in ("current", "new", "confirm"))
    if not check_password_hash(row["password_hash"], cur):
        flash("Current password is incorrect.", "error")
    elif len(new) < 10 or not re.search(r"\d", new) or not re.search(r"[A-Za-z]", new):
        flash("New password must be at least 10 characters with letters and numbers.", "error")
    elif new != conf:
        flash("New password and confirmation do not match.", "error")
    else:
        set_admin_password(db, new)
        audit(f"admin:{session['admin']}", "password change")
        flash("Admin password changed successfully.", "success")
    return redirect(url_for("admin_panel"))


# ------------------------------------------------------------------- errors
ERRORS = {
    400: ("Request could not be processed", "Your session may have expired. Please go back, reload the page and try again."),
    403: ("Access denied", "You do not have permission to view this page."),
    404: ("Page not found", "The page you requested does not exist or is not assigned to you."),
    413: ("File too large", "The uploaded file exceeds the 10 MB limit."),
    500: ("Something went wrong", "An unexpected error occurred. Please try again."),
}


def _error(e):
    code = getattr(e, "code", 500) or 500
    title, msg = ERRORS.get(code, ERRORS[500])
    return render_template("error.html", code=code, title=title, msg=msg), code


for _code in ERRORS:
    app.register_error_handler(_code, _error)

init_db()

if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "reset-admin":
        new = os.environ.get("LAB_ADMIN_PASSWORD") or gen_password(12)
        c = connect()
        set_admin_password(c, new)
        c.close()
        print(f"Admin password reset. Username: admin  Password: {new}")
    elif len(sys.argv) >= 3 and sys.argv[1] == "import":
        c = connect()
        try:
            with open(sys.argv[2], "rb") as fh:
                print("Imported:", import_workbook(io.BytesIO(fh.read()), c))
        except ImportProblem as e:
            print("Import failed:\n  " + "\n  ".join(e.args[0]))
        c.close()
    else:
        host = os.environ.get("LAB_HOST", "127.0.0.1")
        port = int(os.environ.get("LAB_PORT", "5000"))
        print(f"Feedback form running at http://{host}:{port}   (admin: http://{host}:{port}/admin)")
        app.run(host=host, port=port, debug=False)
