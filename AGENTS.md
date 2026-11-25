# Project context for AI coding assistants (Antigravity, Claude Code, Copilot, etc.)

Read this file fully before making changes. It records what the project is, how it works,
the decisions already made, and what is still pending.

## What this is

**Lab Staff Quality Feedback Form** for UPES. Students and faculty log in with their ID and rate the
lab staff assigned to them. The admin uploads a master Excel file that decides who rates whom,
and downloads all responses as Excel.

- Owner: the APO office (non-developer). **Explain things in simple, non-technical language.**
- Users: students, plus faculty / Lab Head / PL / Cluster Head (all treated as the "faculty" role).
- Scale: thousands of students (the real master file has about 16,000 assignment rows).

## Tech stack

- **Python 3.12+ (the VM has 3.14) + Flask 3** (single file app: `app.py`), **openpyxl** for Excel, **SQLite** for storage.
- No JavaScript framework. Server-rendered Jinja templates plus a small `static/app.js`.
- Production server: **waitress** behind **IIS** on a **Windows VM** (see Hosting below).

## Files

| File | Purpose |
|---|---|
| `app.py` | The whole app: config, DB schema and migrations, login, feedback flow, admin, Excel import/export, CLI |
| `workbook_utils.py` | Excel helpers: `safe_cell` (formula-injection escape), `write_sheet`, `build_upload_workbook` (blank template), `demo_data` |
| `icons.py` | Inline SVG icons (`icon(name)`) used by templates |
| `templates/base.html` | Shared layout. All other templates extend it |
| `templates/login.html` | Landing page (`/` redirects here when logged out) |
| `templates/dashboard.html` -> `form.html` -> `preview.html` -> `summary.html` | User flow: pick lab staff, rate, preview + declaration, submitted |
| `templates/admin*.html` | Admin login, panel (upload/download), responses list |
| `static/style.css`, `static/app.js` | Styling and small UI behaviour |
| `web.config` | IIS config: runs the app via HttpPlatformHandler + waitress |
| `HOSTING_GUIDE_IIS.md` | Step-by-step hosting guide written for the university IT team |
| `check_server.ps1` | Read-only checker for the VM (IIS, HttpPlatformHandler, Python, venv, pool, permissions) |
| `run.bat` | Local run on a PC (double-click) |
| `Lab_Feedback_Demo_Data.xlsx` | Small fake dataset for testing (safe to commit) |
| `instance/` | **Live data**: `feedback.db` (SQLite) and `secret.key`. Created at runtime |

## How it works

- **Login**: ID only, no password. The role is detected from the uploaded lists
  (Student ID in `assignments`, otherwise Faculty ID in `assignments` or Reviewer ID in `reviewer_assignments`).
  A simple math captcha is required. Sessions last 30 minutes.
- **Lockout** (in-memory dict `_FAILS`): 5 failures per ID, or 20 per IP address, means a 15-minute lock.
  Because it is in memory, **the server must run as ONE process** (IIS Maximum Worker Processes = 1, waitress threads = 8).
- **Targets**: one feedback card per *lab staff + lab room + course code + cohort* (`target_key`, `get_targets`).
- **Questions**: `STUDENT_QUESTIONS` and `FACULTY_QUESTIONS` in `app.py` (up to `MAX_Q = 7`), rated with `RATING_SCALE`.
- **Duplicates**: blocked by the DB rule `UNIQUE (rater_role, rater_id, staff_id, lab, course_code, cohort)` on `responses`.
- **Admin**: `/admin/login` (username `admin`). Password set with `python app.py reset-admin`.
  Admin can upload the master Excel file (this **replaces** the assignments), download responses
  (4 sheets: Student Responses, Faculty Responses, Staff Summary, Pending), download a blank template,
  view responses, and delete one or all responses.
  *Note: README says "Responses are never deleted", but admin delete routes do exist.*
- **Audit**: logins, uploads, downloads and submissions go to the `audit_log` table, together with the IP.

### Database tables (SQLite, `instance/feedback.db`)

`assignments`, `reviewer_assignments`, `responses`, `admins`, `audit_log`. The schema is in `SCHEMA` / `RESPONSES_SQL`
in `app.py`. `migrate()` adds new columns to older databases, so keep it working when you change the schema.

### Excel upload format

- **Assignments** sheet (required): Sr. No., Student ID, Student Name, Lab Room No., Course Code, Course Name, Cohort,
  Faculty ID, Faculty Name, Lab Staff ID, Lab Staff Name
- **Reviewers** sheet (optional): Reviewer ID, Reviewer Name, Designation, Lab Room No., Lab Staff ID, Lab Staff Name
- Column names are matched loosely (`norm()`, `ASSIGN_MAP`, `REVIEWER_MAP`, `ROOM_ALIASES`).

## Commands

```
pip install -r requirements.txt
python app.py                              # local dev server at http://127.0.0.1:5000
python app.py reset-admin                  # print a new admin password
python app.py import "Lab_Feedback_Demo_Data.xlsx"
python -m waitress --port=8000 app:app     # production-style server
```

Environment variables: `LAB_HTTPS=1` (secure cookies; set when served over HTTPS), `LAB_HOST`, `LAB_PORT`,
`LAB_SECRET_KEY` (optional; otherwise `instance/secret.key` is used), `LAB_ADMIN_PASSWORD` (used by reset-admin).

## Hosting decision (already made, do not change without asking)

- **Target: the university's Windows VM, IIS + HttpPlatformHandler v1.2 + waitress**, configured by `web.config`.
- Site folder assumed: `E:\labfeedback`, with venv at `E:\labfeedback\venv`.
  If the folder is different, update `processPath` and `stdoutLogFile` in `web.config`.
  Real Excel data files are kept OUTSIDE the app folder (for example `E:\LabData\`) and uploaded through the admin page.
- **The VM already hosts another IIS site (the owner's dashboard, on D:; this app is on E:).** The feedback app must be a
  **separate IIS website** with its own app pool `LabFeedback` and its own host name/port. Do NOT add it as a
  sub-application under the dashboard: the app has no sub-path (`SCRIPT_NAME`) support. `iisreset` briefly restarts both sites.
  Folder permissions are locked down (inheritance removed) so the dashboard's app pool cannot read `instance/`.
- IIS has no "Python" option in its UI. `web.config` is what makes the site run Python. There is no `index.html`.
- waitress is started with `--trusted-proxy=*` and `--trusted-proxy-headers="x-forwarded-for x-forwarded-proto"` so that
  `request.remote_addr` is the real visitor IP (tested locally, including `ip:port` values).
- Rejected options: Vercel and Supabase (owner does not want them). PythonAnywhere free and Cloudflare Tunnel were
  discussed as fallbacks. A static-HTML rewrite (Google Sheets / Power Automate backend) is only plan B if IT refuses Python.

## Pending / open items

1. **Verify real IPs on IIS.** After deployment, check that `audit_log.ip` shows real addresses and not `127.0.0.1`
   (`HOSTING_GUIDE_IIS.md`, Step 9). If it is always `127.0.0.1`, HttpPlatformHandler is not forwarding
   `X-Forwarded-For`, and the per-IP lockout would lock out everyone. Fix that before going live.
2. **GitHub**: code is on https://github.com/vikhyatkashyap4999-cloud/Lab-Staff-Feedback-Form (branch `main`).
   The repo is **public**, so never push data or secrets. The owner was advised to consider making it private.
   The GitHub CLI cannot be installed on the owner's laptop (organisation policy). Git and Git Credential Manager work.
3. The README "never deleted" statement versus the admin delete routes (see above) should be made consistent.

## Rules

- **Never commit or share real data**: `instance/`, `data received/`, `*.db`, `*.key`, and any real Excel file.
  In particular, `Lab_Feedback_Upload_Template.xlsx` is **not** a blank template: it holds about 16,000 real rows.
  The blank template is generated by the app (`/admin/download/template`). `.gitignore` already covers these.
- Never delete or overwrite `instance/` when updating the app. It holds all the responses.
- Keep security features intact: CSRF on every POST, captcha, lockout, CSP/security headers, the
  formula-escape on Excel export (`safe_cell`), the server-side check that a user can only open their own targets,
  and the UNIQUE duplicate rule.
- Keep it a single process (in-memory lockout). Do not add gunicorn/multiple workers without moving `_FAILS` into the DB.
- The platform is Windows. Prefer PowerShell/batch instructions, and keep `run.bat` working.
- After changes, test locally with `Lab_Feedback_Demo_Data.xlsx` (see the README for test IDs such as 1001, 5001 and 7003).
