# Lab Staff Quality Feedback Form

Students and Faculty / Lab Head / PL / Cluster Head log in and rate the lab staff assigned to them.
The admin uploads the master Excel file and downloads all responses.

## Run it

```
pip install -r requirements.txt
python app.py
```
Or just double-click `run.bat`.

- Portal: http://127.0.0.1:5000
- Admin:  http://127.0.0.1:5000/admin/login (username `admin`)

On the **first run** the admin password is printed once in the console window. Note it down.
Forgot it? Run `python app.py reset-admin`.

## Test with the demo data

1. Log in as admin, then upload `Lab_Feedback_Demo_Data.xlsx` under **Upload Master Data**.
2. Log out, then log in with just an ID. No password is needed, and the portal works out from the ID whether it is a student or faculty:

| Detected as | ID | Sees |
|---|---|---|
| Student | 1001 (Vikhyat Kashyap) | PQR (PHYS1001, room PHY-101), LMN (CHEM1002, room CHE-204) |
| Student | 1002 | 4 cards, including PQR twice (two different courses in the same room) |
| Student | 1011 | 4 lab staff |
| Faculty | 5001 (ABC) | PQR, Deepak Negi |
| Faculty (Lab Head) | 7001 | PQR, Deepak Negi |
| Faculty (Cluster Head) | 7003 | all 5 lab staff |

## Excel format (`Lab_Feedback_Upload_Template.xlsx`)

- **Assignments** (required): Sr. No., Student ID, Student Name, Lab Room No., Course Code, Course Name, Cohort,
  Faculty ID, Faculty Name, Lab Staff ID, Lab Staff Name
- **Reviewers** (optional): Reviewer ID, Reviewer Name, Designation, Lab Room No., Lab Staff ID, Lab Staff Name

Feedback is collected once per lab staff + course + lab room + cohort, so a lab staff member who supports two
courses gets a separate card for each.

Only IDs in the uploaded file can log in.
Each upload **replaces** the assignment lists. Responses are never deleted.

## Downloads (Admin)

`Lab_Feedback_Responses_<date>.xlsx` has four sheets: Student Responses, Faculty Responses, Staff Summary (averages
and counts per lab staff) and Pending (who has not submitted yet).

## Security built in

- Students and faculty log in with their ID only. The admin page needs a password
- Logins are protected by a captcha, and a computer is blocked for 15 minutes after 20 unknown IDs
- Every form is protected against CSRF, the session cookie is HTTP-only, and sessions expire after 30 minutes
- The server checks every request, so a user can only open lab staff assigned to them
- One submission per rater + lab staff + lab room + course + cohort, enforced by a database UNIQUE rule (no duplicates even if
  the user double-clicks or presses Back)
- Security headers (CSP, no framing), and the Excel export escapes formula characters
- Every login, upload, download and submission is recorded in `audit_log` (in `instance/feedback.db`)

All data is stored in `instance/`. Back up that folder and never share it.
For hosting on a network or the internet, put the app behind HTTPS and set the environment variable `LAB_HTTPS=1`.
