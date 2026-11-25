"""Excel helpers: upload template, demo data workbook and shared styling.

Run directly to (re)create the Excel files in this folder:
    python workbook_utils.py
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

HEAD_FILL = PatternFill("solid", fgColor="4A1145")
HEAD_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
THIN = Side(style="thin", color="D9CFE0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

ASSIGNMENT_HEADERS = ["Sr. No.", "Student ID", "Student Name", "Lab Room No.", "Course Code", "Course Name",
                      "Cohort", "Faculty ID", "Faculty Name", "Lab Staff ID", "Lab Staff Name"]
REVIEWER_HEADERS = ["Reviewer ID", "Reviewer Name", "Designation", "Lab Room No.",
                    "Lab Staff ID", "Lab Staff Name"]

INSTRUCTIONS = [
    ("HOW TO FILL THIS WORKBOOK", ""),
    ("Sheet: Assignments (required)",
     "One row per Student + Course + Lab Room + Cohort. Columns: Sr. No., Student ID, Student Name, "
     "Lab Room No., Course Code, Course Name, Cohort, Faculty ID, Faculty Name, Lab Staff ID, Lab Staff Name. "
     "Students see only the lab staff in their own rows; faculty see the lab staff in rows carrying their "
     "Faculty ID. If the same lab staff supports two courses, each course is rated separately."),
    ("Sheet: Reviewers (optional)",
     "Lab Head / PL / Cluster Head who also rate lab staff. One row per reviewer + lab staff + Lab Room No. "
     "They log in with role 'Faculty / Lab Head' and get the faculty questionnaire."),
    ("Re-uploading",
     "Each upload REPLACES the Assignments and Reviewers lists. Submitted responses are never deleted."),
    ("Login",
     "Students log in with their Student ID; faculty and reviewers with their Faculty / Reviewer ID. "
     "Only IDs present in this workbook can log in."),
]


def safe_cell(v):
    """Neutralise spreadsheet formula injection for text that came from users."""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + v
    return v


def write_sheet(ws, headers, rows, max_width=55):
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for row in rows:
        ws.append([safe_cell(v) for v in row])
    for idx, h in enumerate(headers, start=1):
        longest = max([len(str(h))] + [len(str(r[idx - 1])) for r in rows[:2000] if idx - 1 < len(r) and r[idx - 1] is not None])
        ws.column_dimensions[get_column_letter(idx)].width = min(max(10, longest + 3), max_width)
    ws.freeze_panes = "A2"
    if rows:
        # Excel table style gives banded rows + filter buttons without styling every cell (fast for big sheets)
        name = "T_" + "".join(ch for ch in ws.title if ch.isalnum())
        tab = Table(displayName=name, ref=f"A1:{get_column_letter(len(headers))}{len(rows) + 1}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
        ws.add_table(tab)


def build_upload_workbook(assignments, reviewers):
    wb = Workbook()
    ws = wb.active
    ws.title = "Assignments"
    write_sheet(ws, ASSIGNMENT_HEADERS, [[i + 1, *r] for i, r in enumerate(assignments)])

    ws = wb.create_sheet("Reviewers")
    write_sheet(ws, REVIEWER_HEADERS, reviewers)
    dv = DataValidation(type="list", formula1='"Lab Head,PL,Cluster Head"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C2:C2000")

    ws = wb.create_sheet("Instructions")
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 110
    for r, (a, b) in enumerate(INSTRUCTIONS, start=1):
        ws.cell(r, 1, a).font = Font(bold=True, color="4A1145", size=12 if r == 1 else 11)
        cell = ws.cell(r, 2, b)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 22 if r == 1 else 48
    return wb


# ---------------------------------------------------------------- demo data
# code: (lab room no., course code, course name, cohort, faculty id, faculty name, staff id, staff name)
_LABS = {
    "P1": ("PHY-101", "PHYS1001", "Engineering Physics Lab", "Physics_B1", 5001, "ABC", 6001, "PQR"),
    "P3": ("PHY-101", "PHYS2002", "Applied Optics Lab", "Physics_B1", 5001, "ABC", 6001, "PQR"),
    "P2": ("PHY-102", "PHYS1001", "Engineering Physics Lab", "Physics_B2", 5001, "ABC", 6005, "Deepak Negi"),
    "C1": ("CHE-204", "CHEM1002", "Engineering Chemistry Lab", "Chemistry_B1", 5002, "XYZ", 6002, "LMN"),
    "C2": ("CHE-204", "CHEM1002", "Engineering Chemistry Lab", "Chemistry_B2", 5002, "XYZ", 6002, "LMN"),
    "E1": ("ECE-305", "ECEG2001", "Basic Electronics Lab", "Electronics_B1", 5003, "Dr. Meera Iyer", 6003, "Suresh Kumar"),
    "K1": ("CSE-410", "CSEG1001", "Programming in C Lab", "Computer_B1", 5004, "Prof. Rahul Sen", 6004, "Anita Rawat"),
}
_STUDENTS = [
    (1001, "Vikhyat Kashyap", ["P1", "C1"]),
    (1002, "Aarav Sharma", ["P1", "P3", "C1", "E1"]),
    (1003, "Diya Negi", ["P1", "C1", "K1"]),
    (1004, "Rohan Verma", ["P1", "E1"]),
    (1005, "Ananya Bisht", ["P2", "C2", "K1"]),
    (1006, "Kabir Singh", ["P2", "C2"]),
    (1007, "Ishita Rawat", ["P2", "E1", "K1"]),
    (1008, "Arjun Mehra", ["P2", "C2", "E1"]),
    (1009, "Sneha Joshi", ["P1", "K1"]),
    (1010, "Aditya Thapa", ["P2", "C1"]),
    (1011, "Priya Gupta", ["P1", "C2", "E1", "K1"]),
    (1012, "Manav Chauhan", ["P2", "C1"]),
]
_REVIEWERS = [
    (7001, "Dr. Kavita Joshi", "Lab Head", "PHY-101", 6001, "PQR"),
    (7001, "Dr. Kavita Joshi", "Lab Head", "PHY-102", 6005, "Deepak Negi"),
    (7002, "Prof. Arun Mehta", "PL", "CHE-204", 6002, "LMN"),
    (7003, "Dr. Sanjay Bhatt", "Cluster Head", "PHY-101", 6001, "PQR"),
    (7003, "Dr. Sanjay Bhatt", "Cluster Head", "PHY-102", 6005, "Deepak Negi"),
    (7003, "Dr. Sanjay Bhatt", "Cluster Head", "CHE-204", 6002, "LMN"),
    (7003, "Dr. Sanjay Bhatt", "Cluster Head", "ECE-305", 6003, "Suresh Kumar"),
    (7003, "Dr. Sanjay Bhatt", "Cluster Head", "CSE-410", 6004, "Anita Rawat"),
]


def demo_data():
    assignments = []
    for sid, name, codes in _STUDENTS:
        for code in codes:
            assignments.append([sid, name, *_LABS[code]])
    return assignments, [list(r) for r in _REVIEWERS]


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    a, r = demo_data()
    build_upload_workbook(a, r).save(os.path.join(here, "Lab_Feedback_Demo_Data.xlsx"))
    build_upload_workbook([], []).save(os.path.join(here, "Lab_Feedback_Upload_Template.xlsx"))
    print("Created Lab_Feedback_Demo_Data.xlsx and Lab_Feedback_Upload_Template.xlsx")
