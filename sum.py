import os
import sys
import subprocess
import glob
import argparse

# import subprocess
from datetime import datetime as dt
from datetime import timedelta as td
from decimal import Decimal as dec

# from io import StringIO
import userpaths

# from odfdo import Document, Table, Style, Column
import xlsxwriter
# import webbrowser


class Entry:
    def __init__(self, date, email, user, approver, comment, project, activity, hours):
        self.date = date
        self.email = email
        self.user = user
        self.approver = approver
        self.comment = comment
        self.project = project
        self.activity = activity
        self.hours = hours

    def __str__(self) -> str:
        return "{};{};{};{};{};{};{};{}".format(
            dt.strftime(self.date, "%Y%m%d"),
            self.email,
            self.user,
            self.approver,
            self.comment,
            self.project,
            self.activity,
            self.hours,
        )


dec0 = dec("0")
dec8 = dec("8")

IWORK = 0
IVACATION = 1
ISICK = 2
IHOLIDAY = 3
ISTANDBY = 4

special_projects = {
    "Approved Absence (H)": IVACATION,
    "Vacations": IVACATION,
    "Sick Leave (H)": ISICK,
    "Medical Leave": ISICK,
    "Public Holiday": IHOLIDAY,
}


def read_strings(fn, do_strip=False, do_lower=False):
    strings = []
    try:
        with open(fn, "r", encoding="utf-8") as f:
            strings = f.read().splitlines()
        if do_strip:
            strings = [s.strip() for s in strings]
        if do_lower:
            strings = [s.lower() for s in strings]
        return strings
    except Exception:
        print(f"Could not open '{fn}'")
        return []


def read_dates(fn):
    try:
        dates = []
        with open(fn, "r") as f:
            dates = [dt.strptime(s.strip(), "%Y-%m-%d") for s in f]
        return dates
    except Exception:
        return []


def format_hours(hours):
    if isinstance(hours, int):
        return str(hours)
    elif isinstance(hours, dec):
        if hours.as_integer_ratio()[1] == 1:
            return "{:.0f}".format(hours)
        else:
            return "{:.2f}".format(hours)
    else:
        print(type(hours))
        return "???"


def format_date(t):
    return dt.strftime(t, "%Y-%m-%d")


def format_datetime(t):
    return dt.strftime(t, "%Y-%m-%d   %H:%M:%S")


def get_hour_index(e) -> int:
    if e.project in special_projects.keys():
        return special_projects[e.project]
    elif e.activity == "Standby Hours - Hungary":
        return ISTANDBY
    else:
        return IWORK


def is_working_day(date) -> bool:
    return (
        (date.weekday() < 5)
        and (date not in cfg_weekends)
        and (date not in cfg_holidays)
    ) or (date in cfg_workingdays)


def read_input(inputfilename):
    if inputfilename is None:
        downloads_folder = userpaths.get_downloads()
        matching_files = glob.glob(
            os.path.join(userpaths.get_downloads(), "TimesheetReport_*")
        )
        if len(matching_files) == 0:
            print("Could not find any timesheet in {}".format(downloads_folder))
            return None
        inputfilename = max(matching_files, key=os.path.getctime)
    return read_strings(inputfilename)


# ************
# Read configs
# ************

cfg_users = read_strings(os.path.join("cfg", "users.txt"), do_strip=True, do_lower=True)
cfg_holidays = read_dates(os.path.join("cfg", "holidays.txt"))
cfg_weekends = read_dates(os.path.join("cfg", "weekends.txt"))
cfg_workingdays = read_dates(os.path.join("cfg", "workingdays.txt"))
cfg_projects = read_strings(os.path.join("cfg", "projects.txt"))


# ****************************************************************************

parser = argparse.ArgumentParser("psg - Presence Sheet Generator")
parser.add_argument(
    "filename",
    nargs="?",
    help="Timesheet in CSV format to process. If omitted, latest 'TimesheetReport_*' file is picked from user's default Download folder",
)
args = parser.parse_args()
rawinput = read_input(args.filename)
if rawinput is None or len(rawinput) == 0:
    print("No input to parse, quitting")
    quit(1)


# ************
# Parse header
# ************

header = rawinput[0].split("\t")
idate = header.index("Date")
iemail = header.index("Email Address")
iuser = header.index("User")
iapprover = header.index("Level 1 Approver Name (configured)")
icomment = header.index("Comment")
iproject = header.index("Project")
iprojectdescription = header.index("Project Description")
iactivity = header.index("Activity")
iactivitydescription = header.index("Activity Description")
ihours = header.index("Hours")


# **********
# Parse data
# **********

parsedinput = []
actual_projects = set()

for r in rawinput:
    # TODO: find better way to figure if current line is header, subtotal, footer or data
    if r[0:2] == "20":
        rr = r.split("\t")

        email = rr[iemail][0:-14].strip().lower()

        project_included = (
            (rr[iproject] in special_projects)
            or (len(cfg_projects) == 0)
            or (max([str.find(rr[iproject], p) for p in cfg_projects]) >= 0)
            or (max([str.find(rr[iprojectdescription], p) for p in cfg_projects]) >= 0)
        )
        name_included = (len(cfg_users) == 0) or (email in cfg_users)

        if name_included and project_included:
            parsedinput.append(
                Entry(
                    dt.strptime(rr[idate], "%Y%m%d"),
                    email,
                    rr[iuser],
                    rr[iapprover],
                    rr[icomment],
                    rr[iproject],
                    rr[iactivity],
                    dec(rr[ihours]),
                )
            )
            if rr[iproject] not in special_projects:
                actual_projects.add(
                    "{} {}".format(rr[iproject], rr[iprojectdescription])
                )


# **************
# Summarize data
# **************

suminput = {}

min_date = dt.max
max_date = dt.min

for e in parsedinput:
    if e.email not in suminput.keys():
        suminput[e.email] = {}
        suminput[e.email]["user"] = e.user
        suminput[e.email]["approver"] = e.approver
    if e.date not in suminput[e.email].keys():
        suminput[e.email][e.date] = [dec0, dec0, dec0, dec0, dec0]
    suminput[e.email][e.date][get_hour_index(e)] += e.hours

    if e.date < min_date:
        min_date = e.date

    if e.date > max_date:
        max_date = e.date


# ***********************
# Generate the Excel file
# ***********************

workbook = xlsxwriter.Workbook("sum.xlsx")
worksheet = workbook._add_sheet("Summary")

fmtheaderday = workbook.add_format({"align": "center", "bold": "true"})
fmtheadertxt = workbook.add_format({"align": "left", "bold": "true"})
fmtheadernum = workbook.add_format({"align": "right", "bold": "true"})
fmttxt = workbook.add_format({"align": "left"})
fmtnum = workbook.add_format({"align": "right"})
fmtwork = workbook.add_format({"align": "center", "bg_color": "#90ee90"})
fmtvaca = workbook.add_format({"align": "center", "bg_color": "#ffff00"})
fmtsick = workbook.add_format({"align": "center", "bg_color": "#da70d6"})
fmtno = workbook.add_format({"align": "center", "bg_color": "#d3d3d3"})
fmtunder = workbook.add_format({"align": "center", "bg_color": "#9acd32"})
fmtover = workbook.add_format({"align": "center", "bg_color": "#ffa500"})
fmtquest = workbook.add_format({"align": "center", "bg_color": "#808080"})
fmtempty = workbook.add_format({"align": "center", "bg_color": "#ffffff"})


actual_projects = "All"
if len(cfg_projects) > 0:
    actual_projects = ", ".join(actual_projects)

worksheet.write(0, 0, f"Duration: {format_date(min_date)}-{format_date(max_date)}")
worksheet.write(1, 0, f"Projects: {actual_projects}")
worksheet.write(2, 0, f"Generated: {format_datetime(dt.now())}")

worksheet.write(4, 0, "Name", fmtheadertxt)
worksheet.write(4, 1, "User", fmtheadertxt)
worksheet.write(4, 2, "Approver", fmtheadertxt)
worksheet.write(4, 3, "WorkH", fmtheadernum)
worksheet.write(4, 4, "WorkD", fmtheadernum)
worksheet.write(4, 5, "VacaD", fmtheadernum)
worksheet.write(4, 6, "SickD", fmtheadernum)
worksheet.write(4, 7, "OverH", fmtheadernum)
worksheet.write(4, 8, "StbyH", fmtheadernum)

date = min_date
col = 9
while date <= max_date:
    worksheet.write_number(4, col, date.day, fmtheaderday)
    date = date + td(days=1)
    col += 1

worksheet.set_column(0, 2, width=24)
worksheet.set_column(3, 8, width=8)
worksheet.set_column(9, col - 1, width=6)

missinglist = []
overtimelist = []
sicklist = []
vacationlist = []


row = 5
for email in sorted(suminput.keys()):
    total_hours = [dec0, dec0, dec0, dec0, dec0]
    for date in suminput[email]:
        if isinstance(date, dt):
            for h in range(0, 5):
                total_hours[h] += suminput[email][date][h]

    # if there are filter projects configured, then filter out people with 0 hours against projects
    if len(cfg_projects) > 0 and total_hours[IWORK] == 0:
        continue

    missing = False
    overtime_hours = 0
    standby_hours = 0

    date = min_date
    col = 9
    while date <= max_date:
        hours = []
        if date in suminput[email].keys():
            hours = suminput[email][date]
        else:
            hours = [dec0, dec0, dec0, dec0, dec0]

        if is_working_day(date):
            if hours[0:4] == [dec8, dec0, dec0, dec0]:
                worksheet.write_number(row, col, 8, fmtwork)
            elif hours[0:4] == [dec0, dec8, dec0, dec0]:
                worksheet.write(row, col, "V", fmtvaca)
            elif hours[0:4] == [dec0, dec0, dec8, dec0]:
                worksheet.write(row, col, "S", fmtsick)
            elif hours[0:4] == [dec0, dec0, dec0, dec0]:
                worksheet.write(row, col, "-", fmtno)
                missing = True
            elif (hours[IWORK] < dec8) and (hours[1:4] == [dec0, dec0, dec0]):
                worksheet.write(row, col, format_hours(hours[IWORK]), fmtunder)
                missing = True
            elif (hours[IWORK] > dec8) and (hours[1:4] == [dec0, dec0, dec0]):
                worksheet.write(row, col, format_hours(hours[IWORK] - dec8), fmtover)
                overtime_hours += hours[IWORK] - dec8
            else:
                worksheet.write(row, col, "?", fmtquest)
                missing = True
        else:
            if hours[IWORK] > dec0:
                worksheet.write(row, col, format_hours(hours[IWORK]), fmtover)
                overtime_hours += hours[IWORK]
            elif hours[0:4] == [dec0, dec0, dec0, dec0]:
                worksheet.write(row, col, "", fmtempty)
            elif hours[0:4] == [dec0, dec0, dec0, dec8]:
                worksheet.write(row, col, "", fmtempty)
            else:
                worksheet.write(row, col, "?", fmtquest)

        date = date + td(days=1)
        col += 1

    worksheet.write(row, 0, email, fmttxt)
    worksheet.write(row, 1, suminput[email]["user"])
    worksheet.write(row, 2, suminput[email]["approver"])
    worksheet.write_number(row, 3, int(total_hours[IWORK]), fmtnum)
    worksheet.write_number(
        row, 4, int((total_hours[IWORK] - overtime_hours) // dec8), fmtnum
    )
    worksheet.write_number(row, 5, int(total_hours[IVACATION] // dec8), fmtnum)
    worksheet.write_number(row, 6, int(total_hours[ISICK] // dec8), fmtnum)
    worksheet.write_number(row, 7, int(overtime_hours), fmtnum)
    worksheet.write_number(row, 8, int(total_hours[ISTANDBY]), fmtnum)

    if missing:
        missinglist.append(email)

    if overtime_hours > dec0:
        overtimelist.append((email, overtime_hours))

    if total_hours[ISICK] > dec0:
        sicklist.append((email, total_hours[ISICK] // dec8))

    if total_hours[IVACATION] > dec0:
        vacationlist.append((email, total_hours[IVACATION] // dec8))

    row += 1


for email in sorted(cfg_users):
    if email not in suminput.keys():
        worksheet.write(row, 0, email, fmttxt)

        date = min_date
        col = 9
        while date <= max_date:
            if is_working_day(date):
                worksheet.write(row, col, "-", fmtno)
            else:
                worksheet.write(row, col, "", fmtempty)

            date = date + td(days=1)
            col += 1

        for c in range(1, 7):
            worksheet.write(row, c, 0, fmtnum)

        missinglist.append(email)
        row += 1

worksheet.autofilter(4, 0, row - 1, col - 1)

workbook.close()

if sys.platform == "win32":
    os.system("start excel sum.xlsx")
    # subprocess.Popen(["start", "excel", "sum.xlsx"])
elif sys.platform == "linux":
    subprocess.Popen(["libreoffice", "--calc", "sum.xlsx"])
# webbrowser.open_new_tab("sum_out.html")
