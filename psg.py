import os
import sys
import subprocess
import glob
import argparse
from datetime import datetime as dt
import calendar
from datetime import timedelta as td
from decimal import Decimal as dec
import userpaths
import csv
import xlsxwriter


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

MONTHLYSTANDBYLIMIT = 168


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
    except Exception as exc:
        print(f"Could not open '{fn}'")
        print(f"Exception: {type(exc)}, Arguments: {exc.args}")
        return []


def read_dates(fn):
    try:
        dates = []
        with open(fn, "r") as f:
            dates = [dt.strptime(s.strip(), "%Y-%m-%d") for s in f]
        return dates
    except Exception as exc:
        print(f"Could not parse '{fn}'")
        print(f"Exception: {type(exc)}, Arguments: {exc.args}")
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


def get_hour_index(r) -> int:
    if r["Project"] in special_projects.keys():
        return special_projects[r["Project"]]
    elif r["Activity"] == "Standby Hours - Hungary":
        return ISTANDBY
    else:
        return IWORK


def is_working_day(date) -> bool:
    return (
        (date.weekday() < 5)
        and (date not in cfg_weekends)
        and (date not in cfg_holidays)
    ) or (date in cfg_workingdays)


# returns true if email to be processed
def filter_email(email):
    return len(cfg_users) == 0 or email in cfg_users or email.replace("@capgemini.com", "") in cfg_users


# returns true if project to be processed
def filter_project(project, projectdescription):
    if len(cfg_projects) == 0:
        return True

    proj = project.lower()
    projdesc = projectdescription.lower()

    projectmatch = False
    for p in cfg_projects:
        if p in proj or p in projdesc:
            projectmatch = True
            break

    return projectmatch


# ************
# Read configs
# ************

cfg_users = read_strings(os.path.join("cfg", "users.txt"), do_strip=True, do_lower=True)
cfg_projects = read_strings(os.path.join("cfg", "projects.txt"), do_lower=True)
cfg_holidays = read_dates(os.path.join("cfg", "holidays.txt"))
cfg_weekends = read_dates(os.path.join("cfg", "weekends.txt"))
cfg_workingdays = read_dates(os.path.join("cfg", "workingdays.txt"))


# ****************************************************************************

parser = argparse.ArgumentParser("psg - Presence Sheet Generator")
parser.add_argument(
    "-a",
    "--autoopen",
    action="store_true",
    help="automatically open generated file with Excel/Writer",
)
parser.add_argument(
    "-s", "--standbylimit", action="store_true", help="force monthly standby limit"
)
parser.add_argument(
    "filename",
    nargs="?",
    help="timesheet in CSV format to process; if omitted, latest 'TimesheetReport_*' file is picked from user's default Download folder",
)
args = parser.parse_args()

inputfilename = None
if args.filename is not None:
    if os.path.isfile(args.filename):
        inputfilename = args.filename
    else:
        print(f"Input file does not exist: {args.filename}")
        sys.exit(1)
else:
    downloads_folder = userpaths.get_downloads()
    matching_files = glob.glob(
        os.path.join(userpaths.get_downloads(), "TimesheetReport_*")
    )
    if len(matching_files) == 0:
        print(f"Could not find any timesheet in folder: {format(downloads_folder)}")
        sys.exit(1)
    inputfilename = max(matching_files, key=os.path.getctime)
print(f"Parsing: {inputfilename}")


# **************
# Summarize data
# **************

users = {}
approvers = {}
sumbyuser = {}
sumbyuserandproj = {}

min_date = dt.max
max_date = dt.min



try:
    with open(inputfilename, newline="", encoding="utf-8") as inputfile:
        reader = csv.DictReader(inputfile, delimiter="\t")
        for row in reader:
            # Assumption: if first field can be parsed as date, then it is a data row
            date = None
            try:
                date = dt.strptime(row["Date"], "%Y%m%d")
            except ValueError:
                continue

            email = row["Email Address"].lower()

            if not filter_email(email):
                continue

            if not filter_project(row["Project"], row["Project Description"]):
                continue

            if email not in users.keys():
                users[email] = row["User"]

            if email not in approvers.keys():
                approvers[email] = row["Level 1 Approver Name (configured)"]

            # sumbyuser update
            if email not in sumbyuser.keys():
                sumbyuser[email] = {}
            if date not in sumbyuser[email].keys():
                sumbyuser[email][date] = [dec0, dec0, dec0, dec0, dec0]
            sumbyuser[email][date][get_hour_index(row)] += dec(row["Hours"])

            # sumbyuserandproj update
            projectname = f"{row['Project']} {row['Project Description']}"
            if (email, projectname) not in sumbyuserandproj.keys():
                sumbyuserandproj[email, projectname] = {}
            if date not in sumbyuserandproj[email, projectname].keys():
                sumbyuserandproj[email, projectname][date] = [ dec0, dec0, dec0, dec0, dec0 ]
            sumbyuserandproj[email, projectname][date][get_hour_index(row)] += dec(row["Hours"])

            min_date = min(min_date, date)
            max_date = max(max_date, date)
except Exception as exc:
    print(f"Could not parse input: {inputfilename}")
    print(f"Exception: {type(exc)}, Arguments: {exc.args}")
    sys.exit(1)

# *************************
# Handle 168+ standby hours
# *************************

if args.standbylimit:
    for email in sumbyuser:
        # avoid importing dateutil.relativedelta for now...
        year = min_date.year
        month = min_date.month
        while year < max_date.year or (
            year == max_date.year and month <= max_date.month
        ):
            monthlystandy = sum(
                [
                    sumbyuser[email][d][ISTANDBY]
                    for d in sumbyuser[email]
                    if d.year == year and d.month == month
                ]
            )

            if monthlystandy > MONTHLYSTANDBYLIMIT:
                print(
                    f"Standby hours of {format_hours(monthlystandy)} in {year}-{month:02d} exceeds {MONTHLYSTANDBYLIMIT} hours for {email}"
                )
                for date in sumbyuser[email]:
                    if (
                        date.year == year
                        and date.month == month
                    ):
                        otmulti = 1 if is_working_day(date) else 2
                        if sumbyuser[email][date][ISTANDBY] >= 5 * otmulti:
                            w1 = sumbyuser[email][date][IWORK]
                            s1 = sumbyuser[email][date][ISTANDBY]
                            pluswork = sumbyuser[email][date][ISTANDBY] // (5 * otmulti)
                            minusstandby = pluswork * 5 * otmulti
                            sumbyuser[email][date][ISTANDBY] -= minusstandby
                            sumbyuser[email][date][IWORK] += pluswork
                            w2 = sumbyuser[email][date][IWORK]
                            s2 = sumbyuser[email][date][ISTANDBY]
                            print(
                                f"  {format_date(date)}, {email}: ({format_hours(w1)}, {format_hours(s1)}) + (+{format_hours(pluswork)}, -{format_hours(minusstandby)}) = ({format_hours(w2)}, {format_hours(s2)})"
                            )
                            monthlystandy -= minusstandby
                            if monthlystandy <= MONTHLYSTANDBYLIMIT:
                                print(
                                    f"  Standby hours in {year}-{month:02d} is {format_hours(monthlystandy)} hours for {email}"
                                )
                                break

            month += 1
            if month == 13:
                year += 1
                month = 1

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

worksheet.write(5, 0, "Name", fmtheadertxt)
worksheet.write(5, 1, "User", fmtheadertxt)
worksheet.write(5, 2, "Approver", fmtheadertxt)
worksheet.write(5, 3, "WorkH", fmtheadernum)
worksheet.write(5, 4, "WorkD", fmtheadernum)
worksheet.write(5, 5, "VacaD", fmtheadernum)
worksheet.write(5, 6, "SickD", fmtheadernum)
worksheet.write(5, 7, "OverH", fmtheadernum)
worksheet.write(5, 8, "StbyH", fmtheadernum)

date = min_date
lastmonth = 0
col = 9
while date <= max_date:
    if lastmonth != date.month:
        lastmonth = date.month
        worksheet.write(4, col, calendar.month_name[date.month], fmtheadertxt)
    worksheet.write_number(5, col, date.day, fmtheaderday)
    date = date + td(days=1)
    col += 1

worksheet.set_column(0, 0, width=40)
worksheet.set_column(1, 2, width=24)
worksheet.set_column(3, 8, width=8)
worksheet.set_column(9, col - 1, width=6)

missinglist = []
overtimelist = []
sicklist = []
vacationlist = []


row = 6
for email in sorted(sumbyuser.keys()):
    total_hours = [dec0, dec0, dec0, dec0, dec0]
    for date in sumbyuser[email]:
        for h in range(0, 5):
            total_hours[h] += sumbyuser[email][date][h]

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
        if date in sumbyuser[email].keys():
            hours = sumbyuser[email][date]
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
                worksheet.write(row, col, format_hours(hours[IWORK]), fmtover)
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
    worksheet.write(row, 1, users[email])
    worksheet.write(row, 2, approvers[email])
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
    if email not in sumbyuser.keys() and f"{email}@capgemini.com" not in sumbyuser.keys():
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

worksheet.autofilter(5, 0, row - 1, col - 1)

workbook.close()
print(f"Saved: {os.path.join(os.getcwd(), str(workbook.filename))}")

if args.autoopen:
    if sys.platform == "win32":
        os.system("start excel sum.xlsx")
        # subprocess.Popen(["start", "excel", "sum.xlsx"])
    elif sys.platform == "linux":
        subprocess.Popen(["libreoffice", "--calc", "sum.xlsx"])
