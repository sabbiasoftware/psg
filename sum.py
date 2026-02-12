import os
import glob

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
    def __init__(self, date, name, comment, project, activity, hours):
        self.date = date
        self.name = name
        self.comment = comment
        self.project = project
        self.activity = activity
        self.hours = hours

    def __str__(self) -> str:
        return "{} {} {} {} {} {}".format(
            dt.strftime(self.date, "%Y%m%d"),
            self.name,
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
    with open(fn, "r", encoding="utf-8") as f:
        strings = f.read().splitlines()
    if do_strip:
        strings = [s.strip() for s in strings]
    if do_lower:
        strings = [s.lower() for s in strings]
    return strings


def read_dates(fn):
    dates = []
    with open(fn, "r") as f:
        dates = [dt.strptime(s.strip(), "%Y-%m-%d") for s in f]
    return dates


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


# ***************************
# Read configs and input file
# ***************************

cfg_users = read_strings(os.path.join("cfg", "users.txt"), do_strip=True, do_lower=True)
cfg_holidays = read_dates(os.path.join("cfg", "holidays.txt"))
cfg_weekends = read_dates(os.path.join("cfg", "weekends.txt"))
cfg_workingdays = read_dates(os.path.join("cfg", "workingdays.txt"))
cfg_projects = read_strings(os.path.join("cfg", "projects.txt"))


downloads_folder = userpaths.get_downloads()
raw_files = glob.glob(os.path.join(userpaths.get_downloads(), "TimesheetReport_*"))

if len(raw_files) == 0:
    print("Could not find any timesheet in {}".format(downloads_folder))
    input()
    quit()

latest_raw_filename = max(raw_files, key=os.path.getctime)
raw = read_strings(latest_raw_filename)


# ************
# Parse header
# ************

header = raw[0].split("\t")
idate = header.index("Date")
iname = header.index("Email Address")
icomment = header.index("Comment")
iproject = header.index("Project")
iprojectdescription = header.index("Project Description")
iactivity = header.index("Activity")
iactivitydescription = header.index("Activity Description")
ihours = header.index("Hours")


# **********
# Parse data
# **********

entries = []
actual_projects = set()

for r in raw:
    # TODO: find better way to figure if current line is header, footer or data
    if r[0:2] == "20":
        rr = r.split("\t")

        name = rr[iname][0:-14].strip().lower()

        project_included = (
            (rr[iproject] in special_projects)
            or (len(cfg_projects) == 0)
            or (max([str.find(rr[iproject], p) for p in cfg_projects]) >= 0)
            or (max([str.find(rr[iprojectdescription], p) for p in cfg_projects]) >= 0)
        )
        name_included = (len(cfg_users) == 0) or (name in cfg_users)

        if name_included and project_included:
            entries.append(
                Entry(
                    dt.strptime(rr[idate], "%Y%m%d"),
                    name,
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


daysums = {}

min_date = dt.max
max_date = dt.min

for e in entries:
    if e.name not in daysums.keys():
        daysums[e.name] = {}
    if e.date not in daysums[e.name].keys():
        daysums[e.name][e.date] = [dec0, dec0, dec0, dec0, dec0]
    daysums[e.name][e.date][get_hour_index(e)] += e.hours

    if e.date < min_date:
        min_date = e.date

    if e.date > max_date:
        max_date = e.date


# *******************
# Generate the output
# *******************

workbook = xlsxwriter.Workbook("sum.xlsx")
worksheet = workbook._add_sheet("Summary")

fmtheaderday = workbook.add_format({"align": "center"})
fmtheadername = workbook.add_format({"align": "left"})
fmtheadernum = workbook.add_format({"align": "right"})
fmtname = workbook.add_format({"align": "left"})
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

worksheet.write(0, 0, "Summary")
worksheet.write(1, 0, f"Duration: {format_date(min_date)}-{format_date(max_date)}")
worksheet.write(2, 0, f"Projects: {actual_projects}")
worksheet.write(3, 0, f"Generated: {format_datetime(dt.now())}")

worksheet.write(5, 0, "Name", fmtheadername)
worksheet.write(5, 1, "WorkH", fmtheadernum)
worksheet.write(5, 2, "WorkD", fmtheadernum)
worksheet.write(5, 3, "VacaD", fmtheadernum)
worksheet.write(5, 4, "SickD", fmtheadernum)
worksheet.write(5, 5, "OverH", fmtheadernum)
worksheet.write(5, 6, "StbyH", fmtheadernum)

date = min_date
col = 7
while date <= max_date:
    worksheet.write(5, col, date.day, fmtheaderday)
    date = date + td(days=1)
    col += 1

worksheet.set_column(0, 0, width=16)
worksheet.set_column(1, 6, width=6)
worksheet.set_column(7, col - 1, width=4)

missinglist = []
overtimelist = []
sicklist = []
vacationlist = []


row = 8
for name in sorted(daysums.keys()):
    total_hours = [dec0, dec0, dec0, dec0, dec0]
    for date in daysums[name]:
        for h in range(0, 5):
            total_hours[h] += daysums[name][date][h]

    # if there are filter projects configured, then filter out people with 0 hours against projects
    if len(cfg_projects) > 0 and total_hours[IWORK] == 0:
        continue

    missing = False
    overtime_hours = 0
    standby_hours = 0

    date = min_date
    col = 7
    while date <= max_date:
        hours = []
        if date in daysums[name].keys():
            hours = daysums[name][date]
        else:
            hours = [dec0, dec0, dec0, dec0, dec0]

        if is_working_day(date):
            if hours[0:4] == [dec8, dec0, dec0, dec0]:
                worksheet.write(row, col, "8", fmtwork)
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

    worksheet.write(row, 0, name, fmtname)
    worksheet.write(row, 1, total_hours[IWORK], fmtnum)
    worksheet.write(row, 2, (total_hours[IWORK] - overtime_hours) // dec8, fmtnum)
    worksheet.write(row, 3, total_hours[IVACATION] // dec8, fmtnum)
    worksheet.write(row, 4, total_hours[ISICK] // dec8, fmtnum)
    worksheet.write(row, 5, overtime_hours, fmtnum)
    worksheet.write(row, 6, total_hours[ISTANDBY], fmtnum)

    if missing:
        missinglist.append(name)

    if overtime_hours > dec0:
        overtimelist.append((name, overtime_hours))

    if total_hours[ISICK] > dec0:
        sicklist.append((name, total_hours[ISICK] // dec8))

    if total_hours[IVACATION] > dec0:
        vacationlist.append((name, total_hours[IVACATION] // dec8))

    row += 1


for name in sorted(cfg_users):
    if name not in daysums.keys():
        worksheet.write(row, 0, name, fmtname)

        date = min_date
        col = 7
        while date <= max_date:
            if is_working_day(date):
                worksheet.write(row, col, "-", fmtno)
            else:
                worksheet.write(row, col, "", fmtempty)

            date = date + td(days=1)
            col += 1

        for c in range(1, 7):
            worksheet.write(row, c, 0, fmtnum)

        missinglist.append(name)
        row += 1


workbook.close()

os.system(f"libreoffice --calc {'sum.xlsx'}")
# webbrowser.open_new_tab("sum_out.html")
