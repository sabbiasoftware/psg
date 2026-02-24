import os
import sys
import subprocess
import glob
import argparse
from enum import Enum, auto
from datetime import datetime as dt
import calendar
from datetime import timedelta as td
from decimal import Decimal as dec
import userpaths
import csv
import xlsxwriter
import traceback


dec0 = dec("0")
dec8 = dec("8")

class HourType(Enum):
    WORK = 1
    VACATION = 2
    SICK = 3
    HOLIDAY = 4
    STANDBY = 5

class HourFormat(Enum):
    WORK = auto()
    UNDER = auto()
    OVER = auto()
    VACATION = auto()
    SICK = auto()
    HOLIDAY = auto()
    MISS = auto()
    QUESTION = auto()
    EMPTY = auto()

special_projects = {
    "Approved Absence (H)": HourType.VACATION,
    "Vacations": HourType.VACATION,
    "Sick Leave (H)": HourType.SICK,
    "Medical Leave": HourType.SICK,
    "Public Holiday": HourType.HOLIDAY,
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


# obsolete...
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

def dec_to_number(d):
    if d.as_integer_ratio()[1] == 1:
        return int(d)
    else:
        return float(d)


def format_date(t):
    return dt.strftime(t, "%Y-%m-%d")


def format_datetime(t):
    return dt.strftime(t, "%Y-%m-%d   %H:%M:%S")


def get_hour_type(project, activity) -> HourType:
    if project in special_projects.keys():
        return special_projects[project]
    elif activity == "Standby Hours - Hungary":
        return HourType.STANDBY
    else:
        return HourType.WORK


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


##############
# Read configs
##############

cfg_users = read_strings(os.path.join("cfg", "users.txt"), do_strip=True, do_lower=True)
cfg_projects = read_strings(os.path.join("cfg", "projects.txt"), do_lower=True)
cfg_holidays = read_dates(os.path.join("cfg", "holidays.txt"))
cfg_weekends = read_dates(os.path.join("cfg", "weekends.txt"))
cfg_workingdays = read_dates(os.path.join("cfg", "workingdays.txt"))


##############################################################################

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


################
# Summarize data
################

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

            t = get_hour_type(row["Project"], row["Activity"])

            # sumbyuser update
            if email not in sumbyuser.keys():
                sumbyuser[email] = {}
            if date not in sumbyuser[email].keys():
                sumbyuser[email][date] = {}
            sumbyuser[email][date][t] = sumbyuser[email][date].get(t, 0) + dec(row["Hours"])

            # sumbyuserandproj update
            proj = row['Project']
            desc = row['Project Description']
            project = f"{proj} {desc}" if proj != desc else proj
            activity = row['Activity']
            if (email, project, activity) not in sumbyuserandproj.keys():
                sumbyuserandproj[email, project, activity] = {}
            if date not in sumbyuserandproj.keys():
                sumbyuserandproj[email, project, activity][date] = {}
            sumbyuserandproj[email, project, activity][date][t] = sumbyuserandproj[email, project, activity][date].get(t, 0) + dec(row["Hours"])

            min_date = min(min_date, date)
            max_date = max(max_date, date)
except Exception as exc:
    print(f"Could not parse input: {inputfilename}")
    print(f"Exception: {type(exc)}, Arguments: {exc.args}")
    traceback.print_exc()
    sys.exit(1)

###########################
# Handle 168+ standby hours
###########################

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
                    sumbyuser[email][d].get(HourType.STANDBY, 0)
                    for d in sumbyuser[email]
                    if d.year == year and d.month == month
                ]
            )

            if monthlystandy > MONTHLYSTANDBYLIMIT:
                print(f"Standby hours of {format_hours(monthlystandy)} in {year}-{month:02d} exceeds {MONTHLYSTANDBYLIMIT} hours for {email}")
                for date in sumbyuser[email]:
                    if (
                        date.year == year
                        and date.month == month
                    ):
                        otmulti = 1 if is_working_day(date) else 2
                        if sumbyuser[email][date].get(HourType.STANDBY, 0) >= 5 * otmulti:
                            w1 = sumbyuser[email][date].get(HourType.WORK, 0)
                            s1 = sumbyuser[email][date].get(HourType.STANDBY, 0)
                            pluswork = sumbyuser[email][date][HourType.STANDBY] // (5 * otmulti)
                            minusstandby = pluswork * 5 * otmulti
                            w2 = w1 + pluswork
                            s2 = s1 - minusstandby
                            sumbyuser[email][date][HourType.WORK] = w2
                            sumbyuser[email][date][HourType.STANDBY] = s2
                            print(f"  {format_date(date)}, {email}: ({format_hours(w1)}, {format_hours(s1)}) + (+{format_hours(pluswork)}, -{format_hours(minusstandby)}) = ({format_hours(w2)}, {format_hours(s2)})")
                            monthlystandy -= minusstandby
                            if monthlystandy <= MONTHLYSTANDBYLIMIT:
                                print(f"  Standby hours in {year}-{month:02d} is {format_hours(monthlystandy)} hours for {email}")
                                break

            month += 1
            if month == 13:
                year += 1
                month = 1


###################
# Generate workbook
###################

def generate_title(worksheet, duration, projects, generated):
    worksheet.write(0, 0, f"Duration: {duration}")
    worksheet.write(1, 0, f"Projects: {projects}")
    worksheet.write(2, 0, f"Generated: {generated}")

def generate_header_byuser(worksheet, fmtheaderday, fmtheadertxt, fmtheadernum):
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


def generate_header_byuserandproject(worksheet, fmtheaderday, fmtheadertxt, fmtheadernum):
    worksheet.write(5, 0, "Name", fmtheadertxt)
    worksheet.write(5, 1, "User", fmtheadertxt)
    worksheet.write(5, 2, "Approver", fmtheadertxt)
    worksheet.write(5, 3, "Project", fmtheadertxt)
    worksheet.write(5, 4, "Activity", fmtheadertxt)
    worksheet.write(5, 5, "WorkH", fmtheadernum)
    worksheet.write(5, 6, "VacaD", fmtheadernum)
    worksheet.write(5, 7, "SickD", fmtheadernum)

    date = min_date
    lastmonth = 0
    col = 8
    while date <= max_date:
        if lastmonth != date.month:
            lastmonth = date.month
            worksheet.write(4, col, calendar.month_name[date.month], fmtheadertxt)
        worksheet.write_number(5, col, date.day, fmtheaderday)
        date = date + td(days=1)
        col += 1

    worksheet.set_column(0, 0, width=40)
    worksheet.set_column(1, 2, width=24)
    worksheet.set_column(3, 3, width=48)
    worksheet.set_column(4, 4, width= 12)
    worksheet.set_column(5, 7, width=8)
    worksheet.set_column(8, col - 1, width=6)

workbook = xlsxwriter.Workbook("sum.xlsx")

fmtheaderday = workbook.add_format({"align": "center", "bold": "true"})
fmtheadertxt = workbook.add_format({"align": "left", "bold": "true"})
fmtheadernum = workbook.add_format({"align": "right", "bold": "true"})
fmttxt = workbook.add_format({"align": "left"})
fmtnum = workbook.add_format({"align": "right"})
hourFormats = {
    HourFormat.WORK: workbook.add_format({"align": "center", "bg_color": "#90ee90"}),
    HourFormat.UNDER: workbook.add_format({"align": "center", "bg_color": "#9acd32"}),
    HourFormat.OVER: workbook.add_format({"align": "center", "bg_color": "#ffa500"}),
    HourFormat.VACATION: workbook.add_format({"align": "center", "bg_color": "#ffff00"}),
    HourFormat.SICK: workbook.add_format({"align": "center", "bg_color": "#da70d6"}),
    HourFormat.MISS: workbook.add_format({"align": "center", "bg_color": "#d3d3d3"}),
    HourFormat.QUESTION: workbook.add_format({"align": "center", "bg_color": "#808080"}),
    HourFormat.EMPTY: workbook.add_format({"align": "center", "bg_color": "#ffffff"})
}

wsbyuser = workbook.add_worksheet("By user")
wsbyuserandproj = workbook.add_worksheet("By user and project")

duration = f"{format_date(min_date)}-{format_date(max_date)}"
actual_projects = "All" if len(cfg_projects) == 0 else ", ".join(cfg_projects)
generated = f"{format_datetime(dt.now())}"

generate_title(wsbyuser, duration, actual_projects, generated)
generate_header_byuser(wsbyuser, fmtheaderday, fmtheadertxt, fmtheadernum)
generate_title(wsbyuserandproj, duration, actual_projects, generated)
generate_header_byuserandproject(wsbyuserandproj, fmtheaderday, fmtheadertxt, fmtheadernum)


def no_other_hours(hour_type, hours):
    return sum( [ hours[ht] for ht in hours if ht != hour_type ] ) == 0

# return hours of hours_type if every other hours are 0 except standby
def get_only_hours(hour_type, hours):
    if hour_type not in hours:
        return None
    if sum( [ hours[ht] for ht in hours if ht != hour_type and ht != HourType.STANDBY ] ) > 0:
        return None
    return hours[hour_type]

def get_active_hours(hours):
    return sum( [ hours[ht] for ht in hours if ht != HourType.STANDBY ] )

# return (value, format)
def get_day_cell(date, hours):
    w = get_only_hours(HourType.WORK, hours)
    if is_working_day(date):
        if w is not None:
            if w == 8:
                return w, HourFormat.WORK
            elif w < 8:
                return w, HourFormat.UNDER
            else:
                return w, HourFormat.OVER
        elif get_only_hours(HourType.VACATION, hours) == 8:
            return "V", HourFormat.VACATION
        elif get_only_hours(HourType.SICK, hours) == 8:
            return "S", HourFormat.SICK
        elif get_active_hours(hours) == 0:
            return "-", HourFormat.MISS
        else:
            return "?", HourFormat.QUESTION
    else:
        if w is not None and w > 0:
            return w, HourFormat.OVER
        elif get_only_hours(HourType.HOLIDAY, hours) == 8:
            return "", HourFormat.EMPTY
        elif get_active_hours(hours) == 0:
            return "", HourFormat.EMPTY
        else:
            return "?", HourFormat.QUESTION


#######################
# generate data by user
#######################

row = 6
col = 9
for email in sorted(sumbyuser.keys()):
    total_hours = {}
    for ht in HourType:
        total_hours[ht] = sum( [ sumbyuser[email][date].get(ht, 0) for date in sumbyuser[email] ] )

    # if there are filter projects configured, then filter out people with 0 hours against projects
    if len(cfg_projects) > 0 and total_hours[HourType.WORK] == 0:
        continue

    weekday_overtime_hours = sum( [ sumbyuser[email][date].get(HourType.WORK, 0) - 8 for date in sumbyuser[email] if is_working_day(date) and sumbyuser[email][date].get(HourType.WORK, 0) > 8 ] )
    weekend_overtime_hours = sum( [ sumbyuser[email][date].get(HourType.WORK, 0) for date in sumbyuser[email] if not is_working_day(date) ] )
    overtime_hours = weekday_overtime_hours + weekend_overtime_hours

    date = min_date
    col = 9
    while date <= max_date:
        hours = {}
        if date in sumbyuser[email].keys():
            hours = sumbyuser[email][date]

        value, format = get_day_cell(date, hours)
        if isinstance(value, int) or isinstance(value, float):
            wsbyuser.write_number(row, col, value, hourFormats[format])
        else:
            wsbyuser.write(row, col, value, hourFormats[format])

        date = date + td(days=1)
        col += 1

    wsbyuser.write(row, 0, email, fmttxt)
    wsbyuser.write(row, 1, users[email])
    wsbyuser.write(row, 2, approvers[email])
    wsbyuser.write_number(row, 3, int(total_hours[HourType.WORK]), fmtnum)
    wsbyuser.write_number(row, 4, int((total_hours[HourType.WORK] - overtime_hours) // dec8), fmtnum)
    wsbyuser.write_number(row, 5, int(total_hours[HourType.VACATION] // dec8), fmtnum)
    wsbyuser.write_number(row, 6, int(total_hours[HourType.SICK] // dec8), fmtnum)
    wsbyuser.write_number(row, 7, int(overtime_hours), fmtnum)
    wsbyuser.write_number(row, 8, int(total_hours[HourType.STANDBY]), fmtnum)

    row += 1


for email in sorted(cfg_users):
    if email not in sumbyuser.keys() and f"{email}@capgemini.com" not in sumbyuser.keys():
        wsbyuser.write(row, 0, email, fmttxt)

        date = min_date
        col = 9
        while date <= max_date:
            if is_working_day(date):
                wsbyuser.write(row, col, "-", hourFormats[HourFormat.MISS])
            else:
                wsbyuser.write(row, col, "", hourFormats[HourFormat.EMPTY])

            date = date + td(days=1)
            col += 1

        for c in range(1, 7):
            wsbyuser.write(row, c, 0, fmtnum)

        row += 1

wsbyuser.autofilter(5, 0, row - 1, col - 1)


###################################
# generate data by user and project
###################################

row = 6
col = 9
for email, project, activity in sorted(sumbyuserandproj.keys()):
    if get_hour_type(project, activity) != HourType.STANDBY:
        total_hours = {}
        for ht in HourType:
            total_hours[ht] = sum( [ sumbyuserandproj[email, project, activity][date].get(ht, 0) for date in sumbyuserandproj[email, project, activity] ] )

        # if there are filter projects configured, then filter out people with 0 hours against projects
        if len(cfg_projects) > 0 and total_hours[HourType.WORK] == 0:
            continue

        # weekday_overtime_hours = sum( [ sumbyuser[email][date].get(HourType.WORK, 0) - 8 for date in sumbyuser[email] if is_working_day(date) and sumbyuser[email][date].get(HourType.WORK, 0) > 8 ] )
        # weekend_overtime_hours = sum( [ sumbyuser[email][date].get(HourType.WORK, 0) for date in sumbyuser[email] if not is_working_day(date) ] )
        # overtime_hours = weekday_overtime_hours + weekend_overtime_hours

        date = min_date
        col = 8
        while date <= max_date:
            hours = {}
            if date in sumbyuserandproj[email, project, activity].keys():
                hours = sumbyuserandproj[email, project, activity][date]

            value, format = get_day_cell(date, hours)

            if isinstance(value, int) or isinstance(value, float):
                wsbyuserandproj.write_number(row, col, value, hourFormats[format])
            else:
                wsbyuserandproj.write(row, col, value, hourFormats[format])

            date = date + td(days=1)
            col += 1

        wsbyuserandproj.write(row, 0, email, fmttxt)
        wsbyuserandproj.write(row, 1, users[email])
        wsbyuserandproj.write(row, 2, approvers[email])
        wsbyuserandproj.write(row, 3, project)
        wsbyuserandproj.write(row, 4, activity)
        wsbyuserandproj.write_number(row, 5, int(total_hours[HourType.WORK]), fmtnum)
        wsbyuserandproj.write_number(row, 6, int(total_hours[HourType.VACATION] // dec8), fmtnum)
        wsbyuserandproj.write_number(row, 7, int(total_hours[HourType.SICK] // dec8), fmtnum)

        row += 1


for email in sorted(cfg_users):
    if email not in sumbyuser.keys() and f"{email}@capgemini.com" not in sumbyuser.keys():
        wsbyuserandproj.write(row, 0, email, fmttxt)

        date = min_date
        col = 8
        while date <= max_date:
            if is_working_day(date):
                wsbyuserandproj.write(row, col, "-", hourFormats[HourFormat.MISS])
            else:
                wsbyuserandproj.write(row, col, "", hourFormats[HourFormat.EMPTY])

            date = date + td(days=1)
            col += 1

        for c in range(1, 7):
            wsbyuserandproj.write(row, c, 0, fmtnum)

        row += 1

wsbyuserandproj.autofilter(5, 0, row - 1, col - 1)




workbook.close()
print(f"Saved: {os.path.join(os.getcwd(), str(workbook.filename))}")

if args.autoopen:
    if sys.platform == "win32":
        os.system("start excel sum.xlsx")
        # subprocess.Popen(["start", "excel", "sum.xlsx"])
    elif sys.platform == "linux":
        subprocess.Popen(["libreoffice", "--calc", "sum.xlsx"])
