import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir and script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import subprocess
import glob
import argparse
from datetime import datetime as dt
import userpaths
import csv
import xlsxwriter
from SGByUser import SGByUser
from SGStandbyChanges import SGStandbyChanges
from SGByUserAndProject import SGByUserAndProject
from SGStandby import SGStandby
from SGInfo import SGInfo
from config import Config
from common import HourFormat
import traceback

config = Config()


# returns true if email to be processed
def filter_email(email):
    return len(config.Users) == 0 or email in config.Users or email.replace("@capgemini.com", "") in config.Users


# returns true if project to be processed
def filter_project(project, projectdescription):
    if len(config.Projects) == 0:
        return True

    proj = project.lower()
    projdesc = projectdescription.lower()

    projectmatch = False
    for p in config.Projects:
        if p in proj or p in projdesc:
            projectmatch = True
            break

    return projectmatch


# = {
#     "users": read_strings(os.path.join("cfg", "users.txt"), do_strip=True, do_lower=True),
#     "projects": read_strings(os.path.join("cfg", "projects.txt"), do_lower=True),
#     "holidays": read_dates(os.path.join("cfg", "holidays.txt")),
#     "weekends": read_dates(os.path.join("cfg", "weekends.txt")),
#     "workingdays": read_dates(os.path.join("cfg", "workingdays.txt")),
#     "special_projects": {
#         "Approved Absence (H)": HourType.VACATION,
#         "Vacations": HourType.VACATION,
#         "Sick Leave (H)": HourType.SICK,
#         "Medical Leave": HourType.SICK,
#         "Public Holiday": HourType.HOLIDAY,
#     }
# }

parser = argparse.ArgumentParser("psg - Presence Sheet Generator")
parser.add_argument(
    "-a",
    "--autoopen",
    action="store_true",
    help="automatically open generated file with Excel/Writer",
)
parser.add_argument("-s", "--standbylimit", action="store_true", help="limit monthly standby limit")
parser.add_argument(
    "-m", "--managerfromconfig", action="store_true", help="override manager with manager found in userdata.csv"
)
parser.add_argument(
    "filename",
    nargs="?",
    help="timesheet in CSV format to process; if omitted, latest 'TimesheetReport_*.txt' file is picked from user's default Download folder",
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
    matching_files = glob.glob(os.path.join(userpaths.get_downloads(), "TimesheetReport_*.txt"))
    if len(matching_files) == 0:
        print(f"Could not find any timesheet in folder: {format(downloads_folder)}")
        sys.exit(1)
    inputfilename = max(matching_files, key=os.path.getctime)
print(f"Parsing: {inputfilename}")


workbook = xlsxwriter.Workbook("sum.xlsx")

cellFormats = {
    "headerday": workbook.add_format({"align": "center", "bold": "true"}),
    "headerworkday": workbook.add_format({"align": "center", "bold": "true", "font_color": "black"}),
    "headernonworkday": workbook.add_format({"align": "center", "bold": "true", "font_color": "red"}),
    "headertxt": workbook.add_format({"align": "left", "bold": "true"}),
    "headernum": workbook.add_format({"align": "right", "bold": "true"}),
    "datatxt": workbook.add_format({"align": "left"}),
    "datanum": workbook.add_format({"align": "right"}),
    "hourFormats": {
        HourFormat.WORK: workbook.add_format({"align": "center", "bg_color": "#90ee90"}),
        HourFormat.UNDER: workbook.add_format({"align": "center", "bg_color": "#9acd32"}),
        HourFormat.OVER: workbook.add_format({"align": "center", "bg_color": "#ffa500"}),
        HourFormat.VACATION: workbook.add_format({"align": "center", "bg_color": "#ffff00"}),
        HourFormat.SICK: workbook.add_format({"align": "center", "bg_color": "#da70d6"}),
        HourFormat.MISS: workbook.add_format({"align": "center", "bg_color": "#d3d3d3"}),
        HourFormat.QUESTION: workbook.add_format({"align": "center", "bg_color": "#808080"}),
        HourFormat.EMPTY: workbook.add_format({"align": "center", "bg_color": "#ffffff"}),
        HourFormat.STANDBY: workbook.add_format({"align": "center", "bg_color": "#9090ee"}),
    },
}

sheetGenerators = [
    SGByUser(config, cellFormats, args.standbylimit, args.managerfromconfig),
    SGStandbyChanges(config, cellFormats, args.standbylimit, args.managerfromconfig),
    SGByUserAndProject(config, cellFormats, args.managerfromconfig),
    SGStandby(config, cellFormats, args.managerfromconfig),
    SGInfo(config, cellFormats, False),
]

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

            for g in sheetGenerators:
                g.loadRow(row)

except Exception as exc:
    print(f"Could not parse input: {inputfilename}")
    print(f"Exception: {type(exc)}, Arguments: {exc.args}")
    traceback.print_exc()
    sys.exit(1)

for g in sheetGenerators:
    g.generateSheet(workbook)

workbook.close()
print(f"Saved: {os.path.join(os.getcwd(), str(workbook.filename))}")

if args.autoopen:
    if sys.platform == "win32":
        os.system("start excel sum.xlsx")
        # subprocess.Popen(["start", "excel", "sum.xlsx"])
    elif sys.platform == "linux":
        subprocess.Popen(["libreoffice", "--calc", "sum.xlsx"])

# print(sys.path)
