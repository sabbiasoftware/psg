import os
import glob
#import subprocess
from datetime import datetime as dt
from datetime import timedelta as td
from io import StringIO
import webbrowser

class Entry:
    def __init__(self, date, name, comment, project, hours):
        self.date = date
        self.name = name
        self.comment = comment
        self.project = project
        self.hours = hours
    
    def __str__(self) -> str:
        return "{} {} {} {} {}".format(dt.strftime(self.date, "%Y%m%d"), self.name, self.comment, self.project, self.hours)

IWORK = 0
IVACATION = 1
ISICK = 2
IHOLIDAY = 3

def read_strings(fn, do_strip=False, do_lower=False):
    strings = []
    with open(fn, "r") as f: strings = f.read().splitlines()
    if do_strip:
        strings = [ s.strip() for s in strings ]
    if do_lower:
        strings = [ s.lower() for s in strings ]
    return strings

def read_dates(fn):
    dates = []
    with open(fn, "r") as f: dates = [ dt.strptime(s.strip(), "%Y-%m-%d") for s in f ]
    return dates

special_projects = {}
special_projects["Approved Absence (H)"] = IVACATION
special_projects["Vacations"] = IVACATION
special_projects["Sick Leave (H)"] = ISICK
special_projects["Medical Leave"] = ISICK
special_projects["Public Holiday"] = IHOLIDAY

def get_hour_index(project) -> int:
    if project in special_projects.keys():
        return special_projects[project]
    else:
        return 0

def is_working_day(date) -> bool:
    return ((date.weekday() < 5) and (date not in cfg_weekends) and (date not in cfg_holidays)) or (date in cfg_workingdays)



cfg_users = read_strings("cfg\\users.txt", do_strip=True, do_lower=True)
cfg_holidays = read_dates("cfg\\holidays.txt")
cfg_weekends = read_dates("cfg\\weekends.txt")
cfg_workingdays = read_dates("cfg\\workingdays.txt")
cfg_projects = read_strings("cfg\\projects.txt")




downloads_folder = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')
raw_files = glob.glob(os.path.join(downloads_folder, 'TimesheetReport_*'))

if len(raw_files) == 0:
    print("Could not find any timesheet in {}".format(downloads_folder))
    input()
    quit()

latest_raw_filename = max(raw_files, key=os.path.getctime)
raw = read_strings(latest_raw_filename)
#raw = subprocess.getoutput("powershell.exe -Command Get-Clipboard").splitlines()
#raw = read_strings("raw.txt")



header = raw[0].split("\t")
idate = header.index("Date")
iname = header.index("Email Address")
icomment = header.index("Comment")
iproject = header.index("Project")
iprojectdescription = header.index("Project Description")
ihours = header.index("Hours")

entries = []
actual_projects = set()

for r in raw:
    if r[0:2] == "20":
        rr = r.split("\t")

        name = rr[iname][0:-14].strip().lower()

        project_included = (len(cfg_projects) == 0) or (max( [str.find(rr[iproject], p) for p in cfg_projects] ) >= 0 ) or (max( [str.find(rr[iprojectdescription], p) for p in cfg_projects] ) >= 0 )
        name_included = (len(cfg_users) == 0) or (name in cfg_users)

        if name_included and project_included:
            entries.append(Entry(
                dt.strptime(rr[idate], "%Y%m%d"),
                name,
                rr[icomment],
                rr[iproject],
                int(float(rr[ihours]))))
            if rr[iproject] not in special_projects:
                actual_projects.add("{} {}".format(rr[iproject], rr[iprojectdescription]))



projectsums = {}

min_date = dt.max
max_date = dt.min

for e in entries:
    if e.name not in projectsums.keys():
        projectsums[e.name] = {}
    
    month = dt(e.date.year, e.date.month, 1)

    if month not in projectsums[e.name].keys():
        projectsums[e.name][month] = 0

    projectsums[e.name][month] += e.hours

    if e.date < min_date:
        min_date = e.date

    if e.date > max_date:
        max_date = e.date



html_actual_projects = "All"
if len(cfg_projects) > 0:
    html_actual_projects = ", ".join(actual_projects)

html_title = "Duration: {}  &ndash;  {}<br>Projects: {}<br>Generated: {}".format(dt.strftime(min_date, "%Y-%m-%d"), dt.strftime(max_date, "%Y-%m-%d"), html_actual_projects, dt.strftime(dt.now(), "%Y-%m-%d   %H:%M:%S"))



html_months = StringIO()

month = dt(min_date.year, min_date.month, 1)
while month <= max_date:
    html_months.write('<td class="day">{}</td>'.format(month.month))
    # fixme
    month = month + td(days=31)




html_projectsums = StringIO()

for name in sorted(projectsums.keys()):
    html_projectsums.write("<tr>")

    html_projectsums.write('<td class="name">{}</td>'.format(name))

    month = dt(min_date.year, min_date.month, 1)
    while month <= max_date:
        hour = 0
        if month in projectsums[name].keys():
            hour = projectsums[name][month]

        html_projectsums.write('<td class="work">{}</td>'.format(hour))

        # fixme
        m =  month + td(days=31)
        month = dt(m.year, m.month, 1)

    html_projectsums.write("</tr>")



project_sum = ""
with open("project_sum.html", "r") as f: project_sum = f.read()

project_sum = project_sum.replace("***TITLE***", html_title)
project_sum = project_sum.replace("***MONTHS***", html_months.getvalue())
project_sum = project_sum.replace("***PROJECTSUMS***", html_projectsums.getvalue())

with open("project_sum_out.html", "w+") as f: f.write(project_sum)

webbrowser.open_new_tab("project_sum_out.html")
