import os
import glob
#import subprocess
from datetime import datetime as dt
from datetime import timedelta as td
from decimal import Decimal as dec
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

dec0 = dec("0")
dec8 = dec("8")

IWORK = 0
IVACATION = 1
ISICK = 2
IHOLIDAY = 3

def read_strings(fn, do_strip=False, do_lower=False):
    strings = []
    with open(fn, "r", encoding="utf-8") as f: strings = f.read().splitlines()
    if do_strip:
        strings = [ s.strip() for s in strings ]
    if do_lower:
        strings = [ s.lower() for s in strings ]
    return strings

def read_dates(fn):
    dates = []
    with open(fn, "r") as f: dates = [ dt.strptime(s.strip(), "%Y-%m-%d") for s in f ]
    return dates

def format_hours(hours):
    if type(hours) == int:
        return str(hours)
    elif type(hours) == dec:
        # return "{:.2f}".format(hours)

        if hours.as_integer_ratio()[1] == 1:
            return "{:.0f}".format(hours)
        else:
            return "{:.2f}".format(hours)
    else:
        return "???"


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
        return IWORK

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

        project_included = (rr[iproject] in special_projects) or (len(cfg_projects) == 0) or (max( [str.find(rr[iproject], p) for p in cfg_projects] ) >= 0 ) or (max( [str.find(rr[iprojectdescription], p) for p in cfg_projects] ) >= 0 )
        name_included = (len(cfg_users) == 0) or (name in cfg_users)

        if name_included and project_included:
            entries.append(Entry(
                dt.strptime(rr[idate], "%Y%m%d"),
                name,
                rr[icomment],
                rr[iproject],
                dec(rr[ihours])))
            if rr[iproject] not in special_projects:
                actual_projects.add("{} {}".format(rr[iproject], rr[iprojectdescription]))



daysums = {}

min_date = dt.max
max_date = dt.min

for e in entries:
    if e.name not in daysums.keys():
        daysums[e.name] = {}
    if e.date not in daysums[e.name].keys():
        daysums[e.name][e.date] = [dec0, dec0, dec0, dec0]
    daysums[e.name][e.date][get_hour_index(e.project)] += e.hours

    if e.date < min_date:
        min_date = e.date

    if e.date > max_date:
        max_date = e.date


html_actual_projects = "All"
if len(cfg_projects) > 0:
    html_actual_projects = ", ".join(actual_projects)

html_title = "Duration: {}  &ndash;  {}<br>Projects: {}<br>Generated: {}".format(dt.strftime(min_date, "%Y-%m-%d"), dt.strftime(max_date, "%Y-%m-%d"), html_actual_projects, dt.strftime(dt.now(), "%Y-%m-%d   %H:%M:%S"))



html_days = StringIO()

date = min_date
while date <= max_date:
    html_days.write('<td class="day">{}</td>'.format(date.day))
    date = date + td(days=1)



missinglist = []
overtimelist = []
sicklist = []
vacationlist = []



html_sums = StringIO()

for name in sorted(daysums.keys()):

    total_hours = [dec0, dec0, dec0, dec0]
    for date in daysums[name]:
        for h in range(0,4):
            total_hours[h] += daysums[name][date][h]
    
    # if there are filter projects configured, then filter out people with 0 hours against projects
    if len(cfg_projects) > 0 and total_hours[IWORK] == 0:
        continue

    html_sums.write("<tr>")

    html_sums.write('<td class="name">{}</td>'.format(name))

    missing = False
    overtime_hours = 0

    date = min_date
    while date <= max_date:
        hours = []
        if date in daysums[name].keys():
            hours = daysums[name][date]
        else:
            hours = [dec0, dec0, dec0, dec0]

        if is_working_day(date):
            if hours == [dec8, dec0, dec0, dec0]:
                html_sums.write('<td class="work">{}</td>'.format(format_hours(8)))
            elif hours == [dec0, dec8, dec0, dec0]:
                html_sums.write('<td class="vacation">V</td>')
            elif hours == [dec0, dec0, dec8, dec0]:
                html_sums.write('<td class="sick">S</td>')
            elif hours == [dec0, dec0, dec0, dec0]:
                html_sums.write('<td class="notime">-</td>')
                missing = True
            elif (hours[IWORK] < dec8) and (hours[1:4] == [dec0, dec0, dec0]):
                html_sums.write('<td class="undertime">{}</td>'.format(format_hours(hours[IWORK])))
                missing = True
            elif (hours[IWORK] > dec8) and (hours[1:4] == [dec0, dec0, dec0]):
                html_sums.write('<td class="overtime">{}</td>'.format(format_hours(hours[IWORK])))
                overtime_hours += hours[IWORK] - dec8
            else:
                html_sums.write('<td class="question">?</td>')
                missing = True
        else:
            if (hours[IWORK] > dec0):
                html_sums.write('<td class="overtime">{}</td>'.format(format_hours(hours[IWORK])))
                overtime_hours += hours[IWORK]
            elif hours == [dec0, dec0, dec0, dec0]:
                html_sums.write('<td class="empty"></td>')
            elif hours == [dec0, dec0, dec0, dec8]:
                html_sums.write('<td class="empty"></td>')
            else:
                html_sums.write('<td class="question">?</td>')

        date = date + td(days=1)

    html_sums.write('<td class="total">{}</td>'.format(format_hours((total_hours[IWORK]))))
    html_sums.write('<td class="total">{}</td>'.format(format_hours((total_hours[IWORK] - overtime_hours) // dec8)))
    html_sums.write('<td class="total">{}</td>'.format(format_hours(total_hours[IVACATION] // dec8)))
    html_sums.write('<td class="total">{}</td>'.format(format_hours(total_hours[ISICK] // dec8)))
    html_sums.write('<td class="total">{}</td>'.format(format_hours(overtime_hours)))

    if missing:
        missinglist.append(name)
    
    if overtime_hours > dec0:
        overtimelist.append( (name, overtime_hours) )
    
    if total_hours[ISICK] > dec0:
        sicklist.append( (name, total_hours[ISICK] // dec8) )

    if total_hours[IVACATION] > dec0:
        vacationlist.append( (name, total_hours[IVACATION] // dec8) )

    html_sums.write("</tr>\n")



for name in sorted(cfg_users):
    if name not in daysums.keys():
        html_sums.write("<tr>")

        html_sums.write('<td class="name">{}</td>'.format(name))

        date = min_date
        while date <= max_date:
            if is_working_day(date):
                html_sums.write('<td class="notime">-</td>')
            else:
                html_sums.write('<td class="empty"></td>')

            date = date + td(days=1)

        for h in range(0,4):
            html_sums.write('<td class="total">0</td>')

        missinglist.append(name)

        html_sums.write("</tr>\n")



html_missinglist = str.join("; ", [ "{}@capgemini.com".format(n) for n in sorted(missinglist)])
html_overtimelist = str.join("", [ '<tr><td class="name">{}</td><td class="total">{}</td></tr>\n'.format(n, format_hours(h)) for (n, h) in sorted(overtimelist)])
html_sicklist = str.join("", [ '<tr><td class="name">{}</td><td class="total">{}</td></tr>\n'.format(n, h) for (n, h) in sorted(sicklist)])
html_vacationlist = str.join("", [ '<tr><td class="name">{}</td><td class="total">{}</td></tr>\n'.format(n, h) for (n, h) in sorted(vacationlist)])



html_sum = ""
with open("sum.html", "r") as f: html_sum = f.read()

html_sum = html_sum.replace("***TITLE***", html_title)
html_sum = html_sum.replace("***DAYS***", html_days.getvalue())
html_sum = html_sum.replace("***SUMS***", html_sums.getvalue())
html_sum = html_sum.replace("***MISSINGLIST***", html_missinglist)
html_sum = html_sum.replace("***OVERTIMELIST***", html_overtimelist)
html_sum = html_sum.replace("***SICKLIST***", html_sicklist)
html_sum = html_sum.replace("***VACATIONLIST***", html_vacationlist)

with open("sum_out.html", "w+") as f: f.write(html_sum)

webbrowser.open_new_tab("sum_out.html")
