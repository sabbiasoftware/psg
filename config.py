import os
from common import read_strings, read_dates, HourType

class Config:
    def __init__(self):
        self.Users = read_strings(os.path.join("cfg", "users.txt"), do_strip=True, do_lower=True)
        self.Projects = read_strings(os.path.join("cfg", "projects.txt"), do_lower=True)
        self.Holidays = read_dates(os.path.join("cfg", "holidays.txt"))
        self.Weekends = read_dates(os.path.join("cfg", "weekends.txt"))
        self.Workingdays = read_dates(os.path.join("cfg", "workingdays.txt"))
        self.SpecialProjects = {
            "Approved Absence (H)": HourType.VACATION,
            "Vacations": HourType.VACATION,
            "Sick Leave (H)": HourType.SICK,
            "Medical Leave": HourType.SICK,
            "Public Holiday": HourType.HOLIDAY
        }

        self.Hotlines = {}
        hotlines = read_strings(os.path.join("cfg", "hotlines.txt"), do_strip=True)
        for row in hotlines:
            fields = row.split(",")
            if len(fields) != 2:
                continue
            hotline = fields[0]
            email = fields[1].lower()
            if email not in self.Hotlines.keys():
                self.Hotlines[email] = hotline
