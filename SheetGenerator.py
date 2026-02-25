from datetime import datetime as dt
from common import HourType, HourFormat, format_date, format_datetime

class SheetGenerator:
    def __init__(self, config, cellFormats):
        self.config = config
        self.cellFormats = cellFormats

        self.users = {}
        self.approvers = {}
        self.min_date = dt.max
        self.max_date = dt.min

    def get_hour_type(self, project, activity) -> HourType:
        if project in self.config["special_projects"].keys():
            return self.config["special_projects"][project]
        elif activity == "Standby Hours - Hungary":
            return HourType.STANDBY
        else:
            return HourType.WORK

    def is_working_day(self, date) -> bool:
        return (
            (date.weekday() < 5)
            and (date not in self.config["weekends"])
            and (date not in self.config["holidays"])
        ) or (date in self.config["workingdays"])

    def get_only_hours(self, hour_type, hours):
        if hour_type not in hours:
            return None
        if sum( [ hours[ht] for ht in hours if ht != hour_type and ht != HourType.STANDBY ] ) > 0:
            return None
        return hours[hour_type]

    def get_active_hours(self, hours):
        return sum( [ hours[ht] for ht in hours if ht != HourType.STANDBY ] )

    def get_day_cell(self, date, hours):
        w = self.get_only_hours(HourType.WORK, hours)
        if self.is_working_day(date):
            if w is not None:
                if w == 8:
                    return w, HourFormat.WORK
                elif w < 8:
                    return w, HourFormat.UNDER
                else:
                    return w, HourFormat.OVER
            elif self.get_only_hours(HourType.VACATION, hours) == 8:
                return "V", HourFormat.VACATION
            elif self.get_only_hours(HourType.SICK, hours) == 8:
                return "S", HourFormat.SICK
            elif self.get_active_hours(hours) == 0:
                return "-", HourFormat.MISS
            else:
                return "?", HourFormat.QUESTION
        else:
            if w is not None and w > 0:
                return w, HourFormat.OVER
            elif self.get_only_hours(HourType.HOLIDAY, hours) == 8:
                return "", HourFormat.EMPTY
            elif self.get_active_hours(hours) == 0:
                return "", HourFormat.EMPTY
            else:
                return "?", HourFormat.QUESTION

    def loadRow(self, row):
        date = None
        try:
            date = dt.strptime(row["Date"], "%Y%m%d")
        except ValueError:
            return

        email = row["Email Address"].lower()

        if email not in self.users.keys():
            self.users[email] = row["User"]
        if email not in self.approvers.keys():
            self.approvers[email] = row["Level 1 Approver Name (configured)"]

        self.min_date = min(self.min_date, date)
        self.max_date = max(self.max_date, date)

    def generateTitle(self, worksheet):
        duration = f"{format_date(self.min_date)}-{format_date(self.max_date)}"
        actual_projects = "All" if len(self.config["projects"]) == 0 else ", ".join(self.config["projects"])
        generated = f"{format_datetime(dt.now())}"
        worksheet.write(0, 0, f"Duration: {duration}")
        worksheet.write(1, 0, f"Projects: {actual_projects}")
        worksheet.write(2, 0, f"Generated: {generated}")

    def generateSheet(self, workbook):
        pass
