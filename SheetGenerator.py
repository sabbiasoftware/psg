from abc import abstractmethod
from datetime import datetime as dt, timedelta as td
import calendar
from common import HourType, HourFormat
from config import Config


class SheetGenerator:
    def __init__(self, config: Config, cellFormats):
        self.config = config
        self.cellFormats = cellFormats

        self.users = {}
        self.approvers = {}
        self.min_date = dt.max
        self.max_date = dt.min

    def get_hour_type(self, project, activity) -> HourType:
        if project in self.config.SpecialProjects.keys():
            return self.config.SpecialProjects[project]
        elif activity == "Standby Hours - Hungary":
            return HourType.STANDBY
        else:
            return HourType.WORK

    def is_working_day(self, date) -> bool:
        return ((date.weekday() < 5) and (date not in self.config.Weekends) and (date not in self.config.Holidays)) or (
            date in self.config.Workingdays
        )

    def get_only_hours(self, hour_type, hours):
        if hour_type not in hours:
            return None
        if sum([hours[ht] for ht in hours if ht != hour_type and ht != HourType.STANDBY]) > 0:
            return None
        return hours[hour_type]

    def get_active_hours(self, hours):
        return sum([hours[ht] for ht in hours if ht != HourType.STANDBY])

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

    @abstractmethod
    def generateSheet(self, workbook):
        pass

    def generateColumnHeader(self, worksheet, row, col, headerText, headerFormat, width):
        worksheet.write(row, col, headerText, headerFormat)
        worksheet.set_column(col, col, width=width)

    def generateCommonColumnHeaders(self, worksheet, row, col):
        self.generateColumnHeader(worksheet, row, col + 0, "Email", self.cellFormats["headertxt"], 40)
        self.generateColumnHeader(worksheet, row, col + 1, "Name", self.cellFormats["headertxt"], 24)
        self.generateColumnHeader(worksheet, row, col + 2, "Manager", self.cellFormats["headertxt"], 24)

    def generateHeaderDays(self, worksheet, row, col):
        date = self.min_date
        lastmonth = 0
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                worksheet.write(row - 1, col, calendar.month_name[date.month], self.cellFormats["headertxt"])
            cf = (
                self.cellFormats["headerworkday"] if self.is_working_day(date) else self.cellFormats["headernonworkday"]
            )
            self.generateColumnHeader(worksheet, row, col, date.day, cf, 6)
            date = date + td(days=1)
            col += 1
