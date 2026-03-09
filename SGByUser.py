from datetime import timedelta as td
import calendar
from decimal import Decimal as dec
from SGStandbyLimiter import SGStandbyLimiter
from common import HourType, HourFormat, dec_to_number
from config import Config


class SGByUser(SGStandbyLimiter):
    def __init__(self, config: Config, cellFormats, standbylimit, managerFromConfig) -> None:
        super().__init__(config, cellFormats, standbylimit)
        self.managerFromConfig = managerFromConfig

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
            worksheet.write_number(row, col, date.day, cf)
            date = date + td(days=1)
            col += 1

    def generateHeader(self, worksheet):
        worksheet.write(5, 0, "Name", self.cellFormats["headertxt"])
        worksheet.write(5, 1, "User", self.cellFormats["headertxt"])
        worksheet.write(5, 2, "Manager", self.cellFormats["headertxt"])
        worksheet.write(5, 3, "WorkH", self.cellFormats["headernum"])
        worksheet.write(5, 4, "WorkD", self.cellFormats["headernum"])
        worksheet.write(5, 5, "VacaD", self.cellFormats["headernum"])
        worksheet.write(5, 6, "SickD", self.cellFormats["headernum"])
        worksheet.write(5, 7, "OverH", self.cellFormats["headernum"])
        worksheet.write(5, 8, "StbyH", self.cellFormats["headernum"])

        self.generateHeaderDays(worksheet, 5, 9)

        worksheet.set_column(0, 0, width=40)
        worksheet.set_column(1, 2, width=24)
        worksheet.set_column(3, 8, width=8)
        worksheet.set_column(9, 9 + (self.max_date - self.min_date).days, width=6)

    def generateData(self, worksheet):
        row = 6
        col = 9
        for email in sorted(self.sumbyuser.keys()):
            total_hours = {}
            for ht in HourType:
                total_hours[ht] = sum([self.sumbyuser[email][date].get(ht, 0) for date in self.sumbyuser[email]])

            # if there are filter projects configured, then filter out people with 0 hours against projects
            if len(self.config.Projects) > 0 and total_hours[HourType.WORK] == 0:
                continue

            weekday_overtime_hours = sum(
                [
                    self.sumbyuser[email][date].get(HourType.WORK, 0) - 8
                    for date in self.sumbyuser[email]
                    if self.is_working_day(date) and self.sumbyuser[email][date].get(HourType.WORK, 0) > 8
                ]
            )
            weekend_overtime_hours = sum(
                [
                    self.sumbyuser[email][date].get(HourType.WORK, 0)
                    for date in self.sumbyuser[email]
                    if not self.is_working_day(date)
                ]
            )
            overtime_hours = weekday_overtime_hours + weekend_overtime_hours

            date = self.min_date
            col = 9
            while date <= self.max_date:
                hours = {}
                if date in self.sumbyuser[email].keys():
                    hours = self.sumbyuser[email][date]

                value, format = self.get_day_cell(date, hours)
                if isinstance(value, int) or isinstance(value, float):
                    worksheet.write_number(row, col, value, self.cellFormats["hourFormats"][format])
                elif isinstance(value, dec):
                    worksheet.write_number(row, col, dec_to_number(value), self.cellFormats["hourFormats"][format])
                else:
                    worksheet.write(row, col, value, self.cellFormats["hourFormats"][format])

                date = date + td(days=1)
                col += 1

            manager = self.approvers[email]
            if self.managerFromConfig:
                if email in self.config.UserData.keys():
                    manager = self.config.UserData[email]["Reporting to"]

            worksheet.write(row, 0, email, self.cellFormats["datatxt"])
            worksheet.write(row, 1, self.users[email])
            worksheet.write(row, 2, manager)
            worksheet.write_number(row, 3, int(total_hours[HourType.WORK]), self.cellFormats["datanum"])
            worksheet.write_number(
                row, 4, int((total_hours[HourType.WORK] - overtime_hours) // dec(8)), self.cellFormats["datanum"]
            )
            worksheet.write_number(row, 5, int(total_hours[HourType.VACATION] // dec(8)), self.cellFormats["datanum"])
            worksheet.write_number(row, 6, int(total_hours[HourType.SICK] // dec(8)), self.cellFormats["datanum"])
            worksheet.write_number(row, 7, int(overtime_hours), self.cellFormats["datanum"])
            worksheet.write_number(row, 8, int(total_hours[HourType.STANDBY]), self.cellFormats["datanum"])

            row += 1

        for email in sorted(self.config.Users):
            if email not in self.sumbyuser.keys() and f"{email}@capgemini.com" not in self.sumbyuser.keys():
                worksheet.write(row, 0, email, self.cellFormats["datatxt"])

                date = self.min_date
                col = 9
                while date <= self.max_date:
                    if self.is_working_day(date):
                        worksheet.write(row, col, "-", self.cellFormats["hourFormats"][HourFormat.MISS])
                    else:
                        worksheet.write(row, col, "", self.cellFormats["hourFormats"][HourFormat.EMPTY])

                    date = date + td(days=1)
                    col += 1

                for c in range(1, 7):
                    worksheet.write(row, c, 0, self.cellFormats["datanum"])

                row += 1

        worksheet.autofilter(5, 0, row - 1, col - 1)

    def generateSheet(self, workbook):
        if self.standbylimit:
            self.forceStandbyLimit()
        worksheet = workbook.add_worksheet("By user")
        self.generateTitle(worksheet)
        self.generateHeader(worksheet)
        self.generateData(worksheet)
