from datetime import datetime as dt, timedelta as td
import calendar
from decimal import Decimal as dec
from config import Config
from common import HourType
from SheetGenerator import SheetGenerator


class SGStandbyLimiter(SheetGenerator):
    MONTHLYSTANDBYLIMIT = 168

    def __init__(self, config: Config, cellFormats, standbylimit) -> None:
        super().__init__(config, cellFormats)
        self.standbylimit = standbylimit
        self.sumbyuser = {}
        self.sumstandbydec = {}
        self.sumworkinc = {}

    def loadRow(self, row):
        super().loadRow(row)

        # Assumption: if first field can be parsed as date, then it is a data row
        date = None
        try:
            date = dt.strptime(row["Date"], "%Y%m%d")
        except ValueError:
            return

        email = row["Email Address"].lower()

        t = self.get_hour_type(row["Project"], row["Activity"])

        # sumbyuser update
        if email not in self.sumbyuser.keys():
            self.sumbyuser[email] = {}
        if date not in self.sumbyuser[email].keys():
            self.sumbyuser[email][date] = {}
        self.sumbyuser[email][date][t] = self.sumbyuser[email][date].get(t, 0) + dec(row["Hours"])

    def tryConvertingStandbyToWork(self, email, date):
        minusstandby = 0

        # exchange rate between standby and overtime hours
        # weekday: 15 standby = 2 overtime
        # weekend: 20 standby = 1 overtime
        sbyunit = 15 if self.is_working_day(date) else 10
        ovtunit = 2 if self.is_working_day(date) else 1

        if self.sumbyuser[email][date].get(HourType.STANDBY, 0) >= sbyunit:
            w1 = self.sumbyuser[email][date].get(HourType.WORK, 0)
            s1 = self.sumbyuser[email][date].get(HourType.STANDBY, 0)
            numunit = s1 // sbyunit
            pluswork = numunit * ovtunit
            minusstandby = numunit * sbyunit
            w2 = w1 + pluswork
            s2 = s1 - minusstandby
            self.sumbyuser[email][date][HourType.WORK] = w2
            self.sumbyuser[email][date][HourType.STANDBY] = s2

            if email not in self.sumstandbydec.keys():
                self.sumstandbydec[email] = {}
            self.sumstandbydec[email][date] = self.sumstandbydec[email].get(date, 0) - minusstandby

            if email not in self.sumworkinc.keys():
                self.sumworkinc[email] = {}
            self.sumworkinc[email][date] = self.sumworkinc[email].get(date, 0) + pluswork

        return minusstandby

    def limitStandby(self):
        for email in self.sumbyuser:
            # avoid importing dateutil.relativedelta for now...
            year = self.min_date.year
            month = self.min_date.month
            while year < self.max_date.year or (year == self.max_date.year and month <= self.max_date.month):
                monthlystandy = sum(
                    [
                        self.sumbyuser[email][d].get(HourType.STANDBY, 0)
                        for d in self.sumbyuser[email]
                        if d.year == year and d.month == month
                    ]
                )

                if monthlystandy > self.MONTHLYSTANDBYLIMIT:
                    # first try converting standby hours to overtime on weekends
                    for date in self.sumbyuser[email]:
                        if date.year == year and date.month == month:
                            if not self.is_working_day(date):
                                monthlystandy -= self.tryConvertingStandbyToWork(email, date)
                                if monthlystandy <= self.MONTHLYSTANDBYLIMIT:
                                    break

                    # if limit is still exceeded, try converting standby hours to overtime on weekdays
                    if monthlystandy > self.MONTHLYSTANDBYLIMIT:
                        for date in self.sumbyuser[email]:
                            if date.year == year and date.month == month:
                                if self.is_working_day(date):
                                    monthlystandy -= self.tryConvertingStandbyToWork(email, date)
                                    if monthlystandy <= self.MONTHLYSTANDBYLIMIT:
                                        break

                month += 1
                if month == 13:
                    year += 1
                    month = 1

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
