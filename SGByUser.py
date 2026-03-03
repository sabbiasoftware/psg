from datetime import datetime as dt, timedelta as td
import calendar
from decimal import Decimal as dec
from SheetGenerator import SheetGenerator
from common import HourType, HourFormat, dec_to_number
from config import Config

class SGByUser(SheetGenerator):
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


    def forceStandbyLimit(self):
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
                        if (date.year == year and date.month == month):
                            if not self.is_working_day(date):
                                monthlystandy -= self.tryConvertingStandbyToWork(email, date)
                                if monthlystandy <= self.MONTHLYSTANDBYLIMIT:
                                    break

                    # if limit is still exceeded, try converting standby hours to overtime on weekdays
                    if monthlystandy > self.MONTHLYSTANDBYLIMIT:
                        for date in self.sumbyuser[email]:
                            if (date.year == year and date.month == month):
                                if self.is_working_day(date):
                                    monthlystandy -= self.tryConvertingStandbyToWork(email, date)
                                    if monthlystandy <= self.MONTHLYSTANDBYLIMIT:
                                        break

                month += 1
                if month == 13:
                    year += 1
                    month = 1

    def generateHeaderDays(self, worksheet, row, col):
        date = self.min_date
        lastmonth = 0
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                worksheet.write(row - 1, col, calendar.month_name[date.month], self.cellFormats["headertxt"])
            cf = self.cellFormats["headerworkday"] if self.is_working_day(date) else self.cellFormats["headernonworkday"]
            worksheet.write_number(row, col, date.day, cf)
            date = date + td(days=1)
            col += 1


    def generateHeader(self, worksheet):
        worksheet.write(5, 0, "Name", self.cellFormats["headertxt"])
        worksheet.write(5, 1, "User", self.cellFormats["headertxt"])
        worksheet.write(5, 2, "Approver", self.cellFormats["headertxt"])
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
                total_hours[ht] = sum( [ self.sumbyuser[email][date].get(ht, 0) for date in self.sumbyuser[email] ] )

            # if there are filter projects configured, then filter out people with 0 hours against projects
            if len(self.config.Projects) > 0 and total_hours[HourType.WORK] == 0:
                continue

            weekday_overtime_hours = sum( [ self.sumbyuser[email][date].get(HourType.WORK, 0) - 8 for date in self.sumbyuser[email] if self.is_working_day(date) and self.sumbyuser[email][date].get(HourType.WORK, 0) > 8 ] )
            weekend_overtime_hours = sum( [ self.sumbyuser[email][date].get(HourType.WORK, 0) for date in self.sumbyuser[email] if not self.is_working_day(date) ] )
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

            worksheet.write(row, 0, email, self.cellFormats["datatxt"])
            worksheet.write(row, 1, self.users[email])
            worksheet.write(row, 2, self.approvers[email])
            worksheet.write_number(row, 3, int(total_hours[HourType.WORK]), self.cellFormats["datanum"])
            worksheet.write_number(row, 4, int((total_hours[HourType.WORK] - overtime_hours) // dec(8)), self.cellFormats["datanum"])
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


    def generateStandbyAdjustmentTitle(self, worksheet):
        row = len(sorted(self.sumbyuser.keys())) + 9

        if not self.standbylimit:
            worksheet.write(row, 0, "Forcing monthly standby limit disabled")
            return

        if len(self.sumstandbydec) == 0 and len(self.sumworkinc) == 0:
            worksheet.write(row, 0, f"Forcing monthly standby limit enabled, but monthly limit of {self.MONTHLYSTANDBYLIMIT} was not exceeded")
            return

        worksheet.write(row, 0, f"Forcing monthly standby limit enabled, monthly limit of {self.MONTHLYSTANDBYLIMIT} was exceeded")


    def generateStandbyAdjustmentHeader(self, worksheet):
        row = len(sorted(self.sumbyuser.keys())) + 10
        worksheet.write(row, 0, "Name", self.cellFormats["headertxt"])
        worksheet.write(row, 1, "User", self.cellFormats["headertxt"])
        worksheet.write(row, 2, "Approver", self.cellFormats["headertxt"])
        worksheet.write(row, 3, "Comment", self.cellFormats["headertxt"])
        self.generateHeaderDays(worksheet, row, 9)


    def generateStandbyAdjustmentData(self, worksheet):
        row = len(sorted(self.sumbyuser.keys())) + 11

        for email in sorted(self.sumstandbydec):

            worksheet.write(row + 0, 3, "Standby before")
            worksheet.write(row + 1, 3, "Standby reduction")
            worksheet.write(row + 2, 3, "Standby after")
            worksheet.write(row + 3, 3, "Work before")
            worksheet.write(row + 4, 3, "Work addition")
            worksheet.write(row + 5, 3, "Work after")

            date = self.min_date
            col = 9
            while date <= self.max_date:
                standbyafter = 0
                workafter = 0

                if date in self.sumbyuser[email].keys():
                    standbyafter = self.sumbyuser[email][date].get(HourType.STANDBY, 0)
                    if standbyafter > 0:
                        worksheet.write_number(row + 0, col, standbyafter, self.cellFormats["hourFormats"][HourFormat.STANDBY])
                        worksheet.write_number(row + 2, col, standbyafter, self.cellFormats["hourFormats"][HourFormat.STANDBY])

                    workafter = self.sumbyuser[email][date].get(HourType.WORK, 0)
                    if workafter > 0:
                        worksheet.write_number(row + 3, col, workafter, self.cellFormats["hourFormats"][HourFormat.WORK])
                        worksheet.write_number(row + 5, col, workafter, self.cellFormats["hourFormats"][HourFormat.WORK])

                if date in self.sumstandbydec[email].keys():
                    standbyreduction = self.sumstandbydec[email][date]
                    worksheet.write_number(row + 1, col, standbyreduction, self.cellFormats["hourFormats"][HourFormat.EMPTY])
                    worksheet.write_number(row + 0, col, standbyafter - standbyreduction, self.cellFormats["hourFormats"][HourFormat.STANDBY] )

                if date in self.sumworkinc[email].keys():
                    workaddition = self.sumworkinc[email][date]
                    worksheet.write_number(row + 4, col, workaddition, self.cellFormats["hourFormats"][HourFormat.EMPTY])
                    if workafter - workaddition >= 0:
                        worksheet.write_number(row + 3, col, workafter - workaddition, self.cellFormats["hourFormats"][HourFormat.WORK])

                date = date + td(days=1)
                col += 1

            for rd in range(0, 6):
                worksheet.write(row + rd, 0, email, self.cellFormats["datatxt"])
                worksheet.write(row + rd, 1, self.users[email])
                worksheet.write(row + rd, 2, self.approvers[email])
            s = sum([ self.sumstandbydec[email][date] for date in self.sumstandbydec[email].keys() ])
            worksheet.write_number(row, 7, s, self.cellFormats["datanum"])

            row += 6

        # worksheet.autofilter(len(sorted(self.sumbyuser.keys())) + 10, 0, row - 6, 9 + (self.max_date - self.min_date).days)


    def generateSheet(self, workbook):
        if self.standbylimit:
            self.forceStandbyLimit()
        worksheet = workbook.add_worksheet("By user")
        self.generateTitle(worksheet)
        self.generateHeader(worksheet)
        self.generateData(worksheet)
        self.generateStandbyAdjustmentTitle(worksheet)
        self.generateStandbyAdjustmentHeader(worksheet)
        self.generateStandbyAdjustmentData(worksheet)
