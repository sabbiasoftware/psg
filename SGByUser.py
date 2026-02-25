from datetime import datetime as dt, timedelta as td
import calendar
from decimal import Decimal as dec
from SheetGenerator import SheetGenerator
from common import HourType, HourFormat, format_hours, format_date

class SGByUser(SheetGenerator):
    MONTHLYSTANDBYLIMIT = 168

    def __init__(self, config, cellFormats, standbylimit) -> None:
        super().__init__(config, cellFormats)
        self.sumbyuser = {}
        self.standbylimit = standbylimit

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
                    print(f"Standby hours of {format_hours(monthlystandy)} in {year}-{month:02d} exceeds {self.MONTHLYSTANDBYLIMIT} hours for {email}")
                    for date in self.sumbyuser[email]:
                        if (
                            date.year == year
                            and date.month == month
                        ):
                            otmulti = 1 if self.is_working_day(date) else 2
                            if self.sumbyuser[email][date].get(HourType.STANDBY, 0) >= 5 * otmulti:
                                w1 = self.sumbyuser[email][date].get(HourType.WORK, 0)
                                s1 = self.sumbyuser[email][date].get(HourType.STANDBY, 0)
                                pluswork = self.sumbyuser[email][date][HourType.STANDBY] // (5 * otmulti)
                                minusstandby = pluswork * 5 * otmulti
                                w2 = w1 + pluswork
                                s2 = s1 - minusstandby
                                self.sumbyuser[email][date][HourType.WORK] = w2
                                self.sumbyuser[email][date][HourType.STANDBY] = s2
                                print(f"  {format_date(date)}, {email}: ({format_hours(w1)}, {format_hours(s1)}) + (+{format_hours(pluswork)}, -{format_hours(minusstandby)}) = ({format_hours(w2)}, {format_hours(s2)})")
                                monthlystandy -= minusstandby
                                if monthlystandy <= self.MONTHLYSTANDBYLIMIT:
                                    print(f"  Standby hours in {year}-{month:02d} is {format_hours(monthlystandy)} hours for {email}")
                                    break

                month += 1
                if month == 13:
                    year += 1
                    month = 1

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

        date = self.min_date
        lastmonth = 0
        col = 9
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                worksheet.write(4, col, calendar.month_name[date.month], self.cellFormats["headertxt"])
            worksheet.write_number(5, col, date.day, self.cellFormats["headerday"])
            date = date + td(days=1)
            col += 1

        worksheet.set_column(0, 0, width=40)
        worksheet.set_column(1, 2, width=24)
        worksheet.set_column(3, 8, width=8)
        worksheet.set_column(9, col - 1, width=6)

    def generateData(self, worksheet):
        row = 6
        col = 9
        for email in sorted(self.sumbyuser.keys()):
            total_hours = {}
            for ht in HourType:
                total_hours[ht] = sum( [ self.sumbyuser[email][date].get(ht, 0) for date in self.sumbyuser[email] ] )

            # if there are filter projects configured, then filter out people with 0 hours against projects
            if len(self.config["projects"]) > 0 and total_hours[HourType.WORK] == 0:
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


        for email in sorted(self.config["users"]):
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
