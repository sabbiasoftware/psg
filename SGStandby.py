from datetime import datetime as dt, timedelta as td
import calendar
from decimal import Decimal as dec
from SheetGenerator import SheetGenerator
from common import HourFormat, HourType, dec_to_number
from config import Config


class SGStandby(SheetGenerator):
    def __init__(self, config: Config, cellFormats) -> None:
        super().__init__(config, cellFormats)
        self.sumstandby = {}
        self.sumhotline = {}

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

        if t is HourType.STANDBY:
            proj = row["Project"]
            desc = row["Project Description"]
            project = f"{proj} {desc}" if proj != desc else proj
            hotline = self.config.Hotlines.get(email, "?")
            if (hotline, email, project) not in self.sumstandby.keys():
                self.sumstandby[hotline, email, project] = {}
            self.sumstandby[hotline, email, project][date] = self.sumstandby[hotline, email, project].get(date, 0) + dec(row["Hours"])

            if hotline not in self.sumhotline.keys():
                self.sumhotline[hotline] = {}
            self.sumhotline[hotline][date] = self.sumhotline[hotline].get(date, 0) + dec(row["Hours"])

    def generateHeader(self, worksheet):
        worksheet.write(5, 0, "Hotline", self.cellFormats["headertxt"])
        worksheet.write(5, 1, "Name", self.cellFormats["headertxt"])
        worksheet.write(5, 2, "User", self.cellFormats["headertxt"])
        worksheet.write(5, 3, "Approver", self.cellFormats["headertxt"])
        worksheet.write(5, 4, "Project", self.cellFormats["headertxt"])
        worksheet.write(5, 5, "StbyH", self.cellFormats["headernum"])

        date = self.min_date
        lastmonth = 0
        col = 6
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                worksheet.write(4, col, calendar.month_name[date.month], self.cellFormats["headertxt"])
            worksheet.write_number(5, col, date.day, self.cellFormats["headerday"])
            date = date + td(days=1)
            col += 1

        worksheet.set_column(0, 0, width=24)
        worksheet.set_column(1, 1, width=40)
        worksheet.set_column(2, 3, width=24)
        worksheet.set_column(4, 4, width=48)
        worksheet.set_column(5, 5, width=8)
        worksheet.set_column(6, col - 1, width=6)

    def generateData(self, worksheet):
        row = 6
        col = 5
        for hotline, email, project in sorted(self.sumstandby.keys()):
            date = self.min_date
            col = 6
            while date <= self.max_date:
                hours = dec(0)
                if date in self.sumstandby[hotline, email, project].keys():
                    hours = self.sumstandby[hotline, email, project][date]

                if hours > 0:
                    worksheet.write_number(row, col, dec_to_number(hours), self.cellFormats["hourFormats"][HourFormat.STANDBY])

                date = date + td(days=1)
                col += 1

            worksheet.write(row, 0, hotline)
            worksheet.write(row, 1, email, self.cellFormats["datatxt"])
            worksheet.write(row, 2, self.users[email])
            worksheet.write(row, 3, self.approvers[email])
            worksheet.write(row, 4, project)
            worksheet.write(row, 5, sum([ self.sumstandby[hotline, email, project][date] for date in self.sumstandby[hotline, email, project].keys() ]))

            row += 1

        worksheet.autofilter(5, 0, row - 1, col - 1)


        row += 4
        for hotline in sorted(self.sumhotline.keys()):
            date = self.min_date
            col = 6
            while date <= self.max_date:
                hours = self.sumhotline[hotline].get(date, dec(0))

                if hours > 0:
                    expectedHours = 16 if self.is_working_day(date) else 24
                    hourFormat = HourFormat.WORK
                    if hours < expectedHours:
                        hourFormat = HourFormat.UNDER
                    elif hours > expectedHours:
                        hourFormat = HourFormat.OVER
                    worksheet.write_number(row, col, dec_to_number(hours), self.cellFormats["hourFormats"][hourFormat])

                date = date + td(days=1)
                col += 1

            worksheet.write(row, 0, hotline)
            worksheet.write(row, 5, sum([ self.sumhotline[hotline][date] for date in self.sumhotline[hotline].keys() ]))

            row += 1


    def generateSheet(self, workbook):
        worksheet = workbook.add_worksheet("Standby")
        self.generateTitle(worksheet)
        self.generateHeader(worksheet)
        self.generateData(worksheet)
