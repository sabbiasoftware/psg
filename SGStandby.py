from datetime import datetime as dt, timedelta as td
import calendar
from decimal import Decimal as dec
from SheetGenerator import SheetGenerator
from common import HourFormat, HourType, dec_to_number


class SGStandby(SheetGenerator):
    def __init__(self, config, cellFormats) -> None:
        super().__init__(config, cellFormats)
        self.sumstandby = {}

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
            if (email, project) not in self.sumstandby.keys():
                self.sumstandby[email, project] = {}
            self.sumstandby[email, project][date] = self.sumstandby[email, project].get(date, 0) + dec(row["Hours"])

    def generateHeader(self, worksheet):
        worksheet.write(5, 0, "Name", self.cellFormats["headertxt"])
        worksheet.write(5, 1, "User", self.cellFormats["headertxt"])
        worksheet.write(5, 2, "Approver", self.cellFormats["headertxt"])
        worksheet.write(5, 3, "Project", self.cellFormats["headertxt"])
        worksheet.write(5, 4, "StbyH", self.cellFormats["headernum"])

        date = self.min_date
        lastmonth = 0
        col = 5
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                worksheet.write(4, col, calendar.month_name[date.month], self.cellFormats["headertxt"])
            worksheet.write_number(5, col, date.day, self.cellFormats["headerday"])
            date = date + td(days=1)
            col += 1

        worksheet.set_column(0, 0, width=40)
        worksheet.set_column(1, 2, width=24)
        worksheet.set_column(3, 3, width=48)
        worksheet.set_column(4, 4, width=8)
        worksheet.set_column(5, col - 1, width=6)

    def generateData(self, worksheet):
        row = 6
        col = 5
        for email, project in sorted(self.sumstandby.keys()):
            date = self.min_date
            col = 5
            while date <= self.max_date:
                hours = dec(0)
                if date in self.sumstandby[email, project].keys():
                    hours = self.sumstandby[email, project][date]

                if hours > 0:
                    worksheet.write_number(row, col, dec_to_number(hours), self.cellFormats["hourFormats"][HourFormat.STANDBY])

                date = date + td(days=1)
                col += 1

            worksheet.write(row, 0, email, self.cellFormats["datatxt"])
            worksheet.write(row, 1, self.users[email])
            worksheet.write(row, 2, self.approvers[email])
            worksheet.write(row, 3, project)
            worksheet.write(row, 4, sum([ self.sumstandby[email, project][date] for date in self.sumstandby[email, project].keys() ]))

            row += 1

        worksheet.autofilter(5, 0, row - 1, col - 1)

    def generateSheet(self, workbook):
        worksheet = workbook.add_worksheet("Standby")
        self.generateTitle(worksheet)
        self.generateHeader(worksheet)
        self.generateData(worksheet)
