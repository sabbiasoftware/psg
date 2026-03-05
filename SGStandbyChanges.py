from datetime import timedelta as td
import calendar
from common import HourType, HourFormat
from SGStandbyLimiter import SGStandbyLimiter


class SGStandbyChanges(SGStandbyLimiter):
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

    def generateTitle(self, worksheet):
        row = 0

        if not self.standbylimit:
            worksheet.write(row, 0, "Forcing monthly standby limit disabled")
            return

        if len(self.sumstandbydec) == 0 and len(self.sumworkinc) == 0:
            worksheet.write(
                row,
                0,
                f"Forcing monthly standby limit enabled, but monthly limit of {self.MONTHLYSTANDBYLIMIT} was not exceeded",
            )
            return

        worksheet.write(
            row, 0, f"Forcing monthly standby limit enabled, monthly limit of {self.MONTHLYSTANDBYLIMIT} was exceeded"
        )

    def generateHeader(self, worksheet):
        row = 1
        worksheet.write(row, 0, "Name", self.cellFormats["headertxt"])
        worksheet.write(row, 1, "User", self.cellFormats["headertxt"])
        worksheet.write(row, 2, "Approver", self.cellFormats["headertxt"])
        worksheet.write(row, 3, "Comment", self.cellFormats["headertxt"])
        worksheet.write(row, 4, "StbyH-", self.cellFormats["headernum"])
        worksheet.write(row, 5, "Work+", self.cellFormats["headernum"])
        self.generateHeaderDays(worksheet, row, 6)

        worksheet.set_column(0, 0, width=40)
        worksheet.set_column(1, 3, width=24)
        worksheet.set_column(4, 5, width=8)
        worksheet.set_column(6, 6 + (self.max_date - self.min_date).days, width=6)

    def getWorkCellFormat(self, date, hours):
        normhours = 8 if self.is_working_day(date) else 0
        if hours < normhours:
            return self.cellFormats["hourFormats"][HourFormat.UNDER]
        elif hours > normhours:
            return self.cellFormats["hourFormats"][HourFormat.OVER]
        else:
            return self.cellFormats["hourFormats"][HourFormat.WORK]

    def generateData(self, worksheet):
        row = 2

        for email in sorted(self.sumstandbydec):
            worksheet.write(row + 0, 3, "Standby before")
            worksheet.write(row + 1, 3, "Standby reduction")
            worksheet.write(row + 2, 3, "Standby after")
            worksheet.write(row + 3, 3, "Work before")
            worksheet.write(row + 4, 3, "Work addition")
            worksheet.write(row + 5, 3, "Work after")

            date = self.min_date
            col = 6
            while date <= self.max_date:
                standbyafter = 0
                workafter = 0

                if date in self.sumbyuser[email].keys():
                    standbyafter = self.sumbyuser[email][date].get(HourType.STANDBY, 0)
                    if standbyafter > 0:
                        worksheet.write_number(
                            row + 0, col, standbyafter, self.cellFormats["hourFormats"][HourFormat.STANDBY]
                        )
                        worksheet.write_number(
                            row + 2, col, standbyafter, self.cellFormats["hourFormats"][HourFormat.STANDBY]
                        )

                    workafter = self.sumbyuser[email][date].get(HourType.WORK, 0)
                    if workafter > 0:
                        worksheet.write_number(row + 3, col, workafter, self.getWorkCellFormat(date, workafter))
                        worksheet.write_number(row + 5, col, workafter, self.getWorkCellFormat(date, workafter))

                if date in self.sumstandbydec[email].keys():
                    standbyreduction = self.sumstandbydec[email][date]
                    worksheet.write_number(
                        row + 1, col, standbyreduction, self.cellFormats["hourFormats"][HourFormat.EMPTY]
                    )
                    worksheet.write_number(
                        row + 0,
                        col,
                        standbyafter - standbyreduction,
                        self.cellFormats["hourFormats"][HourFormat.STANDBY],
                    )

                if date in self.sumworkinc[email].keys():
                    workaddition = self.sumworkinc[email][date]
                    worksheet.write_number(
                        row + 4, col, workaddition, self.cellFormats["hourFormats"][HourFormat.EMPTY]
                    )
                    if workafter - workaddition >= 0:
                        worksheet.write_number(
                            row + 3,
                            col,
                            workafter - workaddition,
                            self.getWorkCellFormat(date, workafter - workaddition),
                        )

                date = date + td(days=1)
                col += 1

            for rd in range(0, 6):
                worksheet.write(row + rd, 0, email, self.cellFormats["datatxt"])
                worksheet.write(row + rd, 1, self.users[email])
                worksheet.write(row + rd, 2, self.approvers[email])

            s = sum([self.sumstandbydec[email][date] for date in self.sumstandbydec[email].keys()])
            worksheet.write_number(row + 1, 4, s, self.cellFormats["datanum"])

            w = sum([self.sumworkinc[email][date] for date in self.sumworkinc[email].keys()])
            worksheet.write_number(row + 4, 5, w, self.cellFormats["datanum"])

            row += 6

        worksheet.autofilter(1, 0, row - 6, 6 + (self.max_date - self.min_date).days)

    def generateSheet(self, workbook):
        if self.standbylimit:
            self.forceStandbyLimit()
        worksheet = workbook.add_worksheet("Standby changes")
        self.generateTitle(worksheet)
        self.generateHeader(worksheet)
        self.generateData(worksheet)
