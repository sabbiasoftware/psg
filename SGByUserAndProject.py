from datetime import datetime as dt, timedelta as td
import calendar
from decimal import Decimal as dec
from SheetGenerator import SheetGenerator
from common import HourType, dec_to_number
from config import Config


class SGByUserAndProject(SheetGenerator):
    def __init__(self, config: Config, cellFormats, managerFromConfig) -> None:
        super().__init__(config, cellFormats)
        self.managerFromConfig = managerFromConfig
        self.sumbyuserandproj = {}

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

        proj = row["Project"]
        desc = row["Project Description"]
        project = f"{proj} {desc}" if proj != desc else proj
        activity = row["Activity"]
        if (email, project, activity) not in self.sumbyuserandproj.keys():
            self.sumbyuserandproj[email, project, activity] = {}
        if date not in self.sumbyuserandproj[email, project, activity].keys():
            self.sumbyuserandproj[email, project, activity][date] = {}
        self.sumbyuserandproj[email, project, activity][date][t] = self.sumbyuserandproj[email, project, activity][
            date
        ].get(t, 0) + dec(row["Hours"])

    def generateHeader(self, worksheet):
        worksheet.write(5, 0, "Name", self.cellFormats["headertxt"])
        worksheet.write(5, 1, "User", self.cellFormats["headertxt"])
        worksheet.write(5, 2, "Manager", self.cellFormats["headertxt"])
        worksheet.write(5, 3, "Status", self.cellFormats["headertxt"])
        worksheet.write(5, 4, "Job Title", self.cellFormats["headertxt"])
        worksheet.write(5, 5, "Grade", self.cellFormats["headertxt"])
        worksheet.write(5, 6, "Project", self.cellFormats["headertxt"])
        worksheet.write(5, 7, "Activity", self.cellFormats["headertxt"])
        worksheet.write(5, 8, "WorkH", self.cellFormats["headernum"])
        worksheet.write(5, 9, "VacaD", self.cellFormats["headernum"])
        worksheet.write(5, 10, "SickD", self.cellFormats["headernum"])

        date = self.min_date
        lastmonth = 0
        col = 11
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                worksheet.write(4, col, calendar.month_name[date.month], self.cellFormats["headertxt"])
            cf = (
                self.cellFormats["headerworkday"] if self.is_working_day(date) else self.cellFormats["headernonworkday"]
            )
            worksheet.write_number(5, col, date.day, cf)
            date = date + td(days=1)
            col += 1

        worksheet.set_column(0, 0, width=40)
        worksheet.set_column(1, 2, width=24)
        worksheet.set_column(3, 3, width=12)
        worksheet.set_column(4, 4, width=40)
        worksheet.set_column(5, 5, width=12)
        worksheet.set_column(6, 6, width=48)
        worksheet.set_column(7, 7, width=12)
        worksheet.set_column(8, 10, width=8)
        worksheet.set_column(11, col - 1, width=6)

    def generateData(self, worksheet):
        row = 6
        col = 11
        for email, project, activity in sorted(self.sumbyuserandproj.keys()):
            if self.get_hour_type(project, activity) != HourType.STANDBY:
                total_hours = {}
                for ht in HourType:
                    total_hours[ht] = sum(
                        [
                            self.sumbyuserandproj[email, project, activity][date].get(ht, 0)
                            for date in self.sumbyuserandproj[email, project, activity]
                        ]
                    )

                # if there are filter projects configured, then filter out people with 0 hours against projects
                if len(self.config.Projects) > 0 and total_hours[HourType.WORK] == 0:
                    continue

                date = self.min_date
                col = 11
                while date <= self.max_date:
                    hours = {}
                    if date in self.sumbyuserandproj[email, project, activity].keys():
                        hours = self.sumbyuserandproj[email, project, activity][date]

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
                if email in self.config.UserData.keys():
                    worksheet.write(row, 3, self.config.UserData[email].get("Employment Status", ""))
                    worksheet.write(row, 4, self.config.UserData[email].get("Job Title", ""))
                    worksheet.write(row, 5, self.config.UserData[email].get("Global Grade", ""))
                worksheet.write(row, 6, project)
                worksheet.write(row, 7, activity)
                worksheet.write_number(row, 8, int(total_hours[HourType.WORK]), self.cellFormats["datanum"])
                worksheet.write_number(
                    row, 9, int(total_hours[HourType.VACATION] // dec(8)), self.cellFormats["datanum"]
                )
                worksheet.write_number(row, 10, int(total_hours[HourType.SICK] // dec(8)), self.cellFormats["datanum"])

                row += 1

        worksheet.autofilter(5, 0, row - 1, col - 1)

    def generateSheet(self, workbook):
        worksheet = workbook.add_worksheet("By user and project")
        self.generateTitle(worksheet)
        self.generateHeader(worksheet)
        self.generateData(worksheet)
