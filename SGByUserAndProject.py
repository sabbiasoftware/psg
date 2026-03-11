from datetime import datetime as dt, timedelta as td
from decimal import Decimal as dec
from SGStandbyLimiter import SGStandbyLimiter
from common import HourType, dec_to_number
from config import Config


class SGByUserAndProject(SGStandbyLimiter):
    def __init__(self, config: Config, cellFormats, managerFromConfig) -> None:
        super().__init__(config, cellFormats, False)
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
        self.generateCommonColumnHeaders(worksheet, 5, 0)
        self.generateColumnHeader(worksheet, 5, 3, "Status", self.cellFormats["headertxt"], 12)
        self.generateColumnHeader(worksheet, 5, 4, "Job Title", self.cellFormats["headertxt"], 40)
        self.generateColumnHeader(worksheet, 5, 5, "Grade", self.cellFormats["headertxt"], 12)
        self.generateColumnHeader(worksheet, 5, 6, "Project", self.cellFormats["headertxt"], 48)
        self.generateColumnHeader(worksheet, 5, 7, "Activity", self.cellFormats["headertxt"], 12)
        for i, headerText in enumerate(["WorkH", "VacaD", "SickD"]):
            self.generateColumnHeader(worksheet, 5, 8 + i, headerText, self.cellFormats["headernum"], 8)
        self.generateHeaderDays(worksheet, 5, 11)

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
        worksheet = workbook.add_worksheet("Project details")
        self.generateTitle(worksheet)
        self.generateHeader(worksheet)
        self.generateData(worksheet)
