from datetime import datetime as dt, timedelta as td
from decimal import Decimal as dec
from SGStandbyLimiter import SGStandbyLimiter
from common import HourType, HourFormat, dec_to_number
from config import Config


class SGProjectMonthly(SGStandbyLimiter):
    def __init__(self, config: Config, cellFormats, managerFromConfig) -> None:
        super().__init__(config, cellFormats, False)
        self.managerFromConfig = managerFromConfig
        self.sumprojectmonthly = {}

    def loadRow(self, row):
        super().loadRow(row)

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
        if (email, project, activity) not in self.sumprojectmonthly.keys():
            self.sumprojectmonthly[email, project, activity] = {}
        if date not in self.sumprojectmonthly[email, project, activity].keys():
            self.sumprojectmonthly[email, project, activity][date] = {}
        self.sumprojectmonthly[email, project, activity][date][t] = self.sumprojectmonthly[email, project, activity][
            date
        ].get(t, 0) + dec(row["Hours"])

    def _count_months(self):
        num = 0
        date = self.min_date
        lastmonth = None
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                num += 1
            date += td(days=1)
        return num

    def generateHeader(self, worksheet):
        self.generateCommonColumnHeaders(worksheet, 1, 0)
        self.generateColumnHeader(worksheet, 1, 3, "Status", self.cellFormats["headertxt"], 12)
        self.generateColumnHeader(worksheet, 1, 4, "Job Title", self.cellFormats["headertxt"], 40)
        self.generateColumnHeader(worksheet, 1, 5, "Grade", self.cellFormats["headertxt"], 12)
        self.generateColumnHeader(worksheet, 1, 6, "Rate", self.cellFormats["headernum"], 10)
        self.generateColumnHeader(worksheet, 1, 7, "Project", self.cellFormats["headertxt"], 48)
        self.generateColumnHeader(worksheet, 1, 8, "Activity", self.cellFormats["headertxt"], 12)
        for i, headerText in enumerate(["WorkH", "VacaD", "SickD"]):
            self.generateColumnHeader(worksheet, 1, 9 + i, headerText, self.cellFormats["headernum"], 8)

        num_months = self._count_months()

        worksheet.write(0, 12, "Hours", self.cellFormats["headertxt"])
        col = 12
        date = self.min_date
        lastmonth = None
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                self.generateColumnHeader(worksheet, 1, col, date.strftime("%Y-%m"), self.cellFormats["headertxt"], 10)
                col += 1
            date = date + td(days=1)

        worksheet.write(0, 12 + num_months, "Cost", self.cellFormats["headertxt"])
        col = 12 + num_months
        date = self.min_date
        lastmonth = None
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                self.generateColumnHeader(worksheet, 1, col, date.strftime("%Y-%m"), self.cellFormats["headertxt"], 10)
                col += 1
            date = date + td(days=1)

    def generateData(self, worksheet):
        row = 2
        col = 12
        num_months = self._count_months()
        for email, project, activity in sorted(self.sumprojectmonthly.keys()):
            if self.get_hour_type(project, activity) != HourType.STANDBY:
                total_hours = {}
                for ht in HourType:
                    total_hours[ht] = sum(
                        [
                            self.sumprojectmonthly[email, project, activity][date].get(ht, 0)
                            for date in self.sumprojectmonthly[email, project, activity]
                        ]
                    )

                if len(self.config.Projects) > 0 and total_hours[HourType.WORK] == 0:
                    continue

                manager = self.approvers[email]
                if self.managerFromConfig:
                    if email in self.config.UserData.keys():
                        manager = self.config.UserData[email]["Reporting to"]

                grade = ""
                if email in self.config.UserData.keys():
                    grade = self.config.UserData[email].get("Global Grade", "")
                    worksheet.write(row, 3, self.config.UserData[email].get("Employment Status", ""))
                    worksheet.write(row, 4, self.config.UserData[email].get("Job Title", ""))
                    worksheet.write(row, 5, grade)
                rate_str = self.config.Rates.get(grade, "")
                rate_val = dec(rate_str) if rate_str else None

                worksheet.write(row, 0, email, self.cellFormats["datatxt"])
                worksheet.write(row, 1, self.users[email])
                worksheet.write(row, 2, manager)
                worksheet.write(row, 6, rate_str, self.cellFormats["datausd"])
                worksheet.write(row, 7, project)
                worksheet.write(row, 8, activity)
                worksheet.write_number(row, 9, int(total_hours[HourType.WORK]), self.cellFormats["datanum"])
                worksheet.write_number(
                    row, 10, int(total_hours[HourType.VACATION] // dec(8)), self.cellFormats["datanum"]
                )
                worksheet.write_number(row, 11, int(total_hours[HourType.SICK] // dec(8)), self.cellFormats["datanum"])

                col = 12
                date = self.min_date
                lastmonth = None
                while date <= self.max_date:
                    if lastmonth != date.month:
                        lastmonth = date.month
                        month_hours = {}
                        for d in self.sumprojectmonthly[email, project, activity]:
                            if d.year == date.year and d.month == date.month:
                                for ht, hrs in self.sumprojectmonthly[email, project, activity][d].items():
                                    month_hours[ht] = month_hours.get(ht, 0) + hrs

                        v = self.get_only_hours(HourType.VACATION, month_hours)
                        if v is not None:
                            value = dec_to_number(v) if isinstance(v, dec) else v
                            worksheet.write_number(
                                row, col, value, self.cellFormats["hourFormats"][HourFormat.VACATION]
                            )
                        elif self.get_only_hours(HourType.SICK, month_hours) is not None:
                            s = self.get_only_hours(HourType.SICK, month_hours)
                            value = dec_to_number(s) if isinstance(s, dec) else s
                            worksheet.write_number(row, col, value, self.cellFormats["hourFormats"][HourFormat.SICK])
                        elif self.get_active_hours(month_hours) == 0:
                            worksheet.write(row, col, "", self.cellFormats["hourFormats"][HourFormat.EMPTY])
                        else:
                            w = month_hours.get(HourType.WORK, 0)
                            worksheet.write_number(
                                row,
                                col,
                                dec_to_number(w) if isinstance(w, dec) else w,
                                self.cellFormats["hourFormats"][HourFormat.WORK],
                            )
                        col += 1
                    date = date + td(days=1)

                col = 12 + num_months
                date = self.min_date
                lastmonth = None
                while date <= self.max_date:
                    if lastmonth != date.month:
                        lastmonth = date.month
                        month_hours = {}
                        for d in self.sumprojectmonthly[email, project, activity]:
                            if d.year == date.year and d.month == date.month:
                                for ht, hrs in self.sumprojectmonthly[email, project, activity][d].items():
                                    month_hours[ht] = month_hours.get(ht, 0) + hrs
                        total_hours_for_cost = self.get_active_hours(month_hours)
                        if rate_val is not None and total_hours_for_cost > 0:
                            cost = total_hours_for_cost * rate_val
                            worksheet.write_number(row, col, dec_to_number(cost), self.cellFormats["datausd"])
                        else:
                            worksheet.write(row, col, "", self.cellFormats["datausd"])
                        col += 1
                    date = date + td(days=1)

                row += 1

        worksheet.autofilter(1, 0, row - 1, col - 1)

    def generateSheet(self, workbook):
        worksheet = workbook.add_worksheet("Project monthly")
        self.generateHeader(worksheet)
        self.generateData(worksheet)
