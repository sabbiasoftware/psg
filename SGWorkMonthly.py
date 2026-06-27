from datetime import datetime as dt, timedelta as td
from decimal import Decimal as dec
from SheetGenerator import SheetGenerator
from common import HourType, HourFormat, dec_to_number
from config import Config


class SGWorkMonthly(SheetGenerator):
    def __init__(self, config: Config, cellFormats, managerFromConfig) -> None:
        super().__init__(config, cellFormats)
        self.managerFromConfig = managerFromConfig
        self.sumworkmonthly = {}

    def loadRow(self, row):
        super().loadRow(row)

        date = None
        try:
            date = dt.strptime(row["Date"], "%Y%m%d")
        except ValueError:
            return

        email = row["Email Address"].lower()
        t = self.get_hour_type(row["Project"], row["Activity"])
        if t != HourType.WORK:
            return

        proj = row["Project"]
        desc = row["Project Description"]
        project = f"{proj} {desc}" if proj != desc else proj
        activity = row["Activity"]
        ym = (date.year, date.month)

        if email not in self.sumworkmonthly:
            self.sumworkmonthly[email] = {}
        if ym not in self.sumworkmonthly[email]:
            self.sumworkmonthly[email][ym] = {"projects": set(), "activities": set(), "work_hours": dec(0)}

        self.sumworkmonthly[email][ym]["projects"].add(project)
        self.sumworkmonthly[email][ym]["activities"].add(activity)
        self.sumworkmonthly[email][ym]["work_hours"] += dec(row["Hours"])

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
        self.generateColumnHeader(worksheet, 1, 9, "WorkH", self.cellFormats["headernum"], 8)

        num_months = self._count_months()

        worksheet.write(0, 10, "Hours", self.cellFormats["headertxt"])
        col = 10
        date = self.min_date
        lastmonth = None
        while date <= self.max_date:
            if lastmonth != date.month:
                lastmonth = date.month
                self.generateColumnHeader(worksheet, 1, col, date.strftime("%Y-%m"), self.cellFormats["headertxt"], 10)
                col += 1
            date = date + td(days=1)

        worksheet.write(0, 10 + num_months, "Cost", self.cellFormats["headertxt"])
        col = 10 + num_months
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
        col = 10
        num_months = self._count_months()

        for email in sorted(self.sumworkmonthly.keys()):
            all_projects = set()
            all_activities = set()
            total_work = dec(0)
            for ym_data in self.sumworkmonthly[email].values():
                all_projects.update(ym_data["projects"])
                all_activities.update(ym_data["activities"])
                total_work += ym_data["work_hours"]

            if len(self.config.Projects) > 0 and total_work == 0:
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
            worksheet.write(row, 6, rate_val, self.cellFormats["datausd"])
            worksheet.write(row, 7, ", ".join(sorted(all_projects)))
            worksheet.write(row, 8, ", ".join(sorted(all_activities)))
            worksheet.write_number(row, 9, int(total_work), self.cellFormats["datanum"])

            col = 10
            date = self.min_date
            lastmonth = None
            while date <= self.max_date:
                if lastmonth != date.month:
                    lastmonth = date.month
                    ym = (date.year, date.month)
                    ym_data = self.sumworkmonthly[email].get(ym)
                    if ym_data is not None and ym_data["work_hours"] > 0:
                        w = ym_data["work_hours"]
                        worksheet.write_number(
                            row,
                            col,
                            dec_to_number(w) if isinstance(w, dec) else w,
                            self.cellFormats["hourFormats"][HourFormat.WORK],
                        )
                    else:
                        worksheet.write(row, col, "", self.cellFormats["hourFormats"][HourFormat.EMPTY])
                    col += 1
                date = date + td(days=1)

            col = 10 + num_months
            date = self.min_date
            lastmonth = None
            while date <= self.max_date:
                if lastmonth != date.month:
                    lastmonth = date.month
                    ym = (date.year, date.month)
                    ym_data = self.sumworkmonthly[email].get(ym)
                    total_hours_for_cost = ym_data["work_hours"] if ym_data is not None else 0
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
        worksheet = workbook.add_worksheet("SGWorkMonthly")
        self.generateHeader(worksheet)
        self.generateData(worksheet)
