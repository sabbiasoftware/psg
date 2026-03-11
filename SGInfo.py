from datetime import datetime as dt
from common import format_date, format_datetime
from SGStandbyLimiter import SGStandbyLimiter


class SGInfo(SGStandbyLimiter):
    def generateData(self, worksheet):
        duration = f"{format_date(self.min_date)} - {format_date(self.max_date)}"
        projects = "no filter" if len(self.config.Projects) == 0 else ", ".join(self.config.Projects)
        users = "no filter" if len(self.config.Users) == 0 else f"filtered to {len(self.config.Users)} users"
        generated = f"{format_datetime(dt.now())}"
        worksheet.write(0, 0, f"Duration: {duration}")
        worksheet.write(1, 0, f"Projects: {projects}")
        worksheet.write(2, 0, f"Users: {users}")
        worksheet.write(3, 0, f"Generated: {generated}")

    def generateSheet(self, workbook):
        worksheet = workbook.add_worksheet("Info")
        self.generateData(worksheet)
