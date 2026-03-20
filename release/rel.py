import sys
import os
import datetime
import shutil
from pathlib import Path

zipcmd = "7z" if sys.platform == "linux" else ".\\7za.exe"
nowstr = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
reldir = f"psg_{nowstr}"

os.system(f"{zipcmd} x -o{reldir} py.zip")
for fn in [
    "psg.py",
    "psg.bat",
    "patch.py",
    "patch.bat",
    "common.py",
    "config.py",
    "SGByUser.py",
    "SGByUserAndProject.py",
    "SGStandby.py",
    "SGStandbyChanges.py",
    "SGStandbyLimiter.py",
    "SGInfo.py",
    "SheetGenerator.py",
    "patchfiles.txt",
]:
    shutil.copy(os.path.join("..", fn), reldir)

os.mkdir(os.path.join(reldir, "cfg"))

for fn in ["holidays.txt", "weekends.txt", "workingdays.txt"]:
    shutil.copy(os.path.join("..", "cfg", fn), os.path.join(reldir, "cfg"))

for fn in ["projects.txt", "users.txt", "hotlines.txt", "userdata.csv"]:
    Path(os.path.join(reldir, "cfg", fn)).touch()

os.system(f"{zipcmd} a -tzip -r {reldir}.zip {reldir}")
shutil.rmtree(reldir)
