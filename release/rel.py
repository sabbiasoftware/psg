import sys
import os
import datetime
import shutil

zipcmd = "7z" if sys.platform == "linux" else ".\\7za.exe"
nowstr = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
reldir = f"psg_{nowstr}"

os.system(f"{zipcmd} x -o{reldir} py.zip")
for fn in ["psg.py", "psg.bat", "patch.py", "patch.bat", "patchfiles.txt"]:
    shutil.copy(os.path.join("..", fn), reldir)
shutil.copytree(os.path.join("..", "cfg"), os.path.join(reldir, "cfg"), dirs_exist_ok=True)
os.system(f"{zipcmd} a -tzip -r {reldir}.zip {reldir}")
shutil.rmtree(reldir)
