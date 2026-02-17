import datetime
import os
import shutil

# zipcmd = "7z" if sys.platform == "linux" else "./7za"

nowstr = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
reldir = f"psg_{nowstr}"
os.chdir("release")
os.system(f"7z x -o{reldir} py.zip")
shutil.copy(os.path.join("..", "sum.py"), reldir)
shutil.copy(os.path.join("..", "psg.bat"), reldir)
shutil.copytree(
    os.path.join("..", "cfg"), os.path.join(reldir, "cfg"), dirs_exist_ok=True
)
os.system(f"7z a -tzip -r {reldir}.zip {reldir}")
shutil.rmtree(reldir)
