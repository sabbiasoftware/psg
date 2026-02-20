import os

BASEURI = "https://raw.githubusercontent.com/sabbiasoftware/psg/refs/heads/main/"


def patch_file(fn):
    print(f"Checking {fn}")
    fn_new = f"{fn}.new"
    cmd = f"curl -s -o {fn_new} {BASEURI}{fn}"
    print(cmd)
    if os.system(cmd) != 0:
        print(f"Check failed, skipping {fn}")
        return

    content = ""
    with open(fn, "r") as f:
        content = f.read()

    content_new = ""
    with open(fn_new, "r") as f:
        content_new = f.read()

    if content == content_new:
        os.remove(fn_new)
        print(f"No patch found for {fn}")
    else:
        os.remove(fn)
        os.rename(fn_new, fn)
        print(f"Successfully patched {fn}")


patch_file("psg.bat")
patch_file("psg.py")
