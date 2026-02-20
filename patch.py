import os

BASEURI = "https://raw.githubusercontent.com/sabbiasoftware/psg/refs/heads/main/"
PATCHFILESNAME = "patchfiles.txt"

def patch_file(fn):
    print(f"Checking {fn}: ", end="")
    fn_new = f"{fn}.new"
    cmd = f"curl -f -s -o {fn_new} {BASEURI}{fn}"
    # print(cmd)
    if os.system(cmd) != 0:
        print("failed, skipping")
        return

    content = ""
    with open(fn, "r") as f:
        content = f.read()

    content_new = ""
    with open(fn_new, "r") as f:
        content_new = f.read()

    if content == content_new:
        os.remove(fn_new)
        print("no patch available")
    else:
        os.remove(fn)
        os.rename(fn_new, fn)
        print("successfully patched")


patch_file(PATCHFILESNAME)

with open(PATCHFILESNAME, "r") as f:
    for pfn in f:
        patch_file(pfn.strip())
