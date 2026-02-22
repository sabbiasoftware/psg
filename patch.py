import os

BASEURI = "https://raw.githubusercontent.com/sabbiasoftware/psg/refs/heads/main/"
PATCHFILESNAME = "patchfiles.txt"


def patch_file(fn):
    print(f"Checking {fn}: ", end="")
    fn_new = f"{fn}.new"
    cmd = f"curl -f -s -o {fn_new} {BASEURI}{fn}"
    # print(cmd)
    if os.system(cmd) != 0:
        print("download failed, skipping")
        return

    content = ""
    try:
        with open(fn, "r") as f:
            content = f.read()
    except Exception:
        print("opening original file failed, skipping")
        return

    content_new = ""
    try:
        with open(fn_new, "r") as f:
            content_new = f.read()
    except Exception:
        print("opening new file failed, skipping")
        return

    if content == content_new:
        print("no patch available")
        try:
            os.remove(fn_new)
        except Exception:
            print("failed to delete new file")
    else:
        print("patch available")
        try:
            os.remove(fn)
        except Exception:
            print("failed to delete original (might be in use?), skipping")
            return
        try:
            os.rename(fn_new, fn)
        except Exception:
            print("failed to move new file, but old file was already deleted, skipping")
            return
        print("successfully patched")


patch_file(PATCHFILESNAME)

with open(PATCHFILESNAME, "r") as f:
    for pfn in f:
        patch_file(pfn.strip())
