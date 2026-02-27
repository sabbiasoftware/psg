import os

BASEURI = "https://raw.githubusercontent.com/sabbiasoftware/psg/refs/heads/main/"
PATCHFILESNAME = "patchfiles.txt"


def remove_file(fn):
    print(f"Checking {fn}: ", end="")
    if os.path.exists(fn):
        try:
            os.remove(fn)
            print("successfully removed")
        except Exception:
            print("failed (could not remove)")
    else:
        print("does not exist, no action needed")


def patch_file(fn):
    print(f"Checking {fn}: ", end="")
    fn_new = f"{fn}.new"
    cmd = f"curl -f -s -o {fn_new} {BASEURI}{fn}"
    # print(cmd)
    if os.system(cmd) != 0:
        print("failed (could not download latest version)")
        return

    content = ""
    if os.path.exists(fn):
        try:
            with open(fn, "r") as f:
                content = f.read()
        except Exception:
            print("failed (could not open current version)")
            return

    content_new = ""
    try:
        with open(fn_new, "r") as f:
            content_new = f.read()
    except Exception:
        print("failed (could not open latest verion)")
        return

    if content == content_new:
        print("no patch available")
        try:
            os.remove(fn_new)
        except Exception:
            print("error (could not delete downloaded version)")
    else:
        if os.path.exists(fn):
            try:
                os.remove(fn)
            except Exception:
                print("failed (could not delete current version)")
                return
        try:
            os.rename(fn_new, fn)
        except Exception:
            print("failed (could not move new version in place)")
            return
        print("successfully patched")


patch_file(PATCHFILESNAME)

with open(PATCHFILESNAME, "r") as f:
    for pfn in f:
        if pfn.startswith("-"):
            remove_file(pfn[1:])
        elif pfn.startswith("+"):
            patch_file(pfn[1:])
        else:
            patch_file(pfn.strip())
