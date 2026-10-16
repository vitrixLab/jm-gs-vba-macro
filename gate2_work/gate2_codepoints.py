"""Inspect the exact CLINICIAN'S FEE titles in the GL chart and the raw sundry
spellings in CDJ, including codepoints, so the alias table matches byte-exactly."""
import os
import shutil
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
TMP = os.path.join(os.environ.get("TEMP", "."), "opencode")


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def norm(v):
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip().upper()


def main():
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    work = os.path.join(TMP, "gate2_cp.xlsm")
    if os.path.exists(work):
        os.remove(work)
    shutil.copy2(CANDIDATE, work)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(work, 0, True)
        gl = wb.Worksheets("GL")
        cdj = wb.Worksheets("CDJ")

        print("=== GL titles containing CLINICIAN (exact codepoints) ===")
        grid = gl.Range(gl.Cells(13, 6), gl.Cells(700, 6)).Value2
        for row in grid:
            v = row[0] if isinstance(row, tuple) else row
            if blank(v):
                continue
            s = str(v)
            if "CLINICIAN" in s.upper():
                print("  repr=%r" % s)
                print("  codepoints=%s" % [hex(ord(ch)) for ch in s])
                print()

        print("=== distinct CDJ sundry titles containing CLINICIAN ===")
        cdj_grid = cdj.Range(cdj.Cells(15, 19), cdj.Cells(6700, 19)).Value2
        seen = {}
        for i, row in enumerate(cdj_grid):
            v = row[0] if isinstance(row, tuple) else row
            if blank(v) or isinstance(v, (int, float)):
                continue
            s = str(v)
            if "CLINICIAN" in s.upper():
                seen.setdefault(s, []).append(15 + i)
        for s in sorted(seen):
            rows = seen[s]
            print("  repr=%-46r n=%-3d first=r%d" % (s, len(rows), rows[0]))
            print("     codepoints=%s" % [hex(ord(ch)) for ch in s])

        wb.Close(False)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        del xl
    try:
        os.remove(work)
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
