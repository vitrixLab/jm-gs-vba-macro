"""Gate 2 — probe FindCDJFreeRow / CDJRowIsFree behaviour on a fresh copy.

Case 1 of the synthetic-posting test landed in the ghost zone (row 5765) while
case 2 landed correctly at row 16, so the free-row resolver is state dependent.
This walks the CDJ the same way FindCDJFreeRow does and shows why.
"""
import os
import shutil
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
TMP = os.path.join(os.environ.get("TEMP", "."), "opencode")

MONTHS = {1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
          7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"}
NAMES = {v: k for k, v in MONTHS.items()}
NAMES.update({"JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "JUNE": 6,
              "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10,
              "NOVEMBER": 11, "DECEMBER": 12})


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def norm(v):
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip().upper()


def main():
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    work = os.path.join(TMP, "gate2_free.xlsm")
    if os.path.exists(work):
        os.remove(work)
    shutil.copy2(CANDIDATE, work)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    try:
        wb = xl.Workbooks.Open(work, 0, False)
        cdj = wb.Worksheets("CDJ")

        print("=== FindCDJFreeRow BEFORE any explicit InitializeEngine ===")
        for tm in (0, 1, 2, 3):
            print("  FindCDJFreeRow(%d) = %r" % (tm, xl.Run("FindCDJFreeRow", tm)))

        print()
        print("=== manual walk of rows 15..30 (marker carry-forward + free test) ===")
        print("  CDJRowIsFree / marker resolution as FindCDJFreeRow sees it")
        cur = 0
        for r in range(15, 31):
            c = cdj.Cells(r, 3).Value
            m = NAMES.get(norm(c), 0)
            if m:
                cur = m
            e = cdj.Cells(r, 5).Value
            f = cdj.Cells(r, 6).Value
            aa = cdj.Cells(r, 27).Value
            ab = cdj.Cells(r, 28).Value
            y = cdj.Cells(r, 25).Value
            free = (blank(e) and blank(f) and blank(aa) and blank(ab) and blank(y))
            print("  r%-4d C=%-10r curMonth=%-2d free=%-5s  E=%-22r F=%r AA=%r Y=%r"
                  % (r, c, cur, free, str(e)[:22], f, aa, y))

        print()
        print("=== how far does the January marker reach? ===")
        cur = 0
        first_jan_free = None
        for r in range(15, 6700):
            c = cdj.Cells(r, 3).Value
            m = NAMES.get(norm(c), 0)
            if m:
                cur = m
            if cur != 1:
                continue
            e = cdj.Cells(r, 5).Value
            f = cdj.Cells(r, 6).Value
            aa = cdj.Cells(r, 27).Value
            ab = cdj.Cells(r, 28).Value
            y = cdj.Cells(r, 25).Value
            if blank(e) and blank(f) and blank(aa) and blank(ab) and blank(y):
                first_jan_free = r
                break
        print("  first free row whose carry-forward month is 1:", first_jan_free)

        print()
        print("=== every month marker in column C, with its row ===")
        markers = []
        for r in range(15, 6700):
            c = cdj.Cells(r, 3).Value
            if not blank(c) and NAMES.get(norm(c), 0):
                markers.append((r, str(c), NAMES[norm(c)]))
        print("  count =", len(markers))
        for r, c, m in markers:
            print("     r%-5d %-10s -> month %d" % (r, c, m))

        print()
        print("=== PJ column B: month markers that PJMonthOf would see ===")
        pj = wb.Worksheets("PJ")
        seen = []
        for r in range(1, 401):
            b = pj.Cells(r, 2).Value
            if not blank(b) and NAMES.get(norm(b), 0):
                seen.append((r, str(b), NAMES[norm(b)]))
        print("  count =", len(seen))
        for r, b, m in seen[:30]:
            print("     PJ r%-4d B=%-10s -> month %d" % (r, b, m))
        print("  PJMonthOf(row 400) would return:", seen[-1][2] if seen else 0)

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
