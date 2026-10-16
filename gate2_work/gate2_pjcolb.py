"""Gate 2 — what does PJMonthOf see in PJ column B?"""
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
    work = os.path.join(TMP, "gate2_pjcolb.xlsm")
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
        pj = wb.Worksheets("PJ")
        raw = pj.Range(pj.Cells(1, 2), pj.Cells(450, 2)).Value2
        print("=== PJ column B, rows 1..450: distinct values ===")
        vals = {}
        for i, row in enumerate(raw):
            v = row[0] if isinstance(row, tuple) else row
            if blank(v):
                continue
            key = repr(v)
            vals.setdefault(key, []).append(i + 1)
        print("  distinct non-empty values: %d" % len(vals))
        for k in sorted(vals)[:40]:
            print("    %-28s rows=%s" % (k, vals[k][:8]))

        print()
        print("=== which of those are month markers CDJMarkerMonth would match? ===")
        hits = []
        for i, row in enumerate(raw):
            v = row[0] if isinstance(row, tuple) else row
            if blank(v):
                continue
            m = NAMES.get(norm(v), 0)
            if m:
                hits.append((i + 1, str(v), m))
        print("  count =", len(hits))
        for r, v, m in hits:
            print("    PJ r%-4d B=%-12r -> month %d" % (r, v, m))
        print("  PJMonthOf would return (last hit):", hits[-1][2] if hits else 0)

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
