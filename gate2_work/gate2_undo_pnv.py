"""Gate 2 — undo the mis-posted PNV-00004 ghost row, then re-run an
alias-aware unmapped-sundry analysis.

PNV-00004 was posted to CDJ row 5677 (the ghost zone) because PJ Non-Vat row 9
has no date, so the engine could not resolve a month and fell through to
FindCDJFreeRow(0). That posting is removed and the row reported as PENDING
instead of inventing a month.
"""
import os
import shutil
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
TMP = os.path.join(os.environ.get("TEMP", "."), "opencode")

CDJ_LO, CDJ_HI = 15, 6700
MONTHS = {1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
          7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"}
NAMES = {v: k for k, v in MONTHS.items()}
NAMES.update({"JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "JUNE": 6,
              "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10,
              "NOVEMBER": 11, "DECEMBER": 12})

# V8_GJMap alias table, replicated exactly.
ALIASES = {
    "TRANSPORATION AND TRAVEL": "TRANSPORTATION AND TRAVEL",
    "ADVANCES TO EMPLOYEE": "ADVANCES TO EMPLOYEES",
    "REPAIRS AND MAINTENANCE": "REPAIR AND MAINTENANCE",
    "OUTPUT VAT PAYABLE": "VAT PAYABLE",
    "CLINIC SUPPLIES": "CLINIC MATERIAL AND SUPPLIES",
    "MEDICAL EQUIPMENT": "DENTAL EQUIPMENT",
    "COST OF REVENUE": "CLINIC MATERIAL AND SUPPLIES",
    "SUPPLIES": "CLINIC MATERIAL AND SUPPLIES",
    "BANK CHARGE": "MISCELLANEOUS",
    "CHARGES": "MISCELLANEOUS",
}


def app():
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    return xl


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def norm(v):
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip().upper()


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def main():
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    backup = os.path.join(TMP, "gate2_candidate_backup.xlsm")
    shutil.copy2(CANDIDATE, backup)
    print("backup -> %s" % backup)

    xl = app()
    try:
        xl.AutomationSecurity = 1
        wb = xl.Workbooks.Open(CANDIDATE, 0, False)
        cdj = wb.Worksheets("CDJ")
        gl = wb.Worksheets("GL")

        print("=== undo PNV-00004 ghost posting ===")
        target = None
        for r in range(CDJ_LO, CDJ_HI + 1):
            v = cdj.Cells(r, 28).Value
            if isinstance(v, str) and v.strip().upper() == "PNV-00004":
                target = r
                break
        if target is None:
            print("  PNV-00004 not found in CDJ (nothing to undo)")
        else:
            print("  found at CDJ row %d (E=%r C=%r F=%r)" %
                  (target, cdj.Cells(target, 5).Value, cdj.Cells(target, 3).Value,
                   cdj.Cells(target, 6).Value))
            for c in range(3, 29):
                cdj.Cells(target, c).ClearContents()
            cdj.Cells(target, 22).Formula = "=F%d" % target
            cdj.Cells(target, 23).Formula = "=SUM(G%d:T%d)" % (target, target)
            cdj.Cells(target, 24).Formula = "=V%d+W%d" % (target, target)
            print("  cleared D..AB, restored V/W/X formulas")
        wb.Save()
        print("  saved")

        # ---- alias-aware unmapped analysis -------------------------------
        gl_grid = gl.Range(gl.Cells(13, 6), gl.Cells(700, 6)).Value2
        gl_titles = set()
        for row in gl_grid:
            if isinstance(row, tuple):
                if not blank(row[0]):
                    gl_titles.add(norm(row[0]))
            elif not blank(row):
                gl_titles.add(norm(row))

        def gjmap(raw):
            k = norm(raw)
            if k in ALIASES:
                return ALIASES[k]
            if k in gl_titles:
                return k
            return ""

        cdj_grid = cdj.Range(cdj.Cells(CDJ_LO, 3), cdj.Cells(CDJ_HI, 28)).Value2

        def C(r, c):
            return cdj_grid[r - CDJ_LO][c - 3]

        print()
        print("=== alias-aware unmapped sundry analysis ===")
        groups, zero_rows = {}, []
        for r in range(CDJ_LO, CDJ_HI + 1):
            s = C(r, 19)
            if blank(s) or norm(s) in ("TOTAL", "SUNDRY ACCOUNT"):
                continue
            if isnum(s):
                zero_rows.append(r)
                key = "<NUMERIC S>"
            else:
                key = norm(s)
            amt = 0.0
            if isnum(C(r, 20)):
                amt += C(r, 20)
            if isnum(C(r, 18)):
                amt += C(r, 18)
            g = gjmap(s)
            groups.setdefault(key, {"rows": [], "amt": 0.0, "maps_to": g})
            groups[key]["rows"].append(r)
            groups[key]["amt"] += amt

        unmapped = {k: v for k, v in groups.items() if v["maps_to"] == ""}
        mapped = {k: v for k, v in groups.items() if v["maps_to"] != ""}
        print("  distinct sundry titles        : %d" % len(groups))
        print("  resolved via V8_GJMap / chart : %d" % len(mapped))
        print("  UNMAPPED (blocks the scan)    : %d" % len(unmapped))
        print()
        print("  --- resolved (no action needed) ---")
        for k in sorted(mapped):
            v = mapped[k]
            print("   %-42s -> %-34s n=%-3d sum=%12.2f" %
                  (k, v["maps_to"], len(v["rows"]), v["amt"]))
        print()
        print("  --- UNMAPPED (operator COA decision required) ---")
        tot = 0.0
        for k in sorted(unmapped):
            v = unmapped[k]
            tot += v["amt"]
            print("   %-42s n=%-3d first=r%-4d sum(T+R)=%12.2f" %
                  (k, len(v["rows"]), v["rows"][0], v["amt"]))
        print()
        print("  TOTAL amount blocked by unmapped sundry: %.2f" % tot)
        print()
        print("  rows where column S is numeric instead of a title: %d -> %s"
              % (len(zero_rows), zero_rows[:30]))
        print("    (these are the rows reporting S=0.0; the sundry account name is missing)")

        # ---- where the scan stands now ----------------------------------
        print()
        print("=== scan status after undo ===")
        xl.AutomationSecurity = 1
        try:
            xl.Run("RefreshGL")
        except Exception as exc:
            print("  RefreshGL error:", str(exc)[:200])
        xl.CalculateFull()
        aud = wb.Worksheets("GL_AUDIT")
        last = aud.UsedRange.Row + aud.UsedRange.Rows.Count - 1
        print("  status  :", aud.Cells(last, 2).Value)
        print("  summary :", aud.Cells(last, 3).Value)
        det = aud.Cells(last, 4).Value
        if det:
            print("  detail  :", str(det)[:1500])

        wb.Close(True)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        del xl
    return 0


if __name__ == "__main__":
    sys.exit(main())
