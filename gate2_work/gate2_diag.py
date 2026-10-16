"""Gate 2 diagnostic: unmapped sundry titles, invalid scan cells, and the two
unposted-reference cases (PJ-00586, PNV-00004).

Reads each sheet as one bulk range instead of hammering COM cell-by-cell.
"""
import os
import shutil
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
TMP = os.path.join(os.environ.get("TEMP", "."), "opencode")

CDJ_LO, CDJ_HI = 15, 6700
COL_C, COL_D, COL_E, COL_F = 3, 4, 5, 6


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


def grid(ws, r1, r2, c1, c2):
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value2


def row_of(mat, r, c, c1):
    return mat[r - 1][c - c1]


def main():
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    work = os.path.join(TMP, "gate2_diag.xlsm")
    if os.path.exists(work):
        os.remove(work)
    shutil.copy2(CANDIDATE, work)

    xl = app()
    try:
        wb = xl.Workbooks.Open(work, 0, True)
        cdj = wb.Worksheets("CDJ")
        gl = wb.Worksheets("GL")
        pj = wb.Worksheets("PJ")
        nv = wb.Worksheets("PJ Non-Vat")

        gl_grid = grid(gl, 13, 700, 6, 8)
        gl_titles = set()
        for row in gl_grid:
            if row and not blank(row[0]):
                gl_titles.add(norm(row[0]))
        print("GL chart of accounts: %d distinct titles" % len(gl_titles))
        for t in sorted(gl_titles):
            print("   ", t)
        print()

        cdj_grid = grid(cdj, CDJ_LO, CDJ_HI, COL_C, 28)
        n = len(cdj_grid)
        print("CDJ bulk read: %d rows x %d cols" % (n, 28 - COL_C + 1))
        print()

        def C(r, c):
            return cdj_grid[r - CDJ_LO][c - COL_C]

        print("=== unmapped sundry titles in CDJ column S ===")
        unmapped, mapped = {}, {}
        for r in range(CDJ_LO, CDJ_HI + 1):
            s = C(r, 19)
            if blank(s) or norm(s) in ("TOTAL", "SUNDRY ACCOUNT"):
                continue
            (mapped if norm(s) in gl_titles else unmapped).setdefault(norm(s), []).append(r)
        print("  mapped   : %d distinct" % len(mapped))
        print("  UNMAPPED : %d distinct" % len(unmapped))
        tot = 0.0
        for k in sorted(unmapped):
            rows = unmapped[k]
            amt = 0.0
            for r in rows:
                if isnum(C(r, 20)):
                    amt += C(r, 20)
                if isnum(C(r, 18)):
                    amt += C(r, 18)
            tot += amt
            print("   %-42s n=%-3d first=r%-4d sum(T+R)=%12.2f" % (k, len(rows), rows[0], amt))
        print("  total amount on unmapped sundry rows: %.2f" % tot)
        print("  distinct unmapped titles:")
        for k in sorted(unmapped):
            print("     %s" % k)
        print()

        print("=== invalid (non-numeric) cells in the scanned range ===")
        last = CDJ_LO
        for r in range(CDJ_LO, CDJ_HI + 1):
            if not blank(C(r, COL_E)):
                last = r
        print("  last row with data in col E =", last)
        inv = []
        for r in range(CDJ_LO, last + 1):
            for c in list(range(6, 19)) + [20, 21]:
                v = C(r, c)
                if blank(v):
                    continue
                if not isnum(v):
                    inv.append("CDJ r%d c%d = %r" % (r, c, v))
        print("  non-numeric cells in F..R,T,U: %d" % len(inv))
        for s in inv[:60]:
            print("   ", s)
        if len(inv) > 60:
            print("    ... and %d more" % (len(inv) - 60))
        print()

        print("=== rows where col E is blank but amounts exist (scanner skips these) ===")
        skip = []
        for r in range(CDJ_LO, last + 1):
            if blank(C(r, COL_E)) and any(isnum(C(r, c)) for c in range(6, 22)):
                skip.append(r)
        print("  count=%d rows=%s" % (len(skip), skip[:40]))
        print()

        print("=== PJ-00586 ===")
        pj_grid = grid(pj, 8, 4000, 1, 14)
        for i, row in enumerate(pj_grid):
            if row and str(row[7] or "").strip().upper() == "PJ-00586":
                print("  PJ row %d: B=%s C=%s E=%s G=%s J=%s M=%s N=%s" %
                      (8 + i, row[1], row[2], row[4], row[6], row[9], row[12], row[13]))
        print()

        print("=== PJ Non-Vat rows 8..12 ===")
        nv_grid = grid(nv, 8, 12, 1, 14)
        for i, row in enumerate(nv_grid):
            print("  r%-3d A=%-10s B=%-6s C=%-6s E=%-28s G=%-10s H=%-12s N=%s" %
                  (8 + i, row[0], row[1], row[2], str(row[4])[:28], row[6], row[7], row[13]))
        print("  PNV refs found in CDJ column AB:")
        for r in range(CDJ_LO, CDJ_HI + 1):
            v = C(r, 28)
            if not blank(v):
                print("     CDJ r%-4d AB=%-12s E=%-28s C=%-8s F=%-12s N=%s" %
                      (r, v, str(C(r, 5))[:28], C(r, 3), C(r, 6), C(r, 14)))
        print()

        print("=== row 515 neighbourhood ===")
        for r in (513, 514, 515, 516, 517, 518):
            print("  r%-4d C=%-10s D=%-6s E=%-32s F=%-12s X=%s" %
                  (r, C(r, 3), C(r, 4), str(C(r, 5))[:32], C(r, 6), C(r, 24)))
        print()

        print("=== month partition roll-up (marker carry-forward) ===")
        MONTHS = {1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
                  7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"}
        NAMES = {v: k for k, v in MONTHS.items()}
        NAMES.update({"JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "JUNE": 6,
                      "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10,
                      "NOVEMBER": 11, "DECEMBER": 12})
        cur, cash = 0, {}
        for r in range(CDJ_LO, CDJ_HI + 1):
            m = NAMES.get(norm(C(r, 3)), 0)
            if m:
                cur = m
            if not cur:
                continue
            t = C(r, 5)
            if blank(t) or "TOTAL" in norm(t):
                continue
            v = C(r, 6)
            if isnum(v):
                cash[cur] = cash.get(cur, 0.0) + v
        for m in range(1, 13):
            print("   %s  %14.2f" % (MONTHS[m], cash.get(m, 0.0)))

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
