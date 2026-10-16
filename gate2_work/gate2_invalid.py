"""Gate 2 — locate the 559 Invalid entries the scanner reports.

ScanCDJ, ScanCRJ and ScanGJ each increment `inv`. CDJ was already shown to
contribute 0, so this replicates the CRJ and GJ loops to count and sample them.
"""
import os
import shutil
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
TMP = os.path.join(os.environ.get("TEMP", "."), "opencode")


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def main():
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    work = os.path.join(TMP, "gate2_inv.xlsm")
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

        print("=== CRJ: rows the scanner walks, and why they are Invalid ===")
        crj = wb.Worksheets("CRJ")
        last = 10
        for r in range(10, 4000):
            if not blank(crj.Cells(r, 7).Value):
                last = r
        print("  last CRJ row with data in col G =", last)
        inv_crj = []
        for r in range(10, last + 1):
            inv_no = crj.Cells(r, 7).Value
            tot = crj.Cells(r, 6).Value
            if blank(inv_no):
                continue
            if isinstance(tot, str) and "TOTAL" in tot.upper():
                continue
            d = crj.Cells(r, 15).Value
            is_date = False
            try:
                if d is not None and d != "":
                    is_date = True
            except Exception:
                pass
            # VBA IsDate on a string; replicate the "cannot parse" case
            if not is_date:
                inv_crj.append((r, inv_no, d, "col O (15) not a usable date"))
        print("  rows failing the date check: %d" % len(inv_crj))
        for r, inv_no, d, why in inv_crj[:25]:
            print("     r%-5d G=%-16r O=%-24r %s" % (r, inv_no, d, why))
        if len(inv_crj) > 25:
            print("     ... and %d more" % (len(inv_crj) - 25))

        print()
        print("=== GJ: non-numeric debit/credit cells the scanner walks ===")
        gj = wb.Worksheets("GJ")
        lastg = 11
        for r in range(11, 4000):
            if not blank(gj.Cells(r, 5).Value):
                lastg = r
        print("  last GJ row with data in col E =", lastg)
        inv_gj = []
        for r in range(11, lastg + 1):
            e = gj.Cells(r, 5).Value
            if blank(e) or (isinstance(e, str) and "TOTAL" in e.upper()):
                continue
            for c in (6, 7):
                v = gj.Cells(r, c).Value
                if blank(v):
                    continue
                if not isnum(v):
                    inv_gj.append((r, c, e, v))
        print("  non-numeric debit/credit cells: %d" % len(inv_gj))
        for r, c, e, v in inv_gj[:25]:
            print("     r%-5d c%d=%-24r E=%-34r" % (r, c, v, str(e)[:34]))
        if len(inv_gj) > 25:
            print("     ... and %d more" % (len(inv_gj) - 25))

        print()
        print("TOTAL invalid from CRJ date check: %d" % len(inv_crj))
        print("TOTAL invalid from GJ numbers    : %d" % len(inv_gj))
        print("COMBINED                          : %d" % (len(inv_crj) + len(inv_gj)))

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
