"""Gate 2 — debug why SyncPJToCDJ did not create a row for the synthetic PJ entry.

The procedure's ErrorHandler is an empty `End Sub`, so any error is swallowed
silently. This drives each dependency directly to find which step fails.
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


def main():
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    work = os.path.join(TMP, "gate2_t8dbg.xlsm")
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
        pj = wb.Worksheets("PJ")
        cdj = wb.Worksheets("CDJ")

        def run(label, macro, *args):
            try:
                r = xl.Run(macro, *args)
                print("  %-34s -> %r" % (label, r))
                return r
            except Exception as exc:
                print("  %-34s -> ERROR %s" % (label, str(exc)[:150]))
                return None

        print("=== dependency probes ===")
        run("InitializeEngine('CDJ')", "InitializeEngine", "CDJ")
        run("FindExpenseColumn('Representation')", "FindExpenseColumn", "Representation")
        run("FindExpenseColumn('Clinic Materials...')", "FindExpenseColumn", "Clinic Materials and Supplies")
        try:
            r = xl.Run("CalculateWithholdingTax", "Clinic Materials and Supplies", 1000.0)
            print("  %-34s -> %r" % ("CalculateWithholdingTax", r))
        except Exception as exc:
            print("  %-34s -> ERROR %s" % ("CalculateWithholdingTax", str(exc)[:150]))
        run("FindCDJFreeRow(1)", "FindCDJFreeRow", 1)
        run("FindCDJFreeRow(0)", "FindCDJFreeRow", 0)
        try:
            r = xl.Run("FindCDJRow", "PJ-99001")
            print("  %-34s -> %r" % ("FindCDJRow(PJ-99001)", r))
        except Exception as exc:
            print("  %-34s -> ERROR %s" % ("FindCDJRow", str(exc)[:150]))

        print()
        print("=== prepare synthetic PJ row ===")
        last_pj = 8
        for r in range(8, 4000):
            if not blank(pj.Cells(r, 8).Value):
                last_pj = r
        synth = last_pj + 1
        print("  last PJ row with a ref =", last_pj, " writing at", synth)
        pj.Cells(synth, 2).Value = "JAN"
        pj.Cells(synth, 3).Value = 12
        pj.Cells(synth, 4).Value = "000-000-000-000"
        pj.Cells(synth, 5).Value = "SYNTHETIC SUPPLIER"
        pj.Cells(synth, 6).Value = "SYNTHETIC ADDRESS"
        pj.Cells(synth, 7).Value = "Clinic Materials and Supplies"
        pj.Cells(synth, 8).Value = "PJ-99001"
        pj.Cells(synth, 10).Value = 1000.0
        pj.Cells(synth, 13).Value = 120.0
        pj.Cells(synth, 14).Value = 1120.0
        for c in (2, 3, 4, 5, 6, 7, 8, 10, 13, 14):
            print("    col%-3d = %r" % (c, pj.Cells(synth, c).Value))

        print()
        print("=== AA14 (gNextRow slot) before ===")
        print("  AA14 =", cdj.Cells(14, 27).Value)

        print()
        print("=== call SyncPJToCDJ ===")
        try:
            xl.Run("SyncPJToCDJ", pj, synth)
            print("  call returned normally")
        except Exception as exc:
            print("  call raised:", str(exc)[:300])
        xl.CalculateFull()

        print()
        print("=== AA14 after ===")
        print("  AA14 =", cdj.Cells(14, 27).Value)

        print()
        print("=== search whole CDJ for PJ-99001 ===")
        found = False
        for c in (25, 27, 28):
            for r in range(1, 7000):
                v = cdj.Cells(r, c).Value
                if isinstance(v, str) and v.strip().upper() == "PJ-99001":
                    print("  FOUND at col %d row %d" % (c, r))
                    found = True
        if not found:
            print("  NOT FOUND anywhere in CDJ")

        print()
        print("=== probe row FindCDJFreeRow(1) reports, and dump it ===")
        fr = run("FindCDJFreeRow(1)", "FindCDJFreeRow", 1)
        if isinstance(fr, (int, float)):
            rr = int(fr)
            print("  row %d: C=%r D=%r E=%r F=%r AA=%r AB=%r Y=%r" %
                  (rr, cdj.Cells(rr, 3).Value, cdj.Cells(rr, 4).Value,
                   cdj.Cells(rr, 5).Value, cdj.Cells(rr, 6).Value,
                   cdj.Cells(rr, 27).Value, cdj.Cells(rr, 28).Value, cdj.Cells(rr, 25).Value))

        print()
        print("=== does the engine see the PJ row? LoadCDJReferences + status ===")
        run("LoadCDJReferences", "LoadCDJReferences")

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
