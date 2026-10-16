"""Gate 2 — standalone, verbose run of the T8 synthetic-posting test."""
import os
import shutil
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
TMP = os.path.join(os.environ.get("TEMP", "."), "opencode")
TOL = 0.01


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def norm(v):
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip().upper()


def main():
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    work = os.path.join(TMP, "gate2_t8solo.xlsm")
    if os.path.exists(work):
        os.remove(work)
    shutil.copy2(CANDIDATE, work)
    print("copy -> %s" % work)

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

        last_pj = 8
        for r in range(8, 4000):
            if not blank(pj.Cells(r, 8).Value):
                last_pj = r
        synth = last_pj + 1
        print("last_pj=%d synth=%d" % (last_pj, synth))

        ok_all = True
        for ref, desc, net, vat, gross, day in [
                ("PJ-99001", "Clinic Materials and Supplies", 1000.0, 120.0, 1120.0, 12),
                ("PJ-99002", "Representation", 2000.0, 240.0, 2240.0, 15)]:
            print()
            print("--- case %s (%s) ---" % (ref, desc))
            pj.Cells(synth, 2).Value = "JAN"
            pj.Cells(synth, 3).Value = day
            pj.Cells(synth, 5).Value = "SYNTHETIC SUPPLIER"
            pj.Cells(synth, 7).Value = desc
            pj.Cells(synth, 8).Value = ref
            pj.Cells(synth, 10).Value = net
            pj.Cells(synth, 13).Value = vat
            pj.Cells(synth, 14).Value = gross
            aa14_before = cdj.Cells(14, 27).Value
            print("  AA14 before = %r" % aa14_before)
            try:
                xl.Run("SyncPJToCDJ", pj, synth)
                print("  SyncPJToCDJ returned normally")
            except Exception as exc:
                print("  SyncPJToCDJ RAISED: %s" % str(exc)[:300])
                ok_all = False
            xl.CalculateFull()
            aa14_after = cdj.Cells(14, 27).Value
            print("  AA14 after  = %r" % aa14_after)

            hit = None
            for r in range(15, 810):
                v = cdj.Cells(r, 27).Value
                if isinstance(v, str) and v.strip().upper() == ref:
                    hit = r
                    break
            print("  search 15..810 col27 -> %r" % hit)
            if hit is None:
                for c in (25, 27, 28):
                    for r in range(1, 7000):
                        v = cdj.Cells(r, c).Value
                        if isinstance(v, str) and v.strip().upper() == ref:
                            print("  but FOUND at col %d row %d" % (c, r))
                            hit = r
                            break
                    if hit:
                        break
            if hit is None:
                print("  NOT CREATED - failing this case")
                ok_all = False
                continue

            d = cdj.Cells(hit, 4).Value
            cash = cdj.Cells(hit, 6).Value
            wht = cdj.Cells(hit, 8).Value
            n14 = cdj.Cells(hit, 14).Value
            s = cdj.Cells(hit, 19).Value
            t = cdj.Cells(hit, 20).Value
            rcol = cdj.Cells(hit, 18).Value
            x = cdj.Cells(hit, 24).Value
            print("  r%d day=%r cash=%r wht=%r N=%r S=%r T=%r R=%r X=%r"
                  % (hit, d, cash, wht, n14, s, t, rcol, x))
            if not isnum(x) or abs(x) > TOL:
                print("  FAIL balance"); ok_all = False
            if not isnum(d) or int(d) != day:
                print("  FAIL day"); ok_all = False
            fixed = desc.lower().startswith("clinic")
            if fixed:
                if not isnum(n14) or abs(n14 - net) > TOL:
                    print("  FAIL fixed net"); ok_all = False
                if not blank(t):
                    print("  FAIL T should be empty"); ok_all = False
            else:
                if norm(s) != norm(desc):
                    print("  FAIL sundry S"); ok_all = False
                if not isnum(t) or abs(t - net) > TOL:
                    print("  FAIL sundry T"); ok_all = False
                if not blank(rcol):
                    print("  FAIL sundry R"); ok_all = False

            pj.Cells(synth, 14).Value = gross + 500.0
            pj.Cells(synth, 10).Value = net + 500.0
            xl.Run("SyncPJToCDJ", pj, synth)
            xl.CalculateFull()
            cash2 = cdj.Cells(hit, 6).Value
            t2 = cdj.Cells(hit, 20).Value
            x2 = cdj.Cells(hit, 24).Value
            print("  UPDATE r%d cash=%r S=%r T=%r X=%r" % (hit, cash2, cdj.Cells(hit, 19).Value, t2, x2))
            if not isnum(x2) or abs(x2) > TOL:
                print("  FAIL update balance"); ok_all = False
            if fixed and not blank(t2):
                print("  FAIL update did not clear T"); ok_all = False
            if not fixed and (not isnum(t2) or abs(t2 - (net + 500.0)) > TOL):
                print("  FAIL update did not refresh T"); ok_all = False

            for c in range(3, 29):
                cdj.Cells(hit, c).ClearContents()
            cdj.Cells(hit, 22).Formula = "=F%d" % hit
            cdj.Cells(hit, 23).Formula = "=SUM(G%d:T%d)" % (hit, hit)
            cdj.Cells(hit, 24).Formula = "=V%d+W%d" % (hit, hit)
            for c in range(2, 15):
                pj.Cells(synth, c).ClearContents()

        print()
        print("T8 overall:", "PASS" if ok_all else "FAIL")
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
