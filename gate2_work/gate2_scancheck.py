"""Gate 2 — run RefreshGL on a throwaway copy and report the scan gate result,
plus the distinct unmapped sundry titles that remain."""
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
    work = os.path.join(TMP, "gate2_scan.xlsm")
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
        gl = wb.Worksheets("GL")

        xl.Run("RefreshGL")
        xl.CalculateFull()

        aud = wb.Worksheets("GL_AUDIT")
        last = aud.UsedRange.Row + aud.UsedRange.Rows.Count - 1
        print("=== scan gate ===")
        print("  status :", aud.Cells(last, 2).Value)
        print("  summary:", aud.Cells(last, 3).Value)
        det = aud.Cells(last, 4).Value
        if det:
            parts = [p.strip() for p in str(det).splitlines() if p.strip()]
            print("  detail lines: %d" % len(parts))
            titles = {}
            for p in parts:
                if "->" in p and "unmapped" in p.lower():
                    t = p.split("->", 1)[1].rsplit("unmapped", 1)[0].strip()
                    titles[t] = titles.get(t, 0) + 1
            print("  distinct unmapped sundry titles remaining: %d" % len(titles))
            for t in sorted(titles):
                print("     %-42s x%d" % (t, titles[t]))

        print()
        print("=== alias check: is V8_GJMap resolving the variants now? ===")
        for probe in ["Clinicians Fee- Visiting", "Clinicians Fee - Visiting'",
                      "Visiting Clinicians Fee", "Clinicians Fee- Corporators",
                      "Clinicians Fee- Corporatos", "Clinacians Fee- Corporators",
                      "Clinician's Fee-Visiting", "Clinician's Fee-Corporators",
                      "Representation", "Partners Clinicians Fee",
                      "SSS,HDMF AND PHIC ER SHARE", "Incentives"]:
            try:
                got = xl.Run("V8_GJMap", probe)
            except Exception as exc:
                got = "ERR %s" % str(exc)[:60]
            print("   %-34r -> %r" % (probe, got))

        print()
        print("=== GL Cash in Bank block (rows 13..24) after scan ===")
        for r in range(13, 25):
            print("  r%-4d %-6s %-14s D=%-14s C=%-14s bal=%s" %
                  (r, gl.Cells(r, 2).Value, str(gl.Cells(r, 6).Value)[:14],
                   gl.Cells(r, 7).Value, gl.Cells(r, 8).Value, gl.Cells(r, 9).Value))

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
