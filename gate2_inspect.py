"""Gate 2 step 1 - read-only structural inspection of Global-Smile_2026-v8.4.xlsm.

Macros are force-disabled and events are off; the workbook is opened read-only and
closed without saving. Output is written to gate2_work/inspect_baseline.txt.
"""
import os
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
XLSM = os.path.join(HERE, "Global-Smile_2026-v8.4.xlsm")
OUT = os.path.join(HERE, "gate2_work", "inspect_baseline.txt")

lines = []


def p(s=""):
    lines.append(str(s))
    print(s)


def colletter(c):
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    try:
        xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    except Exception as e:
        p("AutomationSecurity set failed: %s" % e)
    wb = None
    try:
        wb = xl.Workbooks.Open(XLSM, 0, True)  # UpdateLinks=0, ReadOnly=True
        p("Opened read-only: %s" % XLSM)

        p("\n===== SHEETS =====")
        for ws in wb.Sheets:
            p("  %-18s codename=%s index=%s" % (ws.Name, ws.CodeName, ws.Index))

        cdj = wb.Sheets("CDJ")
        p("\n===== CDJ HEADERS rows 12-15 (cols 1..32) =====")
        for c in range(1, 33):
            v13 = cdj.Cells(13, c).Value
            v14 = cdj.Cells(14, c).Value
            v12 = cdj.Cells(12, c).Value
            if v12 not in (None, "") or v13 not in (None, "") or v14 not in (None, ""):
                p("  col %2d %-3s | r12=%r r13=%r r14=%r" % (c, colletter(c), v12, v13, v14))

        last_e = cdj.Cells(cdj.Rows.Count, 5).End(-4162).Row  # xlUp
        p("\nCDJ last row col E (particulars): %d" % last_e)
        p("CDJ AA13 (col27 row13)=%r  AA14 (col27 row14)=%r" % (cdj.Cells(13, 27).Value, cdj.Cells(14, 27).Value))
        p("CDJ Z13 (col26 row13)=%r  Z14=%r" % (cdj.Cells(13, 26).Value, cdj.Cells(14, 26).Value))
        p("CDJ Y14 (col25 row14)=%r  AB14 (col28 row14)=%r" % (cdj.Cells(14, 25).Value, cdj.Cells(14, 28).Value))
        p("CDJ X14 (col24 row14)=%r" % (cdj.Cells(14, 24).Value,))

        p("\n===== CDJ column usage rows 15..%d =====" % last_e)
        for c in range(19, 32):
            count = 0
            samples = []
            for r in range(15, min(last_e, 6000) + 1):
                v = cdj.Cells(r, c).Value
                if v not in (None, ""):
                    count += 1
                    if len(samples) < 5:
                        samples.append((r, v))
            p("  col %2d %-3s non-empty=%4d samples=%s" % (c, colletter(c), count, samples))

        p("\n===== CDJ column C markers (month / TOTAL) rows 1..%d =====" % last_e)
        markers = []
        for r in range(1, last_e + 1):
            v = cdj.Cells(r, 3).Value
            if v not in (None, ""):
                s = str(v).strip().upper()
                if s in ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
                         "JANUARY", "FEBRUARY", "MARCH", "APRIL", "JUNE", "JULY", "AUGUST", "SEPTEMBER",
                         "OCTOBER", "NOVEMBER", "DECEMBER") or "TOTAL" in s or s == "":
                    markers.append((r, v))
        for m in markers:
            p("  C%d = %r" % m)

        p("\n===== CDJ TOTAL rows (col E contains TOTAL) =====")
        for r in range(1, last_e + 1):
            v = cdj.Cells(r, 5).Value
            if v not in (None, "") and "TOTAL" in str(v).upper():
                p("  E%d = %r  F=%r T=%r U=%r V=%r W=%r X=%r (X formula=%r)" % (
                    r, v, cdj.Cells(r, 6).Value, cdj.Cells(r, 20).Value, cdj.Cells(r, 21).Value,
                    cdj.Cells(r, 22).Value, cdj.Cells(r, 23).Value, cdj.Cells(r, 24).Value,
                    cdj.Cells(r, 24).Formula))

        p("\n===== CDJ sample data rows =====")
        for r in (15, 16, 17, 630, 632, 633, 634, 636, 638, 670, 672, 673, 674, 755, 5677, 5821):
            row = []
            for c in list(range(3, 9)) + list(range(14, 29)):
                v = cdj.Cells(r, c).Value
                f = cdj.Cells(r, c).Formula
                if v not in (None, "") or (isinstance(f, str) and f.startswith("=")):
                    row.append("%s=%r%s" % (colletter(c), v, ("[%s]" % f) if isinstance(f, str) and f.startswith("=") else ""))
            p("  row %d: %s" % (r, " | ".join(row)))

        p("\n===== CDJ rows 674..700 col C/E/F =====")
        for r in range(674, 701):
            p("  r%d C=%r E=%r F=%r" % (r, cdj.Cells(r, 3).Value, cdj.Cells(r, 5).Value, cdj.Cells(r, 6).Value))

        # PJ / PJ Non-Vat / PJ Non-Vat2 layouts
        for name in ("PJ", "PJ Non-Vat", "PJ Non-Vat2"):
            ws = wb.Sheets(name)
            p("\n===== %s (codename %s) header rows 1..9 cols 1..20 =====" % (name, ws.CodeName))
            for r in range(1, 10):
                vals = []
                for c in range(1, 21):
                    v = ws.Cells(r, c).Value
                    if v not in (None, ""):
                        vals.append("%s=%r" % (colletter(c), v))
                if vals:
                    p("  r%d: %s" % (r, " | ".join(vals)))
            last = ws.Cells(ws.Rows.Count, 5).End(-4162).Row
            p("  last row col E = %d" % last)
            p("  sample rows:")
            for r in range(min(8, last), min(last, 20) + 1):
                vals = []
                for c in range(1, 21):
                    v = ws.Cells(r, c).Value
                    if v not in (None, ""):
                        vals.append("%s=%r" % (colletter(c), v))
                if vals:
                    p("    r%d: %s" % (r, " | ".join(vals)))

        # SUPPLIERS DATA
        sup = wb.Sheets("SUPPLIERS DATA")
        last_sup = sup.Cells(sup.Rows.Count, 3).End(-4162).Row
        p("\n===== SUPPLIERS DATA (last TIN row %d) =====" % last_sup)
        tins = {}
        placeholders = []
        malformed = []
        for r in range(2, last_sup + 1):
            tin = sup.Cells(r, 3).Value
            nm = sup.Cells(r, 4).Value
            if tin in (None, ""):
                continue
            key = str(tin).strip()
            tins.setdefault(key, []).append((r, nm))
            if "NOT FOUND" in key.upper():
                placeholders.append((r, key, nm))
            if key.count("-") > 2 or "--" in key:
                malformed.append((r, key, nm))
        dups = {k: v for k, v in tins.items() if len(v) > 1}
        p("unique TIN keys: %d ; duplicate keys: %d" % (len(tins), len(dups)))
        for k, v in list(dups.items())[:12]:
            p("  DUP %r -> %s" % (k, v))
        p("placeholders: %d" % len(placeholders))
        for x in placeholders:
            p("  %s" % (x,))
        p("malformed: %d" % len(malformed))
        for x in malformed[:15]:
            p("  %s" % (x,))
        p("header row 2? C2=%r D2=%r" % (sup.Cells(2, 3).Value, sup.Cells(2, 4).Value))
        p("header row 1: C1=%r D1=%r" % (sup.Cells(1, 3).Value, sup.Cells(1, 4).Value))

        # GL structure
        gl = wb.Sheets("GL")
        p("\n===== GL sheet =====")
        p("F11=%r" % gl.Cells(11, 6).Value)
        titles = []
        for b in range(46):
            r = 13 + b * 13
            titles.append((r, gl.Cells(r, 6).Value))
        p("GL account titles (46 blocks):")
        for r, t in titles:
            p("  r%d %r" % (r, t))

        calc = wb.Sheets("GL_V8_CALC")
        p("\n===== GL_V8_CALC =====")
        p("A1:E1 = %r" % [calc.Cells(1, c).Value for c in range(1, 6)])
        p("last row col A = %d" % calc.Cells(calc.Rows.Count, 1).End(-4162).Row)
        for r in range(2, 6):
            p("  r%d = %r" % (r, [calc.Cells(r, c).Value for c in range(1, 6)]))

        aud = wb.Sheets("GL_AUDIT")
        p("\n===== GL_AUDIT last 8 rows =====")
        lastr = aud.Cells(aud.Rows.Count, 1).End(-4162).Row
        p("last row = %d" % lastr)
        for r in range(max(1, lastr - 7), lastr + 1):
            p("  r%d = %r | %r | %r" % (r, aud.Cells(r, 1).Value, aud.Cells(r, 2).Value, aud.Cells(r, 3).Value))

    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception as e:
            p("close error: %s" % e)
        xl.Quit()
        del xl
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    print("\nwrote %s" % OUT)


if __name__ == "__main__":
    main()
