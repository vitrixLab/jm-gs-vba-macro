"""Gate 2 step 1b - second read-only inspection pass (CDJ layout details).

Nothing in the workbook is modified: macros disabled, events off, read-only open,
close without save.
"""
import os

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
XLSM = os.path.join(HERE, "Global-Smile_2026-v8.4.xlsm")
OUT = os.path.join(HERE, "gate2_work", "inspect_baseline2.txt")

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


def fmt(v):
    if isinstance(v, float) and v == int(v):
        return "%d" % v
    return repr(v)


def main():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    try:
        xl.AutomationSecurity = 3
    except Exception:
        pass
    wb = None
    try:
        wb = xl.Workbooks.Open(XLSM, 0, True)
        cdj = wb.Sheets("CDJ")
        last_e = cdj.Cells(cdj.Rows.Count, 5).End(-4162).Row

        p("===== CDJ protection =====")
        p("ProtectContents=%s ProtectDrawingObjects=%s" % (cdj.ProtectContents, cdj.ProtectDrawingObjects))
        for addr in ("X15", "X630", "X755", "F15", "Z15", "AA15", "C15", "T630"):
            p("  %s Locked=%s Formula=%r Value=%r" % (
                addr, cdj.Range(addr).Locked, cdj.Range(addr).Formula, cdj.Range(addr).Value))

        p("\n===== CDJ V/W/X formula presence =====")
        has_v, has_w, has_x = [], [], []
        for r in range(15, last_e + 1):
            for col, lst in ((22, has_v), (23, has_w), (24, has_x)):
                f = cdj.Cells(r, col).Formula
                if isinstance(f, str) and f.startswith("="):
                    lst.append(r)
        for name, lst in (("V", has_v), ("W", has_w), ("X", has_x)):
            if lst:
                # find gaps
                gaps = []
                s = lst[0]
                prev = lst[0]
                for r in lst[1:]:
                    if r != prev + 1:
                        gaps.append((s, prev))
                        s = r
                    prev = r
                gaps.append((s, prev))
                p("  %s formula ranges: %s" % (name, gaps))
            else:
                p("  %s: none" % name)
        x_text = [(r, cdj.Cells(r, 24).Value) for r in range(15, last_e + 1)
                  if cdj.Cells(r, 24).Value not in (None, "")
                  and not (isinstance(cdj.Cells(r, 24).Formula, str) and cdj.Cells(r, 24).Formula.startswith("="))]
        p("  X cells holding literal text/values (notes?): %d %s" % (len(x_text), x_text[:10]))

        p("\n===== CDJ rows 674..760 any content in C..AB =====")
        for r in range(674, 761):
            vals = []
            for c in range(3, 29):
                v = cdj.Cells(r, c).Value
                f = cdj.Cells(r, c).Formula
                if v not in (None, ""):
                    tag = "%s=%s" % (colletter(c), fmt(v))
                    if isinstance(f, str) and f.startswith("="):
                        tag += "[%s]" % f
                    vals.append(tag)
            if vals:
                p("  r%d: %s" % (r, " | ".join(vals)))

        p("\n===== CDJ rows 5700..5821 content =====")
        for r in range(5700, 5822):
            vals = []
            for c in range(1, 31):
                v = cdj.Cells(r, c).Value
                if v not in (None, ""):
                    vals.append("%s=%s" % (colletter(c), fmt(v)))
            if vals:
                p("  r%d: %s" % (r, " | ".join(vals)))

        p("\n===== CDJ rows with H (WHT) <> 0 in data rows =====")
        cnt = 0
        for r in range(15, 675):
            h = cdj.Cells(r, 8).Value
            if isinstance(h, (int, float)) and abs(h) > 0.005:
                cnt += 1
                if cnt <= 30:
                    p("  r%d H=%s E=%r R=%s T=%s F=%s X=%s (Xformula=%r)" % (
                        r, fmt(h), cdj.Cells(r, 5).Value, fmt(cdj.Cells(r, 18).Value),
                        fmt(cdj.Cells(r, 20).Value), fmt(cdj.Cells(r, 6).Value),
                        fmt(cdj.Cells(r, 24).Value), cdj.Cells(r, 24).Formula))
        p("  total rows with H<>0 in 15..674: %d" % cnt)

        p("\n===== CDJ ref columns cross-check =====")
        y_map, aa_map, ab_map = {}, {}, {}
        for r in range(15, last_e + 1):
            for col, mp in ((25, y_map), (27, aa_map), (28, ab_map)):
                v = cdj.Cells(r, col).Value
                if v not in (None, ""):
                    mp.setdefault(str(v).strip(), []).append(r)
        p("  Y refs=%d  AA refs=%d  AB refs=%d" % (len(y_map), len(aa_map), len(ab_map)))
        dupall = {}
        for mp in (y_map, aa_map, ab_map):
            for k, rows in mp.items():
                dupall.setdefault(k, set()).update(rows)
        p("  unique refs across Y/AA/AB: %d" % len(dupall))
        multi = {k: v for k, v in dupall.items() if len(v) > 1}
        p("  refs present in >1 row: %d" % len(multi))
        for k, v in list(multi.items())[:20]:
            p("    %s -> rows %s" % (k, sorted(v)))
        y_only = sorted(set(y_map) - set(aa_map))
        p("  Y-only refs (%d): %s" % (len(y_only), y_only))
        aa_only = sorted(set(aa_map) - set(y_map))
        p("  AA-only count: %d ; sample: %s" % (len(aa_only), aa_only[:10]))
        both = sorted(set(aa_map) & set(y_map))
        p("  in both Y and AA (%d): %s" % (len(both), both[:20]))

        # rows with ref but no particulars / data
        ghost = []
        for r in range(15, last_e + 1):
            e = cdj.Cells(r, 5).Value
            f = cdj.Cells(r, 6).Value
            refs = [str(cdj.Cells(r, c).Value).strip() for c in (25, 27, 28)
                    if cdj.Cells(r, c).Value not in (None, "")]
            if refs and (e in (None, "")) and (f in (None, "")):
                ghost.append((r, refs))
        p("\n  ghost rows (ref present, no E/F): %d %s" % (len(ghost), ghost[:20]))

        # rows with data but no ref
        noref = []
        for r in range(15, last_e + 1):
            e = cdj.Cells(r, 5).Value
            f = cdj.Cells(r, 6).Value
            c = cdj.Cells(r, 3).Value
            refs = [str(cdj.Cells(r, cc).Value).strip() for cc in (25, 27, 28)
                    if cdj.Cells(r, cc).Value not in (None, "")]
            if (f not in (None, "")) and not refs:
                noref.append((r, c, e, f))
        p("  rows with cash but no ref: %d ; sample: %s" % (len(noref), noref[:15]))

        p("\n===== PJ month column B blanks =====")
        pj = wb.Sheets("PJ")
        last_pj = pj.Cells(pj.Rows.Count, 5).End(-4162).Row
        blanks = 0
        filled = 0
        for r in range(8, last_pj + 1):
            b = pj.Cells(r, 2).Value
            if b in (None, ""):
                blanks += 1
            else:
                filled += 1
        p("  PJ rows 8..%d : month filled=%d blank=%d" % (last_pj, filled, blanks))
        # show blank rows that otherwise look like data
        shown = 0
        for r in range(8, last_pj + 1):
            if pj.Cells(r, 2).Value in (None, "") and pj.Cells(r, 8).Value not in (None, ""):
                if shown < 8:
                    p("    r%d B=%r C=%r H=%r E=%r" % (r, pj.Cells(r, 2).Value, pj.Cells(r, 3).Value,
                                                       pj.Cells(r, 8).Value, pj.Cells(r, 5).Value))
                shown += 1
        p("  data rows with blank month: %d" % shown)
        # header block of PJ columns O..R
        p("  PJ r6/r7 cols 14..22: %s" % [(colletter(c), pj.Cells(6, c).Value, pj.Cells(7, c).Value)
                                           for c in range(14, 23)])

        p("\n===== PJ Non-Vat2 columns 1..25 header r6/r7 + sample r8 =====")
        pnv2 = wb.Sheets("PJ Non-Vat2")
        for c in range(1, 26):
            a, b = pnv2.Cells(6, c).Value, pnv2.Cells(7, c).Value
            if a not in (None, "") or b not in (None, ""):
                p("  %s: r6=%r r7=%r" % (colletter(c), a, b))
        p("  r8: %s" % [(colletter(c), pnv2.Cells(8, c).Value) for c in range(1, 26)
                         if pnv2.Cells(8, c).Value not in (None, "")])
        p("  r188: %s" % [(colletter(c), pnv2.Cells(188, c).Value) for c in range(1, 26)
                           if pnv2.Cells(188, c).Value not in (None, "")])
        p("  last row col E = %d" % pnv2.Cells(pnv2.Rows.Count, 5).End(-4162).Row)
        p("  PNV2 protect=%s" % pnv2.ProtectContents)

        p("\n===== GL block 16/17 rows 208..246 col F =====")
        gl = wb.Sheets("GL")
        for r in range(208, 247):
            v = gl.Cells(r, 6).Value
            if v not in (None, ""):
                p("  F%d=%r" % (r, v))

        p("\n===== CDJ total rows detail (111,193,408,516,612,674) =====")
        for r in (111, 193, 408, 516, 612, 674):
            p("  r%d C=%r E=%r" % (r, cdj.Cells(r, 3).Value, cdj.Cells(r, 5).Value))
            for c in (6, 7, 8, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24):
                f = cdj.Cells(r, c).Formula
                if isinstance(f, str) and f.startswith("="):
                    p("    %s: %s = %s" % (colletter(c), f, fmt(cdj.Cells(r, c).Value)))
                elif cdj.Cells(r, c).Value not in (None, ""):
                    p("    %s: literal %s" % (colletter(c), fmt(cdj.Cells(r, c).Value)))

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
