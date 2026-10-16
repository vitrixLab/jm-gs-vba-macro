"""Gate 2 — rebuild CDJ monthly TOTAL rows and append JUL..DEC blocks.

TOTAL rows become multi-range SUMs over the column-C marker partition (decision
D2), including column T (Sundry Debit) and U, which were never summed and made
every TOTAL row's balance check X non-zero.

Row 515 is deliberately excluded from the MAY range (decision D5 / PENDING).

The script closes and releases its Excel instance, then reopens the workbook in
a FRESH instance and asserts the formulas actually landed on disk, because an
earlier run reported success while the on-disk file still held the old formulas.
"""
import io
import os
import sys
import time

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
LOG = os.path.join(ROOT, "gate2_work", "gate2_totals_log.txt")

NUMFMT = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

PARTITIONS = {
    111: ("JANUARY", [(15, 110), (277, 321)]),
    193: ("FEB", [(113, 192), (322, 407)]),
    408: ("MARCH", [(195, 276)]),
    516: ("APRIL", [(410, 514)]),
    612: ("MAY", [(518, 611)]),          # row 515 excluded (D5)
    674: ("JUNE", [(614, 673)]),
}

MONTHS = [("JUL", "JULY"), ("AUG", "AUGUST"), ("SEP", "SEPTEMBER"),
          ("OCT", "OCTOBER"), ("NOV", "NOVEMBER"), ("DEC", "DECEMBER")]
DATA_ROWS_PER_BLOCK = 20

SUM_COLS = list(range(6, 19)) + [20, 21]   # F..R, T, U


def col_letter(col):
    s = ""
    while col:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return s


def expected_formula(col, spans):
    letter = col_letter(col)
    if len(spans) == 1:
        a, b = spans[0]
        return "=SUM(%s%d:%s%d)" % (letter, a, letter, b)
    return "=SUM(%s)" % ",".join("%s%d:%s%d" % (letter, a, letter, b) for a, b in spans)


def expected_juldec_formula(col, first, last):
    letter = col_letter(col)
    return "=SUM(%s%d:%s%d)" % (letter, first, letter, last)


def main():
    lines = []

    def log(msg):
        print(msg)
        lines.append(msg)

    mtime_before = os.path.getmtime(CANDIDATE)

    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = 3
        xl.EnableEvents = False
        xl.ScreenUpdating = False

        log("Opening %s" % CANDIDATE)
        wb = xl.Workbooks.Open(CANDIDATE, 0, False)
        cdj = wb.Worksheets("CDJ")

        log("--- TOTAL ROWS ---")
        for tot, (label, spans) in sorted(PARTITIONS.items()):
            for col in SUM_COLS:
                cdj.Cells(tot, col).Formula = expected_formula(col, spans)
                cdj.Cells(tot, col).NumberFormat = NUMFMT
            cdj.Cells(tot, 22).Formula = "=F%d" % tot
            cdj.Cells(tot, 23).Formula = "=SUM(G%d:T%d)" % (tot, tot)
            cdj.Cells(tot, 24).Formula = "=V%d+W%d" % (tot, tot)
            for col in (22, 23, 24):
                cdj.Cells(tot, col).NumberFormat = NUMFMT
            log("  r%-4d %-9s spans=%s" % (tot, label, spans))

        log("--- JUL..DEC BLOCKS ---")
        cursor = 675
        created = []
        for abbr, full in MONTHS:
            marker_row = cursor
            first_data = cursor + 1
            last_data = cursor + DATA_ROWS_PER_BLOCK
            total_row = last_data + 1

            cdj.Cells(marker_row, 3).Value = abbr
            for r in range(first_data, last_data + 1):
                cdj.Cells(r, 22).Formula = "=F%d" % r
                cdj.Cells(r, 23).Formula = "=SUM(G%d:T%d)" % (r, r)
                cdj.Cells(r, 24).Formula = "=V%d+W%d" % (r, r)
                for col in (22, 23, 24):
                    cdj.Cells(r, col).NumberFormat = NUMFMT

            cdj.Cells(total_row, 3).Value = full + " TOTAL"
            for col in SUM_COLS:
                cdj.Cells(total_row, col).Formula = expected_juldec_formula(col, first_data, last_data)
                cdj.Cells(total_row, col).NumberFormat = NUMFMT
            cdj.Cells(total_row, 22).Formula = "=F%d" % total_row
            cdj.Cells(total_row, 23).Formula = "=SUM(G%d:T%d)" % (total_row, total_row)
            cdj.Cells(total_row, 24).Formula = "=V%d+W%d" % (total_row, total_row)
            for col in (22, 23, 24):
                cdj.Cells(total_row, col).NumberFormat = NUMFMT

            created.append((abbr, marker_row, first_data, last_data, total_row))
            log("  %-3s marker r%-4d data r%d..%-4d TOTAL r%-4d" %
                (abbr, marker_row, first_data, last_data, total_row))
            cursor = total_row + 1

        xl.CalculateFull()

        log("--- IN-MEMORY VALUES ---")
        for tot, (label, spans) in sorted(PARTITIONS.items()):
            f = cdj.Cells(tot, 6).Value
            t = cdj.Cells(tot, 20).Value
            x = cdj.Cells(tot, 24).Value
            log("  r%-4d %-9s F=%s T=%s X=%s" % (
                tot, label,
                round(f, 2) if isinstance(f, (int, float)) else f,
                round(t, 2) if isinstance(t, (int, float)) else t,
                round(x, 6) if isinstance(x, (int, float)) else x))

        wb.Save()
        log("saved")
        wb.Close(True)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        del xl

    time.sleep(3)

    mtime_after = os.path.getmtime(CANDIDATE)
    log("--- DISK ---")
    log("  mtime before = %s" % time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime_before)))
    log("  mtime after  = %s" % time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime_after)))

    # ---- reopen in a FRESH instance and assert the formulas landed --------
    failures = []
    xl2 = win32com.client.DispatchEx("Excel.Application")
    try:
        xl2.Visible = False
        xl2.DisplayAlerts = False
        xl2.AutomationSecurity = 3
        xl2.EnableEvents = False
        xl2.ScreenUpdating = False
        wb2 = xl2.Workbooks.Open(CANDIDATE, 0, True)
        cdj2 = wb2.Worksheets("CDJ")
        xl2.CalculateFull()

        log("--- PERSISTED TOTAL ROWS (fresh instance) ---")
        for tot, (label, spans) in sorted(PARTITIONS.items()):
            for col in SUM_COLS:
                got = cdj2.Cells(tot, col).Formula
                want = expected_formula(col, spans)
                if str(got).upper() != want.upper():
                    failures.append("r%d %s: got %r want %r" % (tot, col_letter(col), got, want))
            for col, want in ((22, "=F%d" % tot),
                              (23, "=SUM(G%d:T%d)" % (tot, tot)),
                              (24, "=V%d+W%d" % (tot, tot))):
                got = cdj2.Cells(tot, col).Formula
                if str(got).upper() != want.upper():
                    failures.append("r%d %s: got %r want %r" % (tot, col_letter(col), got, want))
            f = cdj2.Cells(tot, 6).Value
            t = cdj2.Cells(tot, 20).Value
            x = cdj2.Cells(tot, 24).Value
            log("  r%-4d %-9s F=%s T=%s X=%s" % (
                tot, label,
                round(f, 2) if isinstance(f, (int, float)) else f,
                round(t, 2) if isinstance(t, (int, float)) else t,
                round(x, 6) if isinstance(x, (int, float)) else x))

        log("--- PERSISTED JUL..DEC (fresh instance) ---")
        cursor = 675
        for abbr, full in MONTHS:
            first_data = cursor + 1
            last_data = cursor + DATA_ROWS_PER_BLOCK
            total_row = last_data + 1
            marker = cdj2.Cells(cursor, 3).Value
            tot_label = cdj2.Cells(total_row, 3).Value
            if str(marker).strip().upper() != abbr:
                failures.append("marker r%d: got %r want %r" % (cursor, marker, abbr))
            if str(tot_label).strip().upper() != (full + " TOTAL").upper():
                failures.append("total label r%d: got %r want %r" % (total_row, tot_label, full + " TOTAL"))
            for col in SUM_COLS:
                got = cdj2.Cells(total_row, col).Formula
                want = expected_juldec_formula(col, first_data, last_data)
                if str(got).upper() != want.upper():
                    failures.append("r%d %s: got %r want %r" % (total_row, col_letter(col), got, want))
            log("  %-3s marker r%-4d (%s) TOTAL r%-4d (%s) F111-style formula r%d=%s"
                % (abbr, cursor, marker, total_row, tot_label, total_row,
                   cdj2.Cells(total_row, 6).Formula))
            cursor = total_row + 1

        wb2.Close(False)
    finally:
        try:
            xl2.Quit()
        except Exception:
            pass
        del xl2

    log("--- RESULT ---")
    if failures:
        log("  PERSISTENCE FAILURES (%d):" % len(failures))
        for f in failures[:40]:
            log("    " + f)
        status = 1
    else:
        log("  all formulas persisted correctly")
        status = 0

    with io.open(LOG, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print("\nlog -> %s" % LOG)
    return status


if __name__ == "__main__":
    sys.exit(main())
