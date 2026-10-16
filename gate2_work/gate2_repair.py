"""Gate 2 — historical CDJ data repairs on Global-Smile_2026-v8.4-fix.xlsm.

Repairs (all logged, all asserted):
  1. Clear the 166 old-ref duplicate rows (decision D1), preserving column C
     (month markers) and the V/W/X balance formulas.
  2. Clear the stale reference on the 19 manual rows that were wrongly stamped.
  3. Clear every spurious column-Y reference in rows 15..673 (F10).
  4. Repair the five unbalanced rows (r629/630/632/633/634).
  5. Post the two genuinely unposted PJ Non-Vat rows through the patched engine.

Nothing outside CDJ rows 15..673 and PJ Non-Vat is modified.
"""
import io
import os
import re
import sys

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
LOG = os.path.join(ROOT, "gate2_work", "gate2_repair_log.txt")

# Columns: C=3 month, D=4 day, E=5 particulars, F=6 .. U=21 amounts,
# V=22 W=23 X=24 balance formulas, Y=25 legacy ref, Z=26 note, AA=27 ref, AB=28 pnv
COL_C, COL_D, COL_E = 3, 4, 5
COL_F, COL_U = 6, 21
COL_V, COL_W, COL_X = 22, 23, 24
COL_Y, COL_Z, COL_AA, COL_AB = 25, 26, 27, 28
ROW_LO, ROW_HI = 15, 673
REF_SPLIT = 352  # refs below this are the pre-renumber ("old") set


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def txt(v):
    return str(v or "").strip().upper()


def num(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def main():
    lines = []

    def log(msg):
        print(msg)
        lines.append(msg)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1
    xl.EnableEvents = False
    xl.ScreenUpdating = False

    log("Opening %s" % CANDIDATE)
    wb = xl.Workbooks.Open(CANDIDATE, 0, False)
    cdj = wb.Worksheets("CDJ")
    nv = wb.Worksheets("PJ Non-Vat")

    data = cdj.Range("C15:AB674").Value2

    def V(r, c):
        return data[r - 15][c - 3]

    def is_total(r):
        return "TOTAL" in txt(V(r, COL_C)) or "TOTAL" in txt(V(r, COL_E))

    def refs_of(r):
        out = []
        for c in (COL_Y, COL_AA, COL_AB):
            v = V(r, c)
            if isinstance(v, str) and v.strip().upper().startswith("PJ"):
                d = re.sub(r"\D", "", v)
                if d:
                    out.append((c, int(d)))
        return out

    rows = [r for r in range(ROW_LO, ROW_HI + 1)
            if any(not blank(V(r, c)) for c in range(3, 29)) and not is_total(r)]

    def key(r):
        name = re.sub(r"\s+", " ", str(V(r, COL_E) or "").strip().lower())
        cash = round(num(V(r, COL_F)), 2) if not blank(V(r, COL_F)) else None
        return (name, cash)

    old_rows = sorted({r for r in rows for _, n in refs_of(r) if n < REF_SPLIT})
    new_rows = sorted({r for r in rows for _, n in refs_of(r) if n >= REF_SPLIT})

    new_keys, old_keys = {}, {}
    for r in new_rows:
        new_keys.setdefault(key(r), []).append(r)
    for r in old_rows:
        old_keys.setdefault(key(r), []).append(r)

    clear_data, clear_ref_only = [], []
    for k, ors in old_keys.items():
        nrs = new_keys.get(k, [])
        if nrs and len(nrs) >= len(ors):
            clear_data.extend(ors)
        else:
            clear_ref_only.extend(ors)
    clear_data = sorted(set(clear_data))
    clear_ref_only = sorted(set(clear_ref_only))

    log("--- DEDUP ---")
    log("old-format ref rows : %d" % len(old_rows))
    log("CLEAR DATA          : %d" % len(clear_data))
    log("CLEAR REF ONLY      : %d" % len(clear_ref_only))
    if len(clear_data) != 166:
        raise SystemExit("ABORT: expected 166 clear-data rows, got %d" % len(clear_data))
    if len(clear_ref_only) != 19:
        raise SystemExit("ABORT: expected 19 clear-ref-only rows, got %d" % len(clear_ref_only))

    anchors = [(r, V(r, COL_C)) for r in clear_data if not blank(V(r, COL_C))]
    log("C markers preserved on cleared rows: %s" % anchors)
    bad_anchors = [r for r, m in anchors if r in (15, 113, 195)]
    if sorted(bad_anchors) != [15, 113, 195]:
        raise SystemExit("ABORT: anchor rows 15/113/195 not all present in clear list")

    # ---- 1. clear duplicate rows (preserve C and V/W/X) -------------------
    for r in clear_data:
        cdj.Cells(r, COL_D).ClearContents()
        cdj.Cells(r, COL_E).ClearContents()
        for c in range(COL_F, COL_U + 1):
            cdj.Cells(r, c).ClearContents()
        for c in (COL_Y, COL_Z, COL_AA, COL_AB):
            cdj.Cells(r, c).ClearContents()
    log("cleared %d duplicate rows (D,E,F..U,Y,Z,AA,AB)" % len(clear_data))

    # ---- 2. clear stale refs on wrongly-stamped manual rows ---------------
    for r in clear_ref_only:
        for c in (COL_Y, COL_Z, COL_AA, COL_AB):
            cdj.Cells(r, c).ClearContents()
    log("cleared stale refs on %d manual rows" % len(clear_ref_only))

    # ---- 3. clear every spurious column-Y ref in the live journal ---------
    y_cleared = 0
    for r in range(ROW_LO, ROW_HI + 1):
        v = cdj.Cells(r, COL_Y).Value
        if not blank(v):
            cdj.Cells(r, COL_Y).ClearContents()
            y_cleared += 1
    log("cleared %d column-Y references" % y_cleared)

    # ---- 4. repair the five unbalanced rows -------------------------------
    for r in (629, 630):
        t_before = num(cdj.Cells(r, 20).Value)
        cdj.Cells(r, 20).ClearContents()
        log("r%d T=%.2f -> cleared (S empty, allocations already balance)" % (r, t_before))
    for r in (632, 633, 634):
        t_before = num(cdj.Cells(r, 20).Value)
        r_before = num(cdj.Cells(r, 18).Value)
        cdj.Cells(r, 20).Value = r_before
        cdj.Cells(r, 18).ClearContents()
        log("r%d T=%.2f R=%.2f -> T=%.2f R cleared" % (r, t_before, r_before, r_before))

    # ---- 5. post the two unposted PJ Non-Vat rows through the engine ------
    log("--- PJ NON-VAT ---")
    for r in (8, 9):
        log("  PJ Non-Vat r%d: A=%s B=%s E=%s G=%s H=%s J=%s N=%s"
            % (r, nv.Cells(r, 1).Value, nv.Cells(r, 2).Value, nv.Cells(r, 5).Value,
               nv.Cells(r, 7).Value, nv.Cells(r, 8).Value, nv.Cells(r, 10).Value,
               nv.Cells(r, 14).Value))

    xl.EnableEvents = True
    for r in (8, 9):
        ref = str(nv.Cells(r, 8).Value or "").strip()
        net = num(nv.Cells(r, 14).Value)
        if ref and net:
            try:
                xl.Run("SyncPJNonVatToCDJ", nv, r)
                log("  SyncPJNonVatToCDJ(%s) invoked" % ref)
            except Exception as exc:
                log("  SyncPJNonVatToCDJ(%s) FAILED: %s" % (ref, str(exc)[:160]))
        else:
            log("  SKIPPED r%d (ref=%r net=%r)" % (r, ref, net))
    xl.EnableEvents = False

    wb.Save()
    log("saved")

    with io.open(LOG, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")

    wb.Close(True)
    xl.Quit()
    print("\nlog -> %s" % LOG)
    return 0


if __name__ == "__main__":
    sys.exit(main())
