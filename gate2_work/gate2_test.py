"""Gate 2 — verification suite for Global-Smile_2026-v8.4-fix.xlsm.

T1  module compile / callable surface          (macros must be enabled)
T2  RefreshGL scan gate on baseline vs patched
T3  every CDJ data row balances (X ~ 0)
T4  TOTAL rows equal the recomputed marker-partition sums
T5  CDJ partition cash vs scanned GL Cash in Bank
T6  PJ reference integrity
T7  PJ Non-Vat coverage
T8  end-to-end synthetic posting (insert + update)  (macros must be enabled)
T9  cascade regression (gate wrappers)
T10 source workbook untouched
T11 PENDING items untouched

Every Excel instance is created inside a try/finally so a mid-run exception
cannot leave an orphan process holding the workbook (that is what silently
ate an earlier save).
"""
import datetime
import hashlib
import io
import os
import re
import shutil
import sys
import time

import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CANDIDATE = os.path.join(ROOT, "Global-Smile_2026-v8.4-fix.xlsm")
SOURCE = os.path.join(ROOT, "Global-Smile_2026-v8.4.xlsm")
TMP = os.path.join(os.environ.get("TEMP", "."), "opencode")
OUT = os.path.join(ROOT, "gate2_work", "gate2_test_results.txt")

SOURCE_SHA = "9DF84B7B86482C9D8B7C7FCC6EF4842824E4BF8F47CDFEB848A1BC8D6D56B11C"
TOL = 0.01
REF_SPLIT = 352
PARTITIONS = {
    111: ("JANUARY", [(15, 110), (277, 321)]),
    193: ("FEB", [(113, 192), (322, 407)]),
    408: ("MARCH", [(195, 276)]),
    516: ("APRIL", [(410, 514)]),
    612: ("MAY", [(518, 611)]),
    674: ("JUNE", [(614, 673)]),
}
MONTHS = {1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
          7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"}
NAMES = {v: k for k, v in MONTHS.items()}
NAMES.update({"JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "JUNE": 6,
              "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10,
              "NOVEMBER": 11, "DECEMBER": 12})

results, lines = [], []


def log(msg=""):
    print(msg)
    lines.append(str(msg))


def record(tid, name, status, detail):
    results.append((tid, name, status, detail))
    log("  [%s] %s %s" % (status, tid, name))
    if detail:
        for ln in str(detail).splitlines():
            log("           %s" % ln)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def app(macros=False):
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1 if macros else 3
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    return xl


def quit_xl(xl):
    try:
        xl.Quit()
    except Exception:
        pass


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def norm(v):
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip().upper()


def copy_to(src, name):
    if not os.path.isdir(TMP):
        os.makedirs(TMP)
    dst = os.path.join(TMP, name)
    if os.path.exists(dst):
        os.remove(dst)
    shutil.copy2(src, dst)
    return dst


def cell_mat(ws, r1, r2, c1, c2):
    raw = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value2
    out = []
    for row in raw:
        if isinstance(row, tuple):
            out.append(list(row))
        else:
            out.append([row])
    return out


def main():
    log("GATE 2 VERIFICATION SUITE")
    log("candidate : %s" % CANDIDATE)
    log("mtime     : %s" % time.strftime("%Y-%m-%d %H:%M:%S",
                                        time.localtime(os.path.getmtime(CANDIDATE))))
    log("")

    cdj_cash_by_month = {}

    # ---------------- read-only inspection of the candidate ----------------
    xl = app(macros=False)
    try:
        wb = xl.Workbooks.Open(CANDIDATE, 0, True)
        cdj = wb.Worksheets("CDJ")
        pj = wb.Worksheets("PJ")
        nv = wb.Worksheets("PJ Non-Vat")
        gl = wb.Worksheets("GL")
        xl.CalculateFull()

        log("=== T3 every CDJ data row balances ===")
        rows = cell_mat(cdj, 15, 806, 3, 28)
        bad, worst, n = [], 0.0, 0
        for i, row in enumerate(rows):
            r = 15 + i
            amounts = row[3:19]           # F..U
            if all(blank(v) for v in amounts):
                continue
            n += 1
            x = row[21]                   # X = col 24
            if not isnum(x):
                continue
            worst = max(worst, abs(x))
            if abs(x) > TOL:
                bad.append("r%d X=%.4f S=%s T=%s R=%s" % (r, x, row[16], row[17], row[15]))
        record("T3", "CDJ data rows balance (|X| <= %.2f)" % TOL,
               "PASS" if not bad else "FAIL",
               "rows checked=%d  worst|X|=%.6f  offenders=%d\n%s"
               % (n, worst, len(bad), "\n".join(bad[:20])))

        log("")
        log("=== T4 TOTAL rows equal recomputed partition sums ===")
        t4_fail = []
        for tot, (label, spans) in sorted(PARTITIONS.items()):
            base = tot - 15
            for col in range(6, 22):     # F..U
                want = 0.0
                for a, b in spans:
                    for rr in range(a, b + 1):
                        v = rows[rr - 15][col - 3]
                        if isnum(v):
                            want += v
                got = rows[base][col - 3]
                if not isnum(got):
                    if abs(want) > TOL:
                        t4_fail.append("r%d c%d empty but partition sums %.2f" % (tot, col, want))
                elif abs(got - want) > TOL:
                    t4_fail.append("r%d c%d got=%.4f want=%.4f" % (tot, col, got, want))
            x = rows[base][21]
            if isnum(x) and abs(x) > TOL:
                t4_fail.append("r%d X=%.4f (expected ~0)" % (tot, x))
        record("T4", "TOTAL rows match marker partition and balance",
               "PASS" if not t4_fail else "FAIL",
               "6 TOTAL rows x 16 amount columns + X, all within 0.01"
               if not t4_fail else "\n".join(t4_fail[:25]))

        log("")
        log("=== T6 PJ reference integrity ===")
        pj_last = 8
        for r in range(8, 4000):
            if not blank(pj.Cells(r, 8).Value):
                pj_last = r
        pj_grid = cell_mat(pj, 8, pj_last, 1, 14)
        pj_list = []
        zero_gross = []
        for i, row in enumerate(pj_grid):
            ref = row[7]
            if isinstance(ref, str) and ref.strip().upper().startswith("PJ-"):
                ref = ref.strip().upper()
                pj_list.append(ref)
                if not isnum(row[13]) or abs(row[13]) < 1e-9:
                    zero_gross.append(ref)
        cdj_refs = {}
        for i, row in enumerate(rows):
            r = 15 + i
            for c, idx in ((27, 24), (28, 25)):
                v = row[idx]
                if isinstance(v, str) and v.strip().upper().startswith("PJ-"):
                    cdj_refs.setdefault(v.strip().upper(), set()).add(r)
        old_refs = sorted(x for x in cdj_refs if int(re.sub(r"\D", "", x)) < REF_SPLIT)
        missing = [x for x in sorted(set(pj_list)) if x not in cdj_refs]
        dups = {k: sorted(v) for k, v in cdj_refs.items() if len(v) > 1}
        unexplained = [x for x in missing if x not in zero_gross]
        record("T6", "PJ references: no old refs, no duplicates, all non-zero posted",
               "PASS" if not old_refs and not unexplained and not dups else "FAIL",
               "PJ rows=%d  CDJ refs=%d  old-format=%d  unposted=%d (zero-gross=%d, unexplained=%d)  duplicated=%d\n"
               "  zero-gross, correctly unposted: %s\n  unexplained unposted: %s\n  dups: %s"
               % (len(pj_list), len(cdj_refs), len(old_refs), len(missing), len(zero_gross),
                  len(unexplained), len(dups), zero_gross, unexplained, list(dups.items())[:5]))

        log("")
        log("=== T7 PJ Non-Vat coverage ===")
        nv_grid = cell_mat(nv, 8, 200, 1, 14)
        nv_refs = []
        for i, row in enumerate(nv_grid):
            ref = row[7]
            if isinstance(ref, str) and ref.strip().upper().startswith("PNV-"):
                nv_refs.append((8 + i, ref.strip().upper(), row[13], row[1]))
        posted, pending = [], []
        for r, ref, amt, month in nv_refs:
            hit = None
            for i, row in enumerate(rows):
                v = row[25]
                if isinstance(v, str) and v.strip().upper() == ref:
                    hit = 15 + i
                    break
            if hit:
                posted.append("%s -> CDJ r%d (net=%s)" % (ref, hit, amt))
            else:
                reason = "no month/date in PJ Non-Vat column B" if blank(month) else "not posted"
                pending.append("%s (row %d, net=%s): %s" % (ref, r, amt, reason))
        t7_status = "PASS" if not pending else ("PARTIAL" if posted else "FAIL")
        record("T7", "PJ Non-Vat rows posted to CDJ", t7_status,
               "posted=%d  pending=%d\n  %s%s"
               % (len(posted), len(pending), "\n  ".join(posted),
                  ("\n  PENDING: " + "; ".join(pending)) if pending else ""))

        log("")
        log("=== T11 PENDING items untouched ===")
        t11 = []
        t11.append("r515   C=%s D=%s E=%s F=%s AA=%s" % tuple(rows[515 - 15][c - 3] for c in (3, 4, 5, 6, 27)))
        for r in range(748, 755):
            t11.append("r%-5d C=%s E=%s Y=%s AA=%s" % (r, rows[r - 15][0], rows[r - 15][2],
                                                        rows[r - 15][22], rows[r - 15][24]))
        t11.append("r613   C=%s E=%s X=%s" % (rows[613 - 15][0], rows[613 - 15][2], rows[613 - 15][21]))
        t11.append("r5677  C=%s E=%s AB=%s" % (rows[5677 - 15][0] if 5677 - 15 < len(rows) else None,
                                               rows[5677 - 15][2] if 5677 - 15 < len(rows) else None,
                                               rows[5677 - 15][25] if 5677 - 15 < len(rows) else None))
        f221 = gl.Cells(221, 6).Value
        f222 = gl.Cells(222, 6).Value
        t11.append("GL!F221=%s   (next row GL!F222=%s)" % (f221, f222))
        ok = ("Paid-Up Capital" in str(f221)) and ("Retained Earnings" in str(f222))
        ok = ok and rows[515 - 15][5] is not None
        record("T11", "PENDING items reported, not modified", "PASS" if ok else "FAIL",
               "\n".join(t11))

        # partition cash for T5, computed before the workbook closes
        cur = 0
        for i, row in enumerate(rows):
            m = NAMES.get(norm(row[0]), 0)
            if m:
                cur = m
            if not cur:
                continue
            title = row[2]
            if blank(title) or "TOTAL" in norm(title):
                continue
            v = row[3]
            if isnum(v):
                cdj_cash_by_month[cur] = cdj_cash_by_month.get(cur, 0.0) + v

        wb.Close(False)
    finally:
        quit_xl(xl)
        del xl

    time.sleep(2)

    log("")
    log("=== T1 module compile / callable surface (macros enabled) ===")
    xl = app(macros=True)
    t1_detail, t1_ok = [], True
    try:
        wb = xl.Workbooks.Open(CANDIDATE, 0, True)
        for macro, arg in [("GetFinancialYear", datetime.datetime(2026, 1, 1)),
                           ("FindExpenseColumn", "Clinic Materials and Supplies"),
                           ("ValidateGLConsistency", 2026), ("BuildV8CalcSheet", 2026),
                           ("FindCDJFreeRow", None), ("LoadCDJReferences", None),
                           ("V8_GJMap", "Clinicians Fee- Visiting")]:
            try:
                r = xl.Run(macro, arg) if arg is not None else xl.Run(macro)
                t1_detail.append("%-22s -> %s" % (macro, r))
            except Exception as exc:
                t1_ok = False
                t1_detail.append("%-22s -> FAILED: %s" % (macro, str(exc)[:110]))
        wb.Close(False)
    finally:
        quit_xl(xl)
        del xl
    record("T1", "modules compile and expose the expected surface",
           "PASS" if t1_ok else "FAIL", "\n".join(t1_detail))

    time.sleep(2)

    log("")
    log("=== T9 cascade regression (gate wrappers) ===")
    xl = app(macros=True)
    t9_detail, t9_ok = [], True
    try:
        wb = xl.Workbooks.Open(CANDIDATE, 0, True)
        for macro, arg in [("ValidateGLConsistency", 2026), ("BuildV8CalcSheet", 2026)]:
            try:
                t9_detail.append("%-22s -> %s" % (macro, xl.Run(macro, arg)))
            except Exception as exc:
                t9_ok = False
                t9_detail.append("%-22s -> FAILED: %s" % (macro, str(exc)[:110]))
        wb.Close(False)
    finally:
        quit_xl(xl)
        del xl
    record("T9", "gate wrappers callable from modGLGate", "PASS" if t9_ok else "FAIL",
           "\n".join(t9_detail))

    time.sleep(2)

    log("")
    log("=== T10 source workbook untouched ===")
    got = sha256(SOURCE)
    record("T10", "source SHA-256 unchanged", "PASS" if got == SOURCE_SHA else "FAIL",
           "expected %s\nactual   %s" % (SOURCE_SHA, got))

    log("")
    log("=== T8 end-to-end synthetic posting (throwaway copy, macros enabled) ===")
    t8_copy = copy_to(CANDIDATE, "gate2_t8.xlsm")
    t8_lines, t8_ok = [], True
    xl = app(macros=True)
    try:
        wb = xl.Workbooks.Open(t8_copy, 0, False)
        pj = wb.Worksheets("PJ")
        cdj = wb.Worksheets("CDJ")
        # PJ interleaves month-marker rows and TOTAL rows with the data, so
        # "last row with a reference" is not a safe append point - row 400, for
        # example, already holds the JUNE marker. Find a genuinely empty row.
        last_used = 0
        for r in range(8, 1200):
            for c in range(1, 18):
                if not blank(pj.Cells(r, c).Value):
                    last_used = r
                    break
        synth = last_used + 5
        for c in range(1, 18):
            if not blank(pj.Cells(synth, c).Value):
                t8_ok = False
                t8_lines.append("ABORT: PJ row %d is not empty" % synth)
                raise RuntimeError("no empty PJ row")
        for ref, desc, net, vat, gross, day in [
                ("PJ-99001", "Clinic Materials and Supplies", 1000.0, 120.0, 1120.0, 12),
                ("PJ-99002", "Representation", 2000.0, 240.0, 2240.0, 15)]:
            pj.Cells(synth, 2).Value = "JAN"
            pj.Cells(synth, 3).Value = day
            pj.Cells(synth, 5).Value = "SYNTHETIC SUPPLIER"
            pj.Cells(synth, 7).Value = desc
            pj.Cells(synth, 8).Value = ref
            pj.Cells(synth, 10).Value = net
            pj.Cells(synth, 13).Value = vat
            pj.Cells(synth, 14).Value = gross
            xl.Run("SyncPJToCDJ", pj, synth)
            xl.CalculateFull()
            hit = None
            for r in range(15, 810):
                if str(cdj.Cells(r, 27).Value or "").strip().upper() == ref:
                    hit = r
                    break
            if hit is None:
                t8_ok = False
                t8_lines.append("%s INSERT: row not created" % ref)
                continue
            d = cdj.Cells(hit, 4).Value
            cash = cdj.Cells(hit, 6).Value
            wht = cdj.Cells(hit, 8).Value
            n14 = cdj.Cells(hit, 14).Value
            s = cdj.Cells(hit, 19).Value
            t = cdj.Cells(hit, 20).Value
            rcol = cdj.Cells(hit, 18).Value
            x = cdj.Cells(hit, 24).Value
            t8_lines.append("%s INSERT -> r%d day=%s cash=%s wht=%s N=%s S=%r T=%s R=%s X=%s"
                            % (ref, hit, d, cash, wht, n14, s, t, rcol, x))
            if not isnum(x) or abs(x) > TOL:
                t8_ok = False
                t8_lines.append("   FAIL: does not balance (X=%s)" % x)
            if not isnum(d) or int(d) != day:
                t8_ok = False
                t8_lines.append("   FAIL: day not written to column D (got %s)" % d)
            if not (15 <= hit <= 110 or 277 <= hit <= 321):
                t8_ok = False
                t8_lines.append("   FAIL: row %d is outside the JANUARY partition" % hit)
            fixed = desc.lower().startswith("clinic")
            if fixed:
                if not isnum(n14) or abs(n14 - net) > TOL:
                    t8_ok = False
                    t8_lines.append("   FAIL: fixed path did not place net in N")
                if not blank(t):
                    t8_ok = False
                    t8_lines.append("   FAIL: T must be empty on the fixed path (got %s)" % t)
            else:
                if norm(s) != norm(desc):
                    t8_ok = False
                    t8_lines.append("   FAIL: sundry title not in S (got %r)" % s)
                if not isnum(t) or abs(t - net) > TOL:
                    t8_ok = False
                    t8_lines.append("   FAIL: sundry net not in T (got %s)" % t)
                if not blank(rcol):
                    t8_ok = False
                    t8_lines.append("   FAIL: R must be empty on the sundry path (got %s)" % rcol)

            # UPDATE path - the branch that carried the dangling assignment
            pj.Cells(synth, 14).Value = gross + 500.0
            pj.Cells(synth, 10).Value = net + 500.0
            xl.Run("SyncPJToCDJ", pj, synth)
            xl.CalculateFull()
            cash2 = cdj.Cells(hit, 6).Value
            t2 = cdj.Cells(hit, 20).Value
            x2 = cdj.Cells(hit, 24).Value
            t8_lines.append("%s UPDATE -> r%d cash=%s S=%r T=%s X=%s"
                            % (ref, hit, cash2, cdj.Cells(hit, 19).Value, t2, x2))
            if not isnum(x2) or abs(x2) > TOL:
                t8_ok = False
                t8_lines.append("   FAIL: updated row does not balance (X=%s)" % x2)
            if fixed and not blank(t2):
                t8_ok = False
                t8_lines.append("   FAIL: update branch did not clear T (got %s)" % t2)
            if not fixed and (not isnum(t2) or abs(t2 - (net + 500.0)) > TOL):
                t8_ok = False
                t8_lines.append("   FAIL: update branch did not refresh T (got %s)" % t2)

            for c in range(3, 29):
                cdj.Cells(hit, c).ClearContents()
            cdj.Cells(hit, 22).Formula = "=F%d" % hit
            cdj.Cells(hit, 23).Formula = "=SUM(G%d:T%d)" % (hit, hit)
            cdj.Cells(hit, 24).Formula = "=V%d+W%d" % (hit, hit)
            for c in range(2, 15):
                pj.Cells(synth, c).ClearContents()
        wb.Close(False)
    except Exception as exc:
        t8_ok = False
        t8_lines.append("exception: %s" % str(exc)[:300])
        try:
            wb.Close(False)
        except Exception:
            pass
    finally:
        quit_xl(xl)
        del xl
    record("T8", "synthetic posting through the patched engine",
           "PASS" if t8_ok else "FAIL", "\n".join(t8_lines))
    try:
        os.remove(t8_copy)
    except Exception:
        pass

    time.sleep(2)

    log("")
    log("=== T2 / T5 scan gate (baseline vs patched) ===")
    src_copy = copy_to(SOURCE, "gate2_t2_source.xlsm")
    cand_copy = copy_to(CANDIDATE, "gate2_t2_candidate.xlsm")
    scan = {}
    for label, path in (("baseline", src_copy), ("candidate", cand_copy)):
        xl = app(macros=True)
        try:
            wb = xl.Workbooks.Open(path, 0, False)
            try:
                xl.Run("RefreshGL")
            except Exception as exc:
                scan[label] = ("RUN_ERROR", str(exc)[:200], "", None)
                continue
            xl.CalculateFull()
            aud = wb.Worksheets("GL_AUDIT")
            last = aud.UsedRange.Row + aud.UsedRange.Rows.Count - 1
            status = str(aud.Cells(last, 2).Value)
            summary = str(aud.Cells(last, 3).Value)
            detail = str(aud.Cells(last, 4).Value or "")
            gl = wb.Worksheets("GL")
            cash = {}
            for r in range(13, 25):
                cash[str(gl.Cells(r, 2).Value)] = (gl.Cells(r, 7).Value, gl.Cells(r, 8).Value)
            scan[label] = (status, summary, detail, cash)
            wb.Close(False)
        finally:
            quit_xl(xl)
            del xl
        time.sleep(2)

    for label in ("baseline", "candidate"):
        st, summ, det, _ = scan[label]
        log("  %-9s status=%s" % (label, st))
        log("            %s" % summ)
        if det:
            log("            detail lines=%d" % len([p for p in det.splitlines() if p.strip()]))

    st_b, sum_b, det_b, _ = scan["baseline"]
    st_c, sum_c, det_c, cash_c = scan["candidate"]

    def parse(summary):
        out = {}
        for k in ("Debit", "Credit", "Difference", "Unmapped", "Invalid"):
            m = re.search(k + r"=([0-9.,-]+)", summary)
            if m:
                out[k] = float(m.group(1).replace(",", ""))
        return out

    pb, pc = parse(sum_b), parse(sum_c)
    t2_detail = ["baseline : %s | %s" % (st_b, sum_b),
                 "candidate: %s | %s" % (st_c, sum_c)]
    for k in ("Debit", "Credit", "Difference", "Unmapped", "Invalid"):
        if k in pb and k in pc:
            t2_detail.append("  %-10s %14.2f -> %14.2f   (%+.2f)"
                             % (k, pb[k], pc[k], pc[k] - pb[k]))
    titles = {}
    for p in (det_c or "").splitlines():
        if "->" in p and "unmapped" in p.lower():
            t = p.split("->", 1)[1].rsplit("unmapped", 1)[0].strip()
            titles[t] = titles.get(t, 0) + 1
    t2_detail.append("remaining unmapped sundry titles (%d):" % len(titles))
    for t in sorted(titles):
        t2_detail.append("   %-40s x%d" % (t, titles[t]))
    record("T2", "RefreshGL scan gate: baseline vs patched", "HOLD" if st_c.upper().startswith("HOLD") else ("PASS" if st_c.upper().startswith("PASS") else "FAIL"),
           "\n".join(t2_detail))

    if (cash_c or {}).get("JAN") is None or st_c.upper().startswith("HOLD"):
        record("T5", "CDJ partition cash vs scanned GL", "NOT RUN",
               "scan gate HOLDs, so ScanGL correctly refuses to overwrite the GL; "
               "there is no refreshed GL to compare against.\n"
               "CDJ partition cash (for reference):\n" + "\n".join(
                   "   %s %14.2f" % (MONTHS[m], cdj_cash_by_month.get(m, 0.0)) for m in range(1, 13)))
    else:
        t5_fail = []
        for m in range(1, 7):
            abbr = MONTHS[m]
            d, c = cash_c.get(abbr, (0.0, 0.0))
            gl_net = (d or 0.0) - (c or 0.0)
            cdj_net = cdj_cash_by_month.get(m, 0.0)
            log("    %s CDJ=%14.2f  GLnet=%14.2f  diff=%+.4f" % (abbr, cdj_net, gl_net, cdj_net - gl_net))
            if abs(cdj_net - gl_net) > 0.05:
                t5_fail.append("%s CDJ=%.2f GLnet=%.2f diff=%.2f" % (abbr, cdj_net, gl_net, cdj_net - gl_net))
        record("T5", "CDJ partition cash agrees with scanned GL",
               "PASS" if not t5_fail else "FAIL",
               "\n".join(t5_fail) if t5_fail else "all six months agree within 0.05")

    for p in (src_copy, cand_copy):
        try:
            os.remove(p)
        except Exception:
            pass

    log("")
    log("=== RESULT SUMMARY ===")
    counts = {}
    for tid, name, status, _ in results:
        counts[status] = counts.get(status, 0) + 1
        log("  %-4s %-8s %s" % (tid, status, name))
    log("")
    log("  " + "  ".join("%s=%d" % kv for kv in sorted(counts.items())))

    with io.open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print("\nresults -> %s" % OUT)
    return 0 if counts.get("FAIL", 0) == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
