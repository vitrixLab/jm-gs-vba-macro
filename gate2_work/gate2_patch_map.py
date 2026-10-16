"""Gate 2 — patch modGLWorkbookMap.bas with clinician-fee spelling aliases.

The CDJ sundry column carries several spellings of the two clinician-fee
accounts that already exist in the chart of accounts. They differ only by an
apostrophe, an extra space around the hyphen, or a truncated word, so mapping
them is mechanical and needs no accounting judgement. Without this the scan
gate holds purely on typing differences.

Titles with no chart-of-accounts equivalent (Representation, Partners
Clinicians Fee, Incentives, Outside Services, SSS/HDMF/PHIC ER Share, Training
and Seminar) are deliberately NOT mapped here - they need an operator decision.
"""
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "gate2_work", "vba_baseline", "modGLWorkbookMap.bas")
DST = os.path.join(ROOT, "gate2_work", "vba_patched", "modGLWorkbookMap.bas")

ANCHOR = "        Case Else: If d.Exists(k) Then V8_GJMap = d(k)"

NEW_CASES = """        ' v8.4 Gate 2: CDJ sundry titles that are spelling variants of the two
        ' clinician-fee accounts already in the chart of accounts. They differ
        ' only by an apostrophe, spacing around the hyphen, or a truncated word,
        ' so the mapping is mechanical and needs no accounting judgement.
        Case "CLINICIANS FEE- VISITING", "CLINICIANS FEE - VISITING", _
             "CLINICIANS FEE - VISITING'", "CLINICIANS FEE-VISITING", _
             "VISITING CLINICIANS FEE"
            V8_GJMap = "Clinician's Fee-Visiting"
        Case "CLINICIANS FEE- CORPORATORS", "CLINICIANS FEE -CORPORATORS", _
             "CLINICIANS FEE- CORPORATOS", "CLINACIANS FEE- CORPORATORS"
            V8_GJMap = "Clinician's Fee-Corporators"
"""


def main():
    with io.open(SRC, "r", encoding="utf-8", newline="") as fh:
        text = fh.read()

    if "v8.4 Gate 2" in text:
        print("already patched; writing identical output")
    else:
        if ANCHOR not in text:
            raise SystemExit("ABORT: anchor not found in baseline modGLWorkbookMap.bas")
        if text.count(ANCHOR) != 1:
            raise SystemExit("ABORT: anchor not unique")
        text = text.replace(ANCHOR, NEW_CASES + ANCHOR)
        print("inserted %d alias cases before the Case Else fallback" % 9)

    if not os.path.isdir(os.path.dirname(DST)):
        os.makedirs(os.path.dirname(DST))
    with io.open(DST, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)
    print("OK -> %s" % DST)

    # sanity: the emitted VBA must keep balanced Case / End Select
    n_case = text.count("\n        Case ")
    n_end = text.count("    End Select")
    print("  Case statements in V8_GJMap region: %d, End Select: %d" % (n_case, n_end))
    return 0


if __name__ == "__main__":
    sys.exit(main())
