"""month_labels_v82.py - month-label + block inventory for CDJ/CRJ/GJ/SJ (read-only)."""
import os
import sys

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(BASE, "Global-Smile_2026-v8.2-final-review.xlsm")
wb = openpyxl.load_workbook(WB, keep_vba=True, data_only=True)

print("--- CDJ: month labels in col C, data rows 15..%d ---" % wb["CDJ"].max_row)
cdj = wb["CDJ"]
labels = [(r, cdj.cell(r, 3).value, cdj.cell(r, 5).value) for r in range(15, 1005)
          if str(cdj.cell(r, 3).value or "").strip()]
print("label cells: %d" % len(labels))
for r, c, e in labels:
    print("   C%-4s = %-6s  E=%s" % (r, c, (e or "")[:35]))
data = [r for r in range(15, 1005)
        if any(cdj.cell(r, c).value not in (None, "") for c in range(5, 22))]
print("data rows (E..U non-empty) = %d -> %s" % (len(data), data))

print()
print("--- CRJ: monthly TOTAL blocks (col F=TOTAL, col D=label) ---")
crj = wb["CRJ"]
for r in range(10, 720):
    if str(crj.cell(r, 6).value or "").strip().upper() == "TOTAL":
        print("   row %-4s D=%-5s E=%-3s H(debit)=%-10s J=%-9s L=%s" %
              (r, crj.cell(r, 4).value, crj.cell(r, 5).value,
               crj.cell(r, 8).value, crj.cell(r, 10).value, crj.cell(r, 12).value))
n = [r for r in range(10, 720) if any(crj.cell(r, c).value not in (None, "") for c in (6, 8, 9, 10, 11, 12))]
print("CRJ rows with particulars/amounts = %d -> %s" % (len(n), n))

print()
print("--- GJ: month labels in col B, rows 10..1000 ---")
gj = wb["GJ"]
for r in range(10, 1001):
    v = str(gj.cell(r, 2).value or "").strip()
    if v and not v.replace(".", "").isdigit():
        print("   B%-4s = %-6s  E=%s" % (r, v, (gj.cell(r, 5).value or "")[:40]))
rows = [r for r in range(10, 1001) if gj.cell(r, 5).value not in (None, "")]
print("GJ rows with explanation = %d ; first=%s last=%s" % (len(rows), rows[0] if rows else None,
                                                             rows[-1] if rows else None))
print("   GJ label rows: %s" % [(r, gj.cell(r, 5).value) for r in rows])

print()
print("--- SJ: dates and totals ---")
sj = wb["SJ"]
for r in range(1, 40):
    vals = [(c, sj.cell(r, c).value) for c in range(1, 12) if sj.cell(r, c).value not in (None, "")]
    if vals:
        print("   row %-3s %s" % (r, " | ".join("%s=%s" % v for v in vals)))
wb.close()
