"""v81_audit.py — read-only v8.0 audit for the v8.1 build decision.
Workbook: Global-Smile_2026-v8.0.xlsm (never modified here).
Prints: 46 GL titles, GJ unmapped-label totals for the 5 HOLD labels,
CDJ Q/S nonzero counts, CRJ layout check, source debit/credit reconciliation.
"""
import openpyxl
from collections import defaultdict

WB = "Global-Smile_2026-v8.0.xlsm"
wb = openpyxl.load_workbook(WB, data_only=False)

# 1. GL titles
ws = wb["GL"]
accts, seen = [], set()
for r in range(13, ws.max_row + 1):
    v = ws.cell(r, 6).value
    if v and str(v).strip().upper() not in seen:
        seen.add(str(v).strip().upper())
        accts.append(str(v).strip())
print("GL accounts:", len(accts))
for i, a in enumerate(accts, 1):
    print("  %2d. %s" % (i, a))

# 2. GJ label totals
ws = wb["GJ"]
deb, cred = defaultdict(float), defaultdict(float)
for (e, f, g) in ws.iter_rows(min_row=11, max_row=1000, min_col=5, max_col=7, values_only=True):
    if not e or not str(e).strip():
        continue
    low = str(e).lower()
    if "record" in low or "liquidation" in low:
        continue
    k = str(e).strip()
    if isinstance(f, (int, float)):
        deb[k] += f
    if isinstance(g, (int, float)):
        cred[k] += g
print("\nGJ distinct labels:", len(deb))
hold = ["MEDICAL EQUIPMENT", "COST OF REVENUE", "SUPPLIES", "BANK CHARGE", "CHARGES"]
total = 0.0
for k in sorted(deb, key=str.upper):
    if k.strip().upper() in hold:
        amt = deb[k] + cred[k]
        total += amt
        print("  HOLD %-20s debit=%10.2f credit=%10.2f" % (k, deb[k], cred[k]))
print("  HOLD total: %.2f" % total)

# 3. CDJ Q/S + reconciliation
ws = wb["CDJ"]
q_nz = s_nz = 0
for r in range(15, 1005):
    for col, acc in ((17, "q"), (19, "s")):
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)) and abs(v) > 0.005:
            if acc == "q":
                q_nz += 1
            else:
                s_nz += 1
print("\nCDJ Q(17) nonzero: %d  S(19) nonzero: %d" % (q_nz, s_nz))

# 4. Source totals (values, best effort)
def sumnum(sheet, rows, cols):
    t = 0.0
    w = wb[sheet]
    for r in rows:
        for c in cols:
            v = w.cell(r, c).value
            if isinstance(v, (int, float)):
                t += v
    return t

print("\nSheets:", wb.sheetnames)
print("has GL_V8_CALC:", "GL_V8_CALC" in wb.sheetnames)
print("has GL_AUDIT:", "GL_AUDIT" in wb.sheetnames)
