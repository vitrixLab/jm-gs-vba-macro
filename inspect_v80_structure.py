"""F3/F4 evidence - read-only structural inspection of Global-Smile_2026-v8.0.xlsm.

Dumps the exact header rows and sample data the v8 mapping depends on:
  CDJ rows 13/14 (cols A..T)  -> decide Q(17)/S(19) classification
  CDJ rows 15..20 cols C,E,Q,S -> whether Q/S ever hold posting values
  CRJ rows 7..9 (cols A..Q)   -> canonical v8 CRJ layout consumed by the engine
  CRJ sample rows 10..16 cols D,F,G,H,J,K,L,O,P,Q
  GJ rows 9/10 (cols A..H)
  GL rows 11/13..15 (cols E..F)
  SJ header rows (first 12)   -> what CreateCRJFromSales consumes
  sheet list, GL_V8_CALC presence, GL_AUDIT tail
"""
import os

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, "Global-Smile_2026-v8.0.xlsm")


def colname(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def row_dump(ws, row, c1, c2, label):
    vals = []
    for c in range(c1, c2 + 1):
        v = ws.cell(row=row, column=c).value
        if v is not None:
            vals.append(f"{colname(c)}={v!r}")
    print(f"{label} r{row}: " + (" | ".join(vals) if vals else "(empty)"))


wb = load_workbook(WB, data_only=True, keep_vba=False, read_only=False)
print("sheets:", wb.sheetnames)
print("has GL_V8_CALC:", "GL_V8_CALC" in wb.sheetnames)
print()

if "CDJ" in wb.sheetnames:
    ws = wb["CDJ"]
    for r in (13, 14):
        row_dump(ws, r, 1, 20, "CDJ header")
    print()
    for r in range(15, 21):
        row_dump(ws, r, 3, 5, "CDJ data key")
        row_dump(ws, r, 17, 17, "CDJ col Q")
        row_dump(ws, r, 19, 19, "CDJ col S")
    # scan whole used range for non-zero Q/S
    last = ws.max_row
    qnz = snz = 0
    for r in range(15, last + 1):
        for col, counter in ((17, "q"), (19, "s")):
            v = ws.cell(row=r, column=col).value
            if isinstance(v, (int, float)) and abs(v) > 0.001:
                if counter == "q":
                    qnz += 1
                else:
                    snz += 1
    print(f"CDJ nonzero counts -> Q(17): {qnz}  S(19): {snz}  (rows 15..{last})")
    print()

if "CRJ" in wb.sheetnames:
    ws = wb["CRJ"]
    for r in (7, 8, 9):
        row_dump(ws, r, 1, 17, "CRJ header")
    for r in range(10, 17):
        row_dump(ws, r, 4, 4, "CRJ")
        row_dump(ws, r, 6, 8, "CRJ")
        row_dump(ws, r, 10, 12, "CRJ")
        row_dump(ws, r, 15, 16, "CRJ")
    print()

if "GJ" in wb.sheetnames:
    ws = wb["GJ"]
    for r in (9, 10, 11):
        row_dump(ws, r, 1, 10, "GJ header")
    print()

if "GL" in wb.sheetnames:
    ws = wb["GL"]
    for r in (11, 12, 13, 14, 15):
        row_dump(ws, r, 5, 8, "GL")
    last = ws.max_row
    print(f"GL max_row={last}")
    print()

if "SJ" in wb.sheetnames:
    ws = wb["SJ"]
    print(f"SJ dims: {ws.dimensions}")
    for r in range(1, 13):
        row_dump(ws, r, 1, 14, "SJ")
    print()

if "GL_AUDIT" in wb.sheetnames:
    ws = wb["GL_AUDIT"]
    last = ws.max_row
    print(f"GL_AUDIT max_row={last}, tail:")
    for r in range(max(2, last - 7), last + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 5)]
        print("  ", [str(v)[:120] if v is not None else None for v in vals])
