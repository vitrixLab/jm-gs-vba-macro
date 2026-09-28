"""F2/F4 evidence round 2 - CDJ sundry design, GL titles, codeName mapping."""
import os
import re
import zipfile

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, "Global-Smile_2026-v8.0.xlsm")


def colname(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


wb = load_workbook(WB, data_only=True)

print("== codeName -> sheet name ==")
with zipfile.ZipFile(WB) as z:
    wbxml = z.read("xl/workbook.xml").decode("utf-8", "replace")
names = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml)
with zipfile.ZipFile(WB) as z:
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
rid2target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
for nm, rid in names:
    target = rid2target.get(rid, "")
    if "worksheets" in target:
        path = "xl/" + target.lstrip("/")
        try:
            data = z.read(path).decode("utf-8", "replace") if False else None
        except Exception:
            pass
# simpler: openpyxl internal
for ws in wb.worksheets:
    print(f"  {ws.title!r:<20} codeName={ws.sheet_properties.codeName}")

print()
print("== GL account titles (col F, rows 12..70) ==")
ws = wb["GL"]
for r in range(12, 71):
    v = ws.cell(row=r, column=6).value
    if v is not None and str(v).strip():
        print(f"  r{r}: {v!r}")
    if r > 12 and v is None and ws.cell(row=r, column=5).value is None:
        # stop after first blank gap beyond row 20
        pass

print()
print("== CDJ rows where S (19) is non-empty ==")
ws = wb["CDJ"]
for r in range(15, ws.max_row + 1):
    s = ws.cell(row=r, column=19).value
    if s is not None and str(s).strip():
        full = {colname(c): ws.cell(row=r, column=c).value for c in range(3, 22)}
        full = {k: v for k, v in full.items() if v is not None}
        print(f"  r{r}: {full}")

print()
print("== CDJ rows where Q (17) is non-empty ==")
for r in range(15, ws.max_row + 1):
    q = ws.cell(row=r, column=17).value
    if q is not None and str(q).strip():
        full = {colname(c): ws.cell(row=r, column=c).value for c in range(3, 22)}
        full = {k: v for k, v in full.items() if v is not None}
        print(f"  r{r}: {full}")

print()
print("== CDJ row 17 full (C..U) ==")
vals = []
for c in range(3, 22):
    v = ws.cell(row=17, column=c).value
    vals.append(f"{colname(c)}={v!r}")
print("  " + " | ".join(vals))

print()
print("== CDJ T/U totals rows 13/14 and last data row ==")
for r in (13, 14):
    print(f"  r{r}: T={ws.cell(row=r, column=20).value!r} U={ws.cell(row=r, column=21).value!r}")
last = ws.max_row
print(f"  max_row={last}; r{last}: E={ws.cell(row=last, column=5).value!r} T={ws.cell(row=last, column=20).value!r} U={ws.cell(row=last, column=21).value!r}")
