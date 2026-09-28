import re
import zipfile

from openpyxl import load_workbook

WB = r"D:\citrixlabph\globalsmile\Global-Smile_2026-v8.0.xlsm"

wb = load_workbook(WB, data_only=False)
ws = wb["GL"]

print("== GL rows 442-615 with G/H/I content ==")
for r in range(442, 616):
    out = []
    for c in (2, 3, 5, 6, 7, 8, 9):
        v = ws.cell(row=r, column=c).value
        if v is not None and str(v).strip():
            out.append(f"{chr(64+c)}={v}")
    if any(o[0] in "GHI" for o in out):
        print(f"r{r}:", " | ".join(out))

print()
print("== CDJ headers N13:U14 ==")
cd = wb["CDJ"]
for r in (13, 14):
    print(f"r{r}:", [cd.cell(row=r, column=c).value for c in range(14, 22)])

print()
print("== CRJ rows 5-9, 24-26, 35-37, 43-45, 117-119 cols D..Q ==")
cr = wb["CRJ"]
for r in list(range(5, 10)) + [25, 36, 44, 118]:
    print(f"r{r}:", [cr.cell(row=r, column=c).value for c in range(4, 18)])

print()
print("== drawings/vml with macros ==")
with zipfile.ZipFile(WB) as z:
    for n in z.namelist():
        if "drawing" in n.lower() and n.endswith(".xml"):
            d = z.read(n).decode("utf-8", "replace")
            ms = re.findall(r'macro="([^"]+)"', d)
            if ms:
                print(n, "->", ms)
        if n.endswith(".vml"):
            d = z.read(n).decode("utf-8", "replace")
            ms = re.findall(r'macroname="([^"]+)"', d)
            if ms:
                print(n, "->", ms)
print("(done)")
