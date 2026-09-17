"""Dump data rows 15-30 of a sheet with all columns, to see where values really go."""
import re
import os
import sys

OUT = r"C:\citrixlabph\globalsmile\_vba_extract\v78_unzip"


def read(p):
    with open(p, encoding="utf-8", errors="replace") as f:
        return f.read()


wb = read(os.path.join(OUT, r"xl\workbook.xml"))
rels = read(os.path.join(OUT, r"xl\_rels\workbook.xml.rels"))
ss = read(os.path.join(OUT, r"xl\sharedStrings.xml"))
rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
sheets = dict(re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb))
shared = re.findall(r"<si>(.*?)</si>", ss, re.S)
shared = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", s, re.S)) or "" for s in shared]
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)

name = sys.argv[1]
r1, r2 = int(sys.argv[2]), int(sys.argv[3])
target = rid_to_target[sheets[name]]
xml = read(os.path.join(OUT, "xl", target.replace("/", os.sep)))
grid = {}
for m in CELL_RE.finditer(xml):
    col, rown, attrs, _, inner = m.groups()
    t = re.search(r' t="(\w+)"', attrs)
    t_attr = t.group(1) if t else None
    fm = re.search(r"<f>(.*?)</f>", inner or "", re.S)
    vm = re.search(r"<v>(.*?)</v>", inner or "", re.S)
    raw = vm.group(1) if vm else ""
    if raw == "" and t_attr != "str":
        continue
    if t_attr == "s":
        val = shared[int(raw)]
    else:
        val = raw
    if fm:
        val = f"[f:{fm.group(1)[:30]} -> {val[:20]}]"
    grid[(int(rown), col)] = val

for rn in range(r1, r2 + 1):
    cells = sorted([(c, v) for (r, c), v in grid.items() if r == rn], key=lambda x: (len(x[0]), x[0]))
    if cells:
        print(f"r{rn}: " + " | ".join(f"{c}={v[:42]}" for c, v in cells))
    else:
        print(f"r{rn}: (empty)")
