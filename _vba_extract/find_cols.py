"""Find rows in a sheet where given columns have values (to infer column semantics)."""
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
shared = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", s, re.S)) or "" for s in re.findall(r"<si>(.*?)</si>", ss, re.S)]
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)

name = sys.argv[1]
cols = set(sys.argv[2].split(","))
target = rid_to_target[sheets[name]]
xml = read(os.path.join(OUT, "xl", target.replace("/", os.sep)))
rows = {}
for m in CELL_RE.finditer(xml):
    col, rown, attrs, _, inner = m.groups()
    t = re.search(r' t="(\w+)"', attrs)
    ta = t.group(1) if t else None
    fm = re.search(r"<f>(.*?)</f>", inner or "", re.S)
    vm = re.search(r"<v>(.*?)</v>", inner or "", re.S)
    raw = vm.group(1) if vm else ""
    if raw == "":
        continue
    val = shared[int(raw)] if ta == "s" else raw
    if fm:
        val = f"[f:{fm.group(1)[:40]}={val[:18]}]"
    rn = int(rown)
    if col in cols and rn >= 15:
        rows.setdefault(rn, {})[col] = val
for rn in sorted(rows):
    print(rn, rows[rn])
