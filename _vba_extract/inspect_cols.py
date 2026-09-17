"""Classify per-column content on a sheet: numeric vs text counts + samples."""
import re
import os
import sys
import datetime

BASE = r"C:\citrixlabph\globalsmile\graphify"


def read(p):
    with open(p, encoding="utf-8", errors="replace") as f:
        return f.read()


wb = read(os.path.join(BASE, r"xl\workbook.xml"))
rels = read(os.path.join(BASE, r"xl\_rels\workbook.xml.rels"))
ss = read(os.path.join(BASE, r"xl\sharedStrings.xml"))
rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
sheets = dict(re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb))
shared = re.findall(r"<si>(.*?)</si>", ss, re.S)
shared = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", s, re.S)) or "" for s in shared]

name = sys.argv[1] if len(sys.argv) > 1 else "PJ"
target = rid_to_target[sheets[name]]
xml = read(os.path.join(BASE, "xl", target.replace("/", os.sep)))

CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)

cols = {}
for m in CELL_RE.finditer(xml):
    col, rown, attrs, _, inner = m.groups()
    t = re.search(r' t="(\w+)"', attrs)
    t_attr = t.group(1) if t else None
    vm = re.search(r"<v>(.*?)</v>", inner or "", re.S)
    raw = vm.group(1) if vm else ""
    if raw == "" and t_attr != "str":
        continue
    if t_attr == "s":
        kind, shown = "text", shared[int(raw)][:40]
    elif t_attr in (None, "n"):
        x = float(raw)
        if 20000 < x < 60000 and x == int(x):
            kind = "num(maybe-date)"
        else:
            kind = "num"
        shown = f"{raw[:20]}" + (f" ~{datetime.date(1899,12,30)+datetime.timedelta(days=x)}" if kind.startswith('num(maybe') else "")
    elif t_attr == "str":
        kind, shown = "fx-text", raw[:40]
    elif t_attr == "b":
        kind, shown = "bool", raw
    else:
        kind, shown = t_attr or "?", raw[:40]
    cols.setdefault(col, []).append((int(rown), kind, shown))

for c in sorted(cols, key=lambda s: (len(s), s)):
    items = cols[c]
    counts = {}
    for _, k, _ in items:
        counts[k] = counts.get(k, 0) + 1
    sample = ", ".join(f"r{r}:{s}" for r, _, s in items[:4])
    print(f"{c:>3} n={len(items):4}  {str(counts):60}  rows {min(r for r,_,_ in items)}..{max(r for r,_,_ in items)}  | {sample}")
