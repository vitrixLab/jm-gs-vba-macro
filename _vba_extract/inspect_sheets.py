"""Inspect worksheet layout v2: exact cell addressing, used-range map."""
import re
import os

BASE = r"C:\citrixlabph\globalsmile\graphify"


def read(p):
    with open(p, encoding="utf-8", errors="replace") as f:
        return f.read()


wb = read(os.path.join(BASE, r"xl\workbook.xml"))
rels = read(os.path.join(BASE, r"xl\_rels\workbook.xml.rels"))
ss = read(os.path.join(BASE, r"xl\sharedStrings.xml"))

rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)
shared = re.findall(r"<si>(.*?)</si>", ss, re.S)
shared = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", s, re.S)) or "" for s in shared]

CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)


def decode(t_attr, raw):
    val = raw
    if t_attr == "s" and val != "":
        return shared[int(val)]
    return val


for name, rid in sheets:
    target = rid_to_target.get(rid)
    if not target:
        continue
    xml = read(os.path.join(BASE, "xl", target.replace("/", os.sep)))
    dim = re.search(r'<dimension ref="([^"]+)"', xml)
    print("=" * 72)
    print(f"### {name}  ({target})  dimension={dim.group(1) if dim else '?'}")

    colA = {}
    used_cols = {}
    for m in CELL_RE.finditer(xml):
        col, rown, attrs, _, inner = m.groups()
        t = re.search(r' t="(\w+)"', attrs)
        t_attr = t.group(1) if t else None
        vm = re.search(r"<v>(.*?)</v>", inner or "", re.S)
        val = vm.group(1) if vm else ""
        val = decode(t_attr, val)
        rn = int(rown)
        if val not in ("", None):
            used_cols.setdefault(col, []).append(rn)
            if col == "A":
                colA[rn] = val
    print("used columns:", {c: (len(v), f"{min(v)}..{max(v)}") for c, v in sorted(used_cols.items())})
    aitems = {k: colA[k] for k in sorted(colA)[:8]}
    print("col A first 8:", aitems)
    aitems2 = {k: colA[k] for k in sorted(colA)[-3:]}
    print("col A last 3:", aitems2)

