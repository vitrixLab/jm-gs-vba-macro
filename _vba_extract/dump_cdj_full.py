"""Full row dump: values + formulas for CDJ old (plan) vs CDJ (implementation)."""
import re, os

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

def load(name):
    target = rid_to_target[sheets[name]]
    xml = read(os.path.join(OUT, "xl", target.replace("/", os.sep)))
    grid = {}
    for m in CELL_RE.finditer(xml):
        col, rown, attrs, _, inner = m.groups()
        t = re.search(r' t="(\w+)"', attrs)
        ta = t.group(1) if t else None
        fm = re.search(r"<f[^>]*>(.*?)</f>", inner or "", re.S)
        vm = re.search(r"<v>(.*?)</v>", inner or "", re.S)
        raw = vm.group(1) if vm else ""
        if raw == "" and ta != "str" and not fm:
            continue
        val = shared[int(raw)] if ta == "s" else raw
        grid[(int(rown), col)] = (fm.group(1) if fm else "", val)
    return grid

def cols():
    o = [chr(c) for c in range(ord('A'), ord('Z')+1)]
    o += ["A"+chr(c) for c in range(ord('A'), ord('Z')+1)]
    return o

def dump(name, rows):
    grid = load(name)
    print("="*120)
    print(f"### {name}")
    for r in rows:
        print(f"--- row {r} ---")
        for c in cols():
            f, v = grid.get((r, c), ("", ""))
            if f or v:
                print(f"  {c}{r}: val={v[:70]!r} formula={f[:110]!r}")

import sys
which = sys.argv[1] if len(sys.argv) > 1 else "both"
if which in ("old", "both"):
    dump("CDJ old", [13, 14, 15, 28, 31, 37, 48, 53, 57, 58, 62, 63, 64, 67, 68, 69, 70])
if which in ("new", "both"):
    dump("CDJ", [13, 14, 15, 16, 58, 68])
