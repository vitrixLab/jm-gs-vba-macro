"""Compare TITLE layouts between CDJ old (plan) and CDJ (implementation) — plan-first review."""
import re, os

OUT = r"C:\citrixlabph\globalsmile\_vba_extract\v78_unzip"

def read(p):
    with open(p, encoding="utf-8", errors="replace") as f:
        return f.read()

wb = read(os.path.join(OUT, "xl\workbook.xml"))
rels = read(os.path.join(OUT, "xl\_rels\workbook.xml.rels"))
sspath = os.path.join(OUT, r"xl\sharedStrings.xml")
ss = open(sspath, encoding="utf-8", errors="replace").read()
sheets_for_s = ["CDJ old", "CDJ"]
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
        fm = re.search(r"<f>(.*?)</f>", inner or "", re.S)
        vm = re.search(r"<v>(.*?)</v>", inner or "", re.S)
        raw = vm.group(1) if vm else ""
        if raw == "" and ta != "str":
            continue
        val = shared[int(raw)] if ta == "s" else raw
        grid[(int(rown), col)] = (fm.group(1)[:70] if fm else "", val[:70])
    return grid

def col_order():
    o = []
    for a in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        o.append(a)
    for a in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        o.append("A" + a)
    return o

print("=" * 110)
print(f"### TITLE LAYOUT — CDJ old (PLAN) vs CDJ (IMPLEMENTATION)")
print(f"Rows 13–14 = headers (r13 top, r14 subcategory). Values from r15 snapshot.")
print("=" * 110)

for name in ["CDJ old", "CDJ"]:
    grid = load(name)
    print(f"\n##### {name}")
    print(f"{'Col':<4} {'r13 title (top)':<26} {'r14 title (sub)':<22} {'r15 sample':<48} formula?")
    for c in col_order():
        h13, v13 = grid.get((13, c), ("", ""))
        h14, v14 = grid.get((14, c), ("", ""))
        v15, f15 = grid.get((15, c), ("", ""))
        if not (h13 or h14 or v15):
            continue
        print(f"{c:<4} {h13[:25]:<26} {h14[:21]:<22} {v15[:47]:<48} {f15[:40]}")
