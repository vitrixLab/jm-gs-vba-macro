"""Dump sheet titles/headers for given sheets of the CURRENT v7.8 workbook.

Usage: python inspect_titles.py <SheetName1> [SheetName2 ...]
Rows 1-14 are printed as 'col: value' per sheet, plus numeric density per column.
"""
import re
import os
import sys
import zipfile

XLSM = r"C:\citrixlabph\globalsmile\Global-Smile_2026-v7.8.xlsm"
OUT = r"C:\citrixlabph\globalsmile\_vba_extract\v78_unzip"


def col_letters(idx):
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    if not os.path.isdir(OUT):
        os.makedirs(OUT)
        with zipfile.ZipFile(XLSM) as zf:
            zf.extractall(OUT)

    def read(p):
        with open(p, encoding="utf-8", errors="replace") as f:
            return f.read()

    wb = read(os.path.join(OUT, r"xl\workbook.xml"))
    rels = read(os.path.join(OUT, r"xl\_rels\workbook.xml.rels"))
    ss = read(os.path.join(OUT, r"xl\sharedStrings.xml"))
    rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb)
    print("ALL SHEETS:", " | ".join(n for n, _ in sheets))
    shared = re.findall(r"<si>(.*?)</si>", ss, re.S)
    shared = ["".join(re.findall(r"<t[^>]*>(.*?)</t>", s, re.S)) or "" for s in shared]
    CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)

    wanted = {w.lower(): w for w in sys.argv[1:]}
    for name, rid in sheets:
        if name.lower() not in wanted:
            continue
        target = rid_to_target.get(rid)
        xml = read(os.path.join(OUT, "xl", target.replace("/", os.sep)))
        dim = re.search(r'<dimension ref="([^"]+)"', xml)
        print("\n" + "=" * 78)
        print(f"### {name}  ({target})  dimension={dim.group(1) if dim else '?'}")
        grid = {}
        numeric = {}
        for m in CELL_RE.finditer(xml):
            col, rown, attrs, _, inner = m.groups()
            t = re.search(r' t="(\w+)"', attrs)
            t_attr = t.group(1) if t else None
            vm = re.search(r"<v>(.*?)</v>", inner or "", re.S)
            raw = vm.group(1) if vm else ""
            if raw == "" and t_attr != "str":
                continue
            if t_attr == "s":
                val = shared[int(raw)]
                kind = "text"
            elif t_attr in (None, "n"):
                val = raw
                kind = "num"
            elif t_attr == "str":
                val = raw
                kind = "fx"
            else:
                val = raw
                kind = t_attr
            rn = int(rown)
            grid[(rn, col)] = (kind, val)
            if kind == "num":
                c = numeric.get(col, 0)
                numeric[col] = c + 1
        # numeric density
        print("numeric cells per column:", {c: numeric[c] for c in sorted(numeric, key=lambda s: (len(s), s))})
        # dump rows 1..14
        for rn in range(1, 15):
            cells = sorted([(c, v) for (r, c), v in grid.items() if r == rn],
                           key=lambda x: (len(x[0]), x[0]))
            if not cells:
                continue
            line = ", ".join(f"{c}={v[1][:38]}" for c, v in cells)
            print(f"  r{rn}: {line}")


if __name__ == "__main__":
    main()
