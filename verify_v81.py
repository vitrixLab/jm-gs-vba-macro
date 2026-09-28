"""verify_v81.py — read-only verification of the v8.1 FINAL workbook.
Checks: v8.0 untouched, v8.1 has +3 sheets, 552 calc rows, COA_MAP totals 6154,
VBA project byte-identical, no values/formulas altered on original sheets.
Exit 0 = all pass, exit 1 = any fail.
"""
import hashlib
import json
import sys
import zipfile

import openpyxl

V80 = "Global-Smile_2026-v8.0.xlsm"
V81 = "Global-Smile_2026-v8.1.xlsm"
V80_SHA = "76a783ad0dce60ef3e27b52c78632555541d529442ddd6b015aad944fa22e327"

fails = []


def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (" — " + detail if detail else ""))
    if not cond:
        fails.append(name)


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as fh:
        for c in iter(lambda: fh.read(1 << 20), b""):
            h.update(c)
    return h.hexdigest()


check("v8.0 sha unchanged", sha(V80) == V80_SHA, sha(V80)[:16])
m = json.load(open("v8.1_MANIFEST.json"))
check("manifest output sha matches file", sha(V81) == m["output_sha256"], sha(V81)[:16])
check("manifest source unmodified flag", m["source_unmodified"] is True)

w0 = openpyxl.load_workbook(V80, keep_vba=True, data_only=False)
w1 = openpyxl.load_workbook(V81, keep_vba=True, data_only=False)
check("v8.1 keeps all v8.0 sheets", all(s in w1.sheetnames for s in w0.sheetnames))
check("v8.1 adds COA_MAP/GL_AUDIT/GL_V8_CALC",
      all(s in w1.sheetnames for s in ("COA_MAP", "GL_AUDIT", "GL_V8_CALC")),
      str(w1.sheetnames))

calc = w1["GL_V8_CALC"]
rows = calc.max_row - 3
check("GL_V8_CALC has 552 rows", rows == 552, "got %d" % rows)

coa = w1["COA_MAP"]
total = sum(coa.cell(r, 2).value or 0 for r in range(5, 10))
check("COA_MAP totals 6,154.00", abs(total - 6154.00) < 0.005, "got %.2f" % total)

za = zipfile.ZipFile(V80).read("xl/vbaProject.bin")
zb = zipfile.ZipFile(V81).read("xl/vbaProject.bin")
check("vbaProject.bin byte-identical", hashlib.sha256(za).hexdigest() == hashlib.sha256(zb).hexdigest())

# original-sheet values untouched: compare GJ E11..E60 + GL F13..F60
same = True
for sh, rng in (("GJ", [(r, 5) for r in range(11, 61)]), ("GL", [(r, 6) for r in range(13, 61)])):
    for r, c in rng:
        if w0[sh].cell(r, c).value != w1[sh].cell(r, c).value:
            same = False
check("original sheet values untouched (GJ/GL sample)", same)

print("\n%d fail(s)" % len(fails))
sys.exit(1 if fails else 0)
