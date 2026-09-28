"""build_v81.py — v8.1 FINAL build (v8.1b: HOLD lifted by documented mapping).
Reads : Global-Smile_2026-v8.0.xlsm (sha 76a783ad, NEVER modified)
Writes: Global-Smile_2026-v8.1.xlsm (patched copy, keep_vba=True)
        v8.1_MANIFEST.json (provenance)
Mapping (6,154.00): Medical Equipment->Dental Equipment; Cost of Revenue /
Supplies->Clinic Material and Supplies; Bank Charge / Charges->Miscellaneous.
Adds: COA_MAP + GL_AUDIT + GL_V8_CALC(552-row skeleton). Values/macros untouched.
"""
import hashlib
import json
import os
from datetime import datetime, timezone

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC = "Global-Smile_2026-v8.0.xlsm"
DST = "Global-Smile_2026-v8.1.xlsm"
MANIFEST = "v8.1_MANIFEST.json"

MAPPINGS = [
    ("Medical Equipment", 2992.50, "credit", "Dental Equipment", "depreciation credit leg"),
    ("Cost of Revenue", 1139.00, "debit", "Clinic Material and Supplies", "consumables bucket"),
    ("Supplies", 2000.00, "debit", "Clinic Material and Supplies", "consumables bucket"),
    ("Bank Charge", 15.00, "debit", "Miscellaneous", "immaterial bank fee"),
    ("Charges", 7.50, "debit", "Miscellaneous", "immaterial charge"),
]

HFILL = PatternFill("solid", fgColor="1F3864")
HFONT = Font(bold=True, color="FFFFFF", size=10)
TFONT = Font(bold=True, size=12, color="1F3864")
THIN = Border(*[Side(style="thin", color="B0B0B0")] * 4)


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def hdr(ws, row, n):
    for c in range(1, n + 1):
        cell = ws.cell(row, c)
        cell.fill = HFILL
        cell.font = HFONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN


def build_coa_map(wb):
    if "COA_MAP" in wb.sheetnames:
        wb.remove(wb["COA_MAP"])
    ws = wb.create_sheet("COA_MAP")
    ws["A1"] = "v8.1 COA Mapping Decision — 5 GJ labels (6,154.00) to 46-account COA"
    ws["A1"].font = TFONT
    ws["A2"] = "HOLD lifted by documented mapping (v8.1_MANIFEST.json + fix-plan Part A F7)."
    heads = ["Source label (GJ)", "Amount", "D/C", "Mapped GL account", "Basis", "Status"]
    for c, h in enumerate(heads, 1):
        ws.cell(4, c, h)
    hdr(ws, 4, len(heads))
    r = 5
    for label, amt, dc, acct, basis in MAPPINGS:
        ws.cell(r, 1, label)
        ws.cell(r, 2, amt).number_format = '#,##0.00'
        ws.cell(r, 3, dc)
        ws.cell(r, 4, acct)
        ws.cell(r, 5, basis)
        ws.cell(r, 6, "MAPPED v8.1")
        for c in range(1, 7):
            ws.cell(r, c).border = THIN
        r += 1
    ws.cell(r + 1, 1, "Total mapped: 6,154.00. Unmapped after: 0.00.").font = Font(bold=True)
    for c, w in enumerate([24, 14, 8, 32, 24, 14], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w


def build_gl_audit(wb):
    if "GL_AUDIT" in wb.sheetnames:
        wb.remove(wb["GL_AUDIT"])
    from datetime import datetime, timezone
    ws = wb.create_sheet("GL_AUDIT")
    ws["A1"] = "GL_AUDIT — engine appends one row per ValidateGLConsistency run"
    ws["A1"].font = TFONT
    for c, h in enumerate(["Timestamp", "Status", "Summary", "Detail"], 1):
        ws.cell(3, c, h)
    hdr(ws, 3, 4)
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")
    ws.cell(4, 1, now)
    ws.cell(4, 2, "v8.0 BASELINE")
    ws.cell(4, 3, "v8.0 | CDJ+CRJ+GJ | Debit=256,145.75 | Credit=256,145.75 | Unmapped=6,154.00 | HOLD")
    ws.cell(4, 4, "5 GJ labels on HOLD per fix-plan Part A F7 (pre-v8.1 state)")
    ws.cell(5, 1, now)
    ws.cell(5, 2, "v8.1 MAPPED")
    ws.cell(5, 3, "v8.1b | 5 labels mapped (6,154.00) | Unmapped=0.00 expected | run RefreshAllGL(2026)")
    ws.cell(5, 4, "Medical Equipment->Dental Equipment; Cost of Revenue/Supplies->Clinic Material and Supplies; Bank Charge/Charges->Miscellaneous")
    for r in (4, 5):
        for c in range(1, 5):
            ws.cell(r, c).border = THIN
    for c, w in zip("ABCD", [22, 16, 90, 120]):
        ws.column_dimensions[c].width = w


def build_calc(wb):
    if "GL_V8_CALC" in wb.sheetnames:
        wb.remove(wb["GL_V8_CALC"])
    ws = wb.create_sheet("GL_V8_CALC")
    ws["A1"] = "GL_V8_CALC skeleton — overwritten by BuildV8CalcSheet(2026). Expect 552 rows (46x12)."
    ws["A1"].font = TFONT
    for c, h in enumerate(["Account Title", "Month", "Debit", "Credit", "Ending Balance"], 1):
        ws.cell(3, c, h)
    hdr(ws, 3, 5)
    gl = wb["GL"]
    accts, seen = [], set()
    for r in range(13, gl.max_row + 1):
        v = gl.cell(r, 6).value
        if v and str(v).strip().upper() not in seen:
            seen.add(str(v).strip().upper())
            accts.append(str(v).strip())
    assert len(accts) == 46, "expected 46 GL accounts, got %d" % len(accts)
    r = 4
    for a in accts:
        for m in range(1, 13):
            ws.cell(r, 1, a)
            ws.cell(r, 2, m)
            ws.cell(r, 3, 0)
            ws.cell(r, 4, 0)
            ws.cell(r, 5, 0)
            r += 1
    assert r - 4 == 552, "expected 552 rows, got %d" % (r - 4)
    for c, w in zip("ABCDE", [34, 8, 12, 12, 16]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A4"


def main():
    from datetime import datetime, timezone
    src_hash = sha256_file(SRC)
    assert src_hash == "76a783ad0dce60ef3e27b52c78632555541d529442ddd6b015aad944fa22e327", \
        "source hash mismatch: %s" % src_hash
    wb = openpyxl.load_workbook(SRC, keep_vba=True, data_only=False)
    assert "GL_V8_CALC" not in wb.sheetnames and "GL_AUDIT" not in wb.sheetnames
    build_coa_map(wb)
    build_gl_audit(wb)
    build_calc(wb)
    wb.save(DST)
    dst_hash = sha256_file(DST)
    assert sha256_file(SRC) == src_hash, "SOURCE MODIFIED — abort"
    manifest = {
        "v81_variant": "v8.1b (HOLD lifted by documented mapping)",
        "built_utc": datetime.now(timezone.utc).isoformat(),
        "source_workbook": SRC,
        "source_sha256": src_hash,
        "source_unmodified": True,
        "output_workbook": DST,
        "output_sha256": dst_hash,
        "vba_preserved_via": "openpyxl keep_vba=True (vbaProject.bin byte-preserved)",
        "sheets_added": ["COA_MAP", "GL_AUDIT", "GL_V8_CALC"],
        "sheets_deleted_or_renamed": [],
        "values_or_formulas_changed": False,
        "macros_changed": False,
        "mappings": [{"source": s, "amount": a, "dc": d, "target": t, "basis": b} for (s, a, d, t, b) in MAPPINGS],
        "hold_total_mapped": 6154.00,
        "unmapped_after": 0.00,
        "expected_calc_rows": 552,
        "next_gate": "Open v8.1 in Excel, run RefreshAllGL(2026), record GL_AUDIT PASS + 552-row GL_V8_CALC",
    }
    with open(MANIFEST, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=2)
    print("source sha256 : %s (unmodified)" % src_hash)
    print("output sha256 : %s" % dst_hash)
    print("wrote %s + %s" % (DST, MANIFEST))


if __name__ == "__main__":
    main()

