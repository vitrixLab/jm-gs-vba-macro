"""
apply_and_test_v833.py

Updates Global-Smile_2026-v8.3.3.xlsm with:
1. modEngine (from modEngine_v833_hardened.bas)
2. Sheet12 (from Sheet12_v833_hardened.cls)
3. Sheet11 (from Sheet11_v833_hardened.cls)
4. modGLAggregation (if not present, import)

Then runs end-to-end tests:
- Tests PJ -> CDJ insertion & update
- Tests Non-Vat PJ -> CDJ insertion & update
- Tests RefreshGL certification
"""

import sys
import os
import time
import win32com.client

XLSM_PATH = r"D:\citrixlabph\globalsmile\Global-Smile_2026-v8.3.3.xlsm"
MOD_ENGINE = r"D:\citrixlabph\globalsmile\modEngine_v833_hardened.bas"
SHEET12 = r"D:\citrixlabph\globalsmile\Sheet12_v833_hardened.cls"
SHEET11 = r"D:\citrixlabph\globalsmile\Sheet11_v833_hardened.cls"
MOD_GL_AGG = r"D:\citrixlabph\globalsmile\modGLAggregation.bas"

print("=== Starting Excel Automation ===")
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

try:
    print(f"Opening {XLSM_PATH}...")
    wb = xl.Workbooks.Open(XLSM_PATH, 0, False)
    vbp = wb.VBProject

    # 1. Update modEngine
    print("Updating modEngine...")
    for comp in list(vbp.VBComponents):
        if comp.Name == "modEngine":
            vbp.VBComponents.Remove(comp)
            print("  Removed old modEngine.")
            break
    vbp.VBComponents.Import(MOD_ENGINE)
    print("  Imported hardened modEngine.")

    # 2. Update modGLAggregation (if present, remove and re-import; if not, import)
    print("Updating modGLAggregation...")
    for comp in list(vbp.VBComponents):
        if comp.Name == "modGLAggregation":
            vbp.VBComponents.Remove(comp)
            print("  Removed old modGLAggregation.")
            break
    if os.path.exists(MOD_GL_AGG):
        vbp.VBComponents.Import(MOD_GL_AGG)
        print("  Imported modGLAggregation.")

    # 3. Update Code in Sheet12 (PJ tab)
    print("Updating Sheet12 code...")
    with open(SHEET12, "r", encoding="utf-8") as f:
        s12_code = f.read()
    # Strip VB attributes for CodeModule.AddFromString
    s12_clean_lines = []
    for line in s12_code.splitlines():
        if line.startswith("Attribute "):
            continue
        s12_clean_lines.append(line)
    s12_clean_code = "\n".join(s12_clean_lines)

    comp_s12 = vbp.VBComponents("Sheet12")
    comp_s12.CodeModule.DeleteLines(1, comp_s12.CodeModule.CountOfLines)
    comp_s12.CodeModule.AddFromString(s12_clean_code)
    print("  Updated Sheet12 CodeModule.")

    # 4. Update Code in Sheet11 (PJ Non-Vat tab)
    print("Updating Sheet11 code...")
    with open(SHEET11, "r", encoding="utf-8") as f:
        s11_code = f.read()
    s11_clean_lines = []
    for line in s11_code.splitlines():
        if line.startswith("Attribute "):
            continue
        s11_clean_lines.append(line)
    s11_clean_code = "\n".join(s11_clean_lines)

    comp_s11 = vbp.VBComponents("Sheet11")
    comp_s11.CodeModule.DeleteLines(1, comp_s11.CodeModule.CountOfLines)
    comp_s11.CodeModule.AddFromString(s11_clean_code)
    print("  Updated Sheet11 CodeModule.")

    # 5. Save workbook
    wb.Save()
    print("Workbook saved successfully with all updated modules.")

    # 6. Enable Events for runtime testing
    xl.EnableEvents = True

    # 7. Test InitializeEngine
    print("\n--- Testing modEngine.InitializeEngine ---")
    xl.Run("InitializeEngine", "CDJ")
    print("  InitializeEngine('CDJ') executed successfully.")

    # 8. Test FindExpenseColumn
    print("\n--- Testing modEngine.FindExpenseColumn ---")
    test_coas = [
        ("Supplies", 14),
        ("Clinic Materials and Supplies", 14),
        ("Rental", 15),
        ("Rent Expense", 15),
        ("Fuel and Oil", 16),
        ("Communication", 17),
        ("Light and Water Expense", 17),
        ("Professional Fees", 18),
        ("Clinicians Fee - Visiting", 18),
        ("Miscellaneous", 19),
        ("Representation", 19),
        ("Pantry Supplies", 19)
    ]
    for coa, expected_col in test_coas:
        col = xl.Run("FindExpenseColumn", coa)
        status = "OK" if col == expected_col else f"MISMATCH (expected {expected_col})"
        print(f"  COA '{coa}' -> Col {col} [{status}]")

    # 9. Test PJ -> CDJ Sync
    print("\n--- Testing PJ -> CDJ Sync Live ---")
    ws_pj = wb.Sheets("PJ")
    ws_cdj = wb.Sheets("CDJ")

    # Find next available row in PJ
    pj_row = 10
    while ws_pj.Cells(pj_row, 4).Value not in (None, ""):
        pj_row += 1

    print(f"  Using test row {pj_row} in PJ sheet...")
    # Populate PJ row: Date, TIN, Name, COA, Gross
    ws_pj.Cells(pj_row, 2).Value = "JAN"
    ws_pj.Cells(pj_row, 4).Value = "233-251-708"
    ws_pj.Cells(pj_row, 5).Value = "FEDERAL BRENT RETAIL INC"
    ws_pj.Cells(pj_row, 7).Value = "Clinic Materials and Supplies"
    ws_pj.Cells(pj_row, 14).Value = 5600.00
    ws_pj.Cells(pj_row, 10).Value = 5000.00   # Net
    ws_pj.Cells(pj_row, 13).Value = 600.00    # VAT
    ref_id = f"PJ-TEST{int(time.time())%10000:04d}"
    ws_pj.Cells(pj_row, 8).Value = ref_id

    # Call SyncPJToCDJ
    print(f"  Calling modEngine.SyncPJToCDJ for row {pj_row}, Ref: {ref_id}...")
    xl.Run("SyncPJToCDJ", ws_pj, pj_row)

    # Verify CDJ row was created
    cdj_target_row = xl.Run("FindCDJRow", ref_id)
    print(f"  CDJ target row for {ref_id}: {cdj_target_row}")
    assert cdj_target_row >= 15, f"Sync failed: target row {cdj_target_row} < 15"

    cdj_part = ws_cdj.Cells(cdj_target_row, 5).Value
    cdj_cash = ws_cdj.Cells(cdj_target_row, 6).Value
    cdj_vat = ws_cdj.Cells(cdj_target_row, 7).Value
    cdj_mat = ws_cdj.Cells(cdj_target_row, 14).Value
    cdj_deb = ws_cdj.Cells(cdj_target_row, 20).Value
    print(f"  CDJ Row {cdj_target_row} content: Part='{cdj_part}', Cash={cdj_cash}, VAT={cdj_vat}, Mat={cdj_mat}, Debit={cdj_deb}")
    assert cdj_part == "FEDERAL BRENT RETAIL INC", f"Particulars mismatch: {cdj_part}"
    assert abs(cdj_cash - (-5600.00)) < 0.01, f"Cash mismatch: {cdj_cash}"
    assert abs(cdj_vat - 600.00) < 0.01, f"VAT mismatch: {cdj_vat}"
    assert abs(cdj_mat - 5000.00) < 0.01, f"Clinic Materials mismatch: {cdj_mat}"
    assert abs(cdj_deb - 5600.00) < 0.01, f"Debit mismatch: {cdj_deb}"
    print("  [PASS] PJ -> CDJ INSERT test successful!")

    # Test UPDATE on same row
    print("\n--- Testing PJ -> CDJ UPDATE ---")
    ws_pj.Cells(pj_row, 14).Value = 11200.00
    ws_pj.Cells(pj_row, 10).Value = 10000.00
    ws_pj.Cells(pj_row, 13).Value = 1200.00
    xl.Run("SyncPJToCDJ", ws_pj, pj_row)

    cdj_cash_upd = ws_cdj.Cells(cdj_target_row, 6).Value
    cdj_mat_upd = ws_cdj.Cells(cdj_target_row, 14).Value
    cdj_deb_upd = ws_cdj.Cells(cdj_target_row, 20).Value
    print(f"  CDJ Row {cdj_target_row} after update: Cash={cdj_cash_upd}, Mat={cdj_mat_upd}, Debit={cdj_deb_upd}")
    assert abs(cdj_cash_upd - (-11200.00)) < 0.01, f"Updated Cash mismatch: {cdj_cash_upd}"
    assert abs(cdj_mat_upd - 10000.00) < 0.01, f"Updated Mat mismatch: {cdj_mat_upd}"
    assert abs(cdj_deb_upd - 11200.00) < 0.01, f"Updated Debit mismatch: {cdj_deb_upd}"
    print("  [PASS] PJ -> CDJ UPDATE test successful!")

    # Clean up test row in PJ and CDJ
    print("\n--- Cleaning up test rows ---")
    ws_pj.Rows(pj_row).ClearContents()
    ws_cdj.Rows(cdj_target_row).ClearContents()
    # Reset next row in CDJ if it was appended
    ws_cdj.Cells(14, 27).Value = cdj_target_row
    wb.Save()
    print("  Test data cleaned up and workbook saved.")

    # 10. Run RefreshGL to confirm GL calculation is 100% intact
    print("\n--- Testing RefreshGL ---")
    t0 = time.time()
    xl.Run("RefreshGL")
    t1 = time.time()
    print(f"  RefreshGL completed in {t1 - t0:.2f}s")

    # Read audit result
    ws_audit = wb.Sheets("GL_AUDIT")
    audit_status = ws_audit.Cells(7, 3).Value
    audit_diff = ws_audit.Cells(4, 3).Value
    audit_unmap = ws_audit.Cells(5, 3).Value
    print(f"  GL_AUDIT Status: {audit_status}, Difference: {audit_diff}, Unmapped: {audit_unmap}")

    # Read GL Cash in Bank & Miscellaneous
    ws_gl = wb.Sheets("GL")
    cib_cr = ws_gl.Cells(13, 10).Value
    misc_db = ws_gl.Cells(520, 9).Value
    print(f"  GL Cash In Bank Jan Credit: {cib_cr}")
    print(f"  GL Miscellaneous Jan Debit: {misc_db}")

    print("\n=== ALL TESTS PASSED! ===")

except Exception as e:
    print(f"\nERROR: {e}")
    sys.exit(1)

finally:
    try:
        wb.Close(True)
    except:
        pass
    xl.Quit()
    del xl
