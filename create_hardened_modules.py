"""
create_hardened_modules.py

Generates:
1. modEngine_v833_hardened.bas
2. Sheet12_v833_hardened.cls (PJ tab)
3. Sheet11_v833_hardened.cls (PJ Non-Vat tab)
4. modGLAggregation.bas (ensure it exists and is clean)
"""

import os
from oletools.olevba import VBA_Parser

# Extract current modules from xlsm
vp = VBA_Parser('Global-Smile_2026-v8.3.3.xlsm')
macros = {}
for (_, _, fname, code) in vp.extract_macros():
    macros[fname] = code
vp.close()

print(f"Extracted {len(macros)} modules from xlsm.")

# ==============================================================================
# 1. HARDEN modEngine.bas
# ==============================================================================
mod_engine = macros.get('modEngine.bas', '')
# Normalize newlines
lines = [l.strip('\r\n') for l in mod_engine.replace('\r\r\n', '\n').split('\n')]
text = '\n'.join(lines)

# 1.1 Fix InitializeEngine default
old_init = """Public Sub InitializeEngine(Optional targetSheet As String = "CDJ")
    On Error Resume Next
    gSyncSheet = targetSheet"""

new_init = """Public Sub InitializeEngine(Optional targetSheet As String = "CDJ")
    On Error Resume Next
    If Trim$(targetSheet) = "" Then targetSheet = "CDJ"
    gSyncSheet = targetSheet"""

if old_init in text:
    text = text.replace(old_init, new_init, 1)
    print("[modEngine] Fixed InitializeEngine default handling.")
else:
    print("[modEngine] WARNING: old_init not found")

# 1.2 Fix all calls InitializeEngine(gSyncSheet) -> InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
text = text.replace('Call InitializeEngine(gSyncSheet)', 'Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))')
print("[modEngine] Replaced all InitializeEngine(gSyncSheet) calls with fallback.")

# 1.3 Fix FindExpenseColumn
old_fec = """Public Function FindExpenseColumn(description As String) As Integer
    Dim descLower As String, i As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    descLower = LCase(Trim(description))
    If descLower = "" Then FindExpenseColumn = 0: Exit Function
    For i = 1 To gHeaderCount
        If arrHeaders(i) <> "" Then
            If InStr(descLower, arrHeaders(i)) > 0 Or InStr(arrHeaders(i), descLower) > 0 Then
                FindExpenseColumn = arrCols(i)
                Exit Function
            End If
        End If
    Next i
    FindExpenseColumn = 0
End Function"""

new_fec = """Public Function FindExpenseColumn(description As String) As Integer
    Dim descLower As String, i As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    descLower = LCase(Trim(description))
    If descLower = "" Then FindExpenseColumn = 0: Exit Function
    
    ' --- Pass 1: live CDJ header substring match ---
    For i = 1 To gHeaderCount
        If arrHeaders(i) <> "" Then
            If InStr(descLower, arrHeaders(i)) > 0 Or InStr(arrHeaders(i), descLower) > 0 Then
                FindExpenseColumn = arrCols(i)
                Exit Function
            End If
        End If
    Next i
    
    ' --- Pass 2: COA alias mapping to CDJ columns (14-19) ---
    Select Case True
        Case InStr(descLower, "clinic") > 0 Or InStr(descLower, "medical") > 0 _
          Or (InStr(descLower, "suppl") > 0 And InStr(descLower, "pantry") = 0)
            FindExpenseColumn = CDJ_EXPENSE_START       ' Col 14: Clinic Materials & Supplies
            
        Case InStr(descLower, "rent") > 0 Or InStr(descLower, "lease") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 1   ' Col 15: Rent Expense
            
        Case InStr(descLower, "fuel") > 0 Or InStr(descLower, "gas") > 0 _
          Or InStr(descLower, "oil") > 0 Or InStr(descLower, "toll") > 0 _
          Or InStr(descLower, "transport") > 0 Or InStr(descLower, "travel") > 0 _
          Or InStr(descLower, "delivery") > 0 Or InStr(descLower, "parking") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 2   ' Col 16: Gas, Oil, Fuel, Toll, Parking
            
        Case InStr(descLower, "communication") > 0 Or InStr(descLower, "light") > 0 _
          Or InStr(descLower, "water") > 0 Or InStr(descLower, "utilit") > 0 _
          Or InStr(descLower, "electric") > 0 Or InStr(descLower, "internet") > 0 _
          Or InStr(descLower, "power") > 0 Or InStr(descLower, "phone") > 0 Or InStr(descLower, "tel") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 3   ' Col 17: Communication, Light and Water
            
        Case InStr(descLower, "professional") > 0 Or InStr(descLower, "consult") > 0 _
          Or InStr(descLower, "clinician") > 0 Or InStr(descLower, "doctor") > 0 _
          Or InStr(descLower, "legal") > 0 Or InStr(descLower, "audit") > 0 _
          Or InStr(descLower, "accounting") > 0 Or InStr(descLower, "retainer") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 4   ' Col 18: Professional Fees
            
        Case Else
            ' All other COAs (Miscellaneous, Representation, Pantry, Taxes, Repairs, Dues, etc.) -> Sundry
            FindExpenseColumn = CDJ_EXPENSE_END         ' Col 19: Sundry Account
    End Select
End Function"""

if old_fec in text:
    text = text.replace(old_fec, new_fec, 1)
    print("[modEngine] Replaced FindExpenseColumn with comprehensive COA alias mapping.")
else:
    print("[modEngine] WARNING: old_fec not found")

# 1.4 Fix FindCDJRow
old_fcr = """Public Function FindCDJRow(ref As String) As Long
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    ref = Trim(ref)
    If ref = "" Then FindCDJRow = 0: Exit Function
    If dictRef.Exists(ref) Then FindCDJRow = dictRef(ref) Else FindCDJRow = 0
End Function"""

new_fcr = """Public Function FindCDJRow(ref As String) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    ref = Trim$(ref)
    If ref = "" Then FindCDJRow = 0: Exit Function
    If dictRef.Exists(ref) Then
        FindCDJRow = dictRef(ref)
        Exit Function
    End If
    ' Fallback scan of CDJ sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
        If lastRow >= CDJ_DATA_START Then
            For i = CDJ_DATA_START To lastRow
                If Trim$(CStr(ws.Cells(i, CDJ_COL_REF).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_PNV).Value)) = ref Then
                    dictRef(ref) = i
                    FindCDJRow = i
                    Exit Function
                End If
            Next i
        End If
    End If
    FindCDJRow = 0
End Function"""

if old_fcr in text:
    text = text.replace(old_fcr, new_fcr, 1)
    print("[modEngine] Added CDJ sheet fallback scan to FindCDJRow.")
else:
    print("[modEngine] WARNING: old_fcr not found")

# 1.5 Fix SyncPJToCDJ
old_sync_pj = """Public Sub SyncPJToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim grossAmount As Double, netAmount As Double, vatAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value)
    If ref = "" Or ref = "TOTAL" Then Exit Sub
    supplierName = Trim(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value)
    description = Trim(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value)
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    grossAmount = sourceSheet.Cells(sourceRow, PJ_COL_GROSS).Value
    netAmount = sourceSheet.Cells(sourceRow, PJ_COL_NET).Value
    vatAmount = sourceSheet.Cells(sourceRow, PJ_COL_VAT).Value
    If grossAmount = 0 Or Not IsNumeric(grossAmount) Then Exit Sub
    expenseCol = FindExpenseColumn(description)
    If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
        whtax = CalculateWithholdingTax( _
            Trim(Trim(wsTarget.Cells(13, expenseCol).Value) & " " & _
                 Trim(wsTarget.Cells(14, expenseCol).Value)), netAmount)
    Else
        whtax = 0
    End If
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = grossAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated from PJ Row " & sourceRow & " | " & Now
        End With
    Else
        targetRow = gNextRow
        If targetRow < CDJ_DATA_START Then
            targetRow = CDJ_DATA_START
            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).Value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = grossAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted from PJ Row " & sourceRow & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub"""

new_sync_pj = """Public Sub SyncPJToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim grossAmount As Double, netAmount As Double, vatAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value))
    If ref = "" Or UCase$(ref) = "TOTAL" Then Exit Sub
    supplierName = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value))
    description = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value))
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_GROSS).Value) Then
        grossAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_GROSS).Value)
    Else
        grossAmount = 0
    End If
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_NET).Value) Then
        netAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_NET).Value)
    Else
        netAmount = 0
    End If
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_VAT).Value) Then
        vatAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_VAT).Value)
    Else
        vatAmount = 0
    End If
    If grossAmount = 0 Then Exit Sub
    expenseCol = FindExpenseColumn(description)
    If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
        whtax = CalculateWithholdingTax(description, netAmount)
    Else
        whtax = 0
    End If
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = grossAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated from PJ Row " & sourceRow & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
    Else
        targetRow = gNextRow
        If targetRow < CDJ_DATA_START Then
            targetRow = CDJ_DATA_START
            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).Value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = grossAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = grossAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted from PJ Row " & sourceRow & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub"""

if old_sync_pj in text:
    text = text.replace(old_sync_pj, new_sync_pj, 1)
    print("[modEngine] Hardened SyncPJToCDJ.")
else:
    print("[modEngine] WARNING: old_sync_pj not found")

# 1.6 Fix SyncPJNonVatToCDJ
old_sync_pnv = """Public Sub SyncPJNonVatToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim netAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value)
    If ref = "" Or ref = "TOTAL" Then Exit Sub
    supplierName = Trim(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value)
    description = Trim(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value)
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    netAmount = sourceSheet.Cells(sourceRow, 14).Value
    If netAmount = 0 Or Not IsNumeric(netAmount) Then Exit Sub
    whtax = CalculateEWT(description, netAmount)
    expenseCol = FindExpenseColumn(description)
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = netAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 2).Value
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c
            .Cells(targetRow, expenseCol).Value = netAmount
            .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
        End With
    Else
        targetRow = gNextRow
        If targetRow < CDJ_DATA_START Then
            targetRow = CDJ_DATA_START
            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).Value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = netAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 2).Value
            .Cells(targetRow, expenseCol).Value = netAmount
            .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub"""

new_sync_pnv = """Public Sub SyncPJNonVatToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim netAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value))
    If ref = "" Or UCase$(ref) = "TOTAL" Then Exit Sub
    supplierName = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value))
    description = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value))
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    If IsNumeric(sourceSheet.Cells(sourceRow, 14).Value) Then
        netAmount = CDbl(sourceSheet.Cells(sourceRow, 14).Value)
    Else
        netAmount = 0
    End If
    If netAmount = 0 Then Exit Sub
    whtax = CalculateEWT(description, netAmount)
    expenseCol = FindExpenseColumn(description)
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = netAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 2).Value
            For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
    Else
        targetRow = gNextRow
        If targetRow < CDJ_DATA_START Then
            targetRow = CDJ_DATA_START
            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).Value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = netAmount * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 2).Value
            If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount + whtax
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub"""

if old_sync_pnv in text:
    text = text.replace(old_sync_pnv, new_sync_pnv, 1)
    print("[modEngine] Hardened SyncPJNonVatToCDJ.")
else:
    print("[modEngine] WARNING: old_sync_pnv not found")

with open('modEngine_v833_hardened.bas', 'w', encoding='utf-8') as f:
    f.write(text)

# ==============================================================================
# 2. HARDEN Sheet12.cls (Tab "PJ")
# ==============================================================================
sheet12 = macros.get('Sheet12.cls', '')
s12_lines = [l.strip('\r\n') for l in sheet12.replace('\r\r\n', '\n').split('\n')]
s12_text = '\n'.join(s12_lines)

# Fix Worksheet_Change in Sheet12 to include Column B (Date) and sync on update as well as insert
old_s12_wc = """    ' ---- PART 1: LOOKUP, VAT, ID STAMPING ----
    Set rngLookup = Intersect(Target, Union(Me.Range("D:D"), Me.Range("E:E"), _
                                            Me.Range("F:F"), Me.Range("G:G"), _
                                            Me.Range("N:N")))
    If Not rngLookup Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupLookup
        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")

        For Each cell In rngLookup
            If cell.Row >= START_ROW Then
                If cell.Column = COL_TIN Or cell.Column = COL_NAME Then
                    Dim tinVal As String, nameVal As String
                    tinVal = Trim(Me.Cells(cell.Row, COL_TIN).Value)
                    nameVal = Trim(Me.Cells(cell.Row, COL_NAME).Value)
                    If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then
                        Call CalculateTotals(cell.Row)
                        GoTo NextLookup
                    End If
                    If tinVal <> "" Then
                        Call LookupByTIN(cell.Row, tinVal)
                    ElseIf nameVal <> "" Then
                        Call LookupByName(cell.Row, nameVal)
                    Else
                        Call ClearSupplierRow(cell.Row)
                    End If
                End If

                If Me.Cells(cell.Row, COL_GROSS).Value <> "" And _
                   Me.Cells(cell.Row, COL_TIN).Value <> "TOTAL" And _
                   Me.Cells(cell.Row, COL_NAME).Value <> "TOTAL" Then
                    Me.Cells(cell.Row, COL_DEBIT).Formula = "=N" & cell.Row & "/1.12"
                    Me.Cells(cell.Row, COL_CREDIT).Formula = "=N" & cell.Row & "-J" & cell.Row
                    Me.Cells(cell.Row, COL_DEBIT_MIRROR).Formula = "=J" & cell.Row
                    Me.Cells(cell.Row, COL_CREDIT_MIRROR).Formula = "=M" & cell.Row

                    ' ---- ID STAMPING ----
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" And _
                       Me.Cells(cell.Row, COL_REF).Value = "" Then
                        newRef = modEngine.GetNextPJNumber()
                        Me.Cells(cell.Row, COL_REF).Value = newRef
                        Me.Cells(cell.Row, COL_FULL_ID).Value = "PJINV-" & currentTimestamp & "-" & Right(newRef, 5)

                        ' ====== FIX: SYNC IMMEDIATELY ======
                        Call modEngine.SyncPJToCDJ(Me, cell.Row)
                    End If
                End If
            End If
NextLookup:
        Next cell

CleanupLookup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If"""

new_s12_wc = """    ' ---- PART 1: LOOKUP, VAT, ID STAMPING AND SYNC ----
    Set rngLookup = Intersect(Target, Union(Me.Range("B:B"), Me.Range("D:D"), Me.Range("E:E"), _
                                            Me.Range("F:F"), Me.Range("G:G"), _
                                            Me.Range("N:N")))
    If Not rngLookup Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupLookup
        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")

        For Each cell In rngLookup
            If cell.Row >= START_ROW Then
                If cell.Column = COL_TIN Or cell.Column = COL_NAME Then
                    Dim tinVal As String, nameVal As String
                    tinVal = Trim(Me.Cells(cell.Row, COL_TIN).Value)
                    nameVal = Trim(Me.Cells(cell.Row, COL_NAME).Value)
                    If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then
                        Call CalculateTotals(cell.Row)
                        GoTo NextLookup
                    End If
                    If tinVal <> "" Then
                        Call LookupByTIN(cell.Row, tinVal)
                    ElseIf nameVal <> "" Then
                        Call LookupByName(cell.Row, nameVal)
                    Else
                        Call ClearSupplierRow(cell.Row)
                    End If
                End If

                If Me.Cells(cell.Row, COL_GROSS).Value <> "" And _
                   Me.Cells(cell.Row, COL_TIN).Value <> "TOTAL" And _
                   Me.Cells(cell.Row, COL_NAME).Value <> "TOTAL" Then
                    Me.Cells(cell.Row, COL_DEBIT).Formula = "=N" & cell.Row & "/1.12"
                    Me.Cells(cell.Row, COL_CREDIT).Formula = "=N" & cell.Row & "-J" & cell.Row
                    Me.Cells(cell.Row, COL_DEBIT_MIRROR).Formula = "=J" & cell.Row
                    Me.Cells(cell.Row, COL_CREDIT_MIRROR).Formula = "=M" & cell.Row

                    ' ---- ID STAMPING AND IMMEDIATE SYNC ----
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" Then
                        If Me.Cells(cell.Row, COL_REF).Value = "" Then
                            newRef = modEngine.GetNextPJNumber()
                            Me.Cells(cell.Row, COL_REF).Value = newRef
                            Me.Cells(cell.Row, COL_FULL_ID).Value = "PJINV-" & currentTimestamp & "-" & Right(newRef, 5)
                        End If

                        ' ====== SYNC TO CDJ (New or Updated) ======
                        Call modEngine.SyncPJToCDJ(Me, cell.Row)
                    End If
                End If
            End If
NextLookup:
        Next cell

CleanupLookup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If"""

if old_s12_wc in s12_text:
    s12_text = s12_text.replace(old_s12_wc, new_s12_wc, 1)
    print("[Sheet12] Updated Worksheet_Change to include Column B and support continuous updates.")
else:
    print("[Sheet12] WARNING: old_s12_wc not found")

with open('Sheet12_v833_hardened.cls', 'w', encoding='utf-8') as f:
    f.write(s12_text)

# ==============================================================================
# 3. HARDEN Sheet11.cls (Tab "PJ Non-Vat")
# ==============================================================================
sheet11 = macros.get('Sheet11.cls', '')
s11_lines = [l.strip('\r\n') for l in sheet11.replace('\r\r\n', '\n').split('\n')]
s11_text = '\n'.join(s11_lines)

old_s11_wc = """    ' ---- PART 1: LOOKUP, AMOUNT (NO VAT), ID STAMPING ----
    Set rngLookup = Intersect(Target, Union(Me.Range("D:D"), Me.Range("E:E"), _
                                            Me.Range("F:F"), Me.Range("G:G"), _
                                            Me.Range("N:N")))    ' Column N = Net Amount
    If Not rngLookup Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupLookup
        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")

        For Each cell In rngLookup
            If cell.Row >= START_ROW Then
                If cell.Column = COL_TIN Or cell.Column = COL_NAME Then
                    Dim tinVal As String, nameVal As String
                    tinVal = Trim(Me.Cells(cell.Row, COL_TIN).Value)
                    nameVal = Trim(Me.Cells(cell.Row, COL_NAME).Value)
                    If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then
                        Call CalculateTotals(cell.Row)
                        GoTo NextLookup
                    End If
                    If tinVal <> "" Then
                        Call LookupByTIN(cell.Row, tinVal)
                    ElseIf nameVal <> "" Then
                        Call LookupByName(cell.Row, nameVal)
                    Else
                        Call ClearSupplierRow(cell.Row)
                    End If
                End If

                ' ---- NON-VAT AMOUNT (NO VAT) ----
                If Me.Cells(cell.Row, COL_GROSS).Value <> "" And _
                   Me.Cells(cell.Row, COL_TIN).Value <> "TOTAL" And _
                   Me.Cells(cell.Row, COL_NAME).Value <> "TOTAL" Then

                    ' Set J = N, M = 0
                    Me.Cells(cell.Row, COL_DEBIT).Value = Me.Cells(cell.Row, COL_GROSS).Value
                    Me.Cells(cell.Row, COL_CREDIT).Value = 0
                    Me.Cells(cell.Row, COL_DEBIT_MIRROR).Value = Me.Cells(cell.Row, COL_GROSS).Value
                    Me.Cells(cell.Row, COL_CREDIT_MIRROR).Value = 0

                    ' ---- ID STAMPING (PNV-XXXXX) ----
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" And _
                       Me.Cells(cell.Row, COL_REF).Value = "" Then
                        newRef = modEngine.GetNextPJNumber()       ' Uses sequence from AA13
                        newRef = "PNV-" & Right(newRef, 5)         ' Change prefix to PNV
                        Me.Cells(cell.Row, COL_REF).Value = newRef
                        Me.Cells(cell.Row, COL_FULL_ID).Value = "PNVINV-" & currentTimestamp & "-" & Right(newRef, 5)

                        ' ---- SYNC IMMEDIATELY ----
                        Call modEngine.SyncPJNonVatToCDJ(Me, cell.Row)
                    End If
                End If
            End If
NextLookup:
        Next cell

CleanupLookup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If"""

new_s11_wc = """    ' ---- PART 1: LOOKUP, AMOUNT (NO VAT), ID STAMPING AND SYNC ----
    Set rngLookup = Intersect(Target, Union(Me.Range("B:B"), Me.Range("D:D"), Me.Range("E:E"), _
                                            Me.Range("F:F"), Me.Range("G:G"), _
                                            Me.Range("N:N")))    ' Column N = Net Amount
    If Not rngLookup Is Nothing Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        On Error GoTo CleanupLookup
        currentTimestamp = Format(Now, "yyyymmdd-HHnnSS")

        For Each cell In rngLookup
            If cell.Row >= START_ROW Then
                If cell.Column = COL_TIN Or cell.Column = COL_NAME Then
                    Dim tinVal As String, nameVal As String
                    tinVal = Trim(Me.Cells(cell.Row, COL_TIN).Value)
                    nameVal = Trim(Me.Cells(cell.Row, COL_NAME).Value)
                    If LCase(tinVal) = "total" Or LCase(nameVal) = "total" Then
                        Call CalculateTotals(cell.Row)
                        GoTo NextLookup
                    End If
                    If tinVal <> "" Then
                        Call LookupByTIN(cell.Row, tinVal)
                    ElseIf nameVal <> "" Then
                        Call LookupByName(cell.Row, nameVal)
                    Else
                        Call ClearSupplierRow(cell.Row)
                    End If
                End If

                ' ---- NON-VAT AMOUNT (NO VAT) ----
                If Me.Cells(cell.Row, COL_GROSS).Value <> "" And _
                   Me.Cells(cell.Row, COL_TIN).Value <> "TOTAL" And _
                   Me.Cells(cell.Row, COL_NAME).Value <> "TOTAL" Then

                    ' Set J = N, M = 0
                    Me.Cells(cell.Row, COL_DEBIT).Value = Me.Cells(cell.Row, COL_GROSS).Value
                    Me.Cells(cell.Row, COL_CREDIT).Value = 0
                    Me.Cells(cell.Row, COL_DEBIT_MIRROR).Value = Me.Cells(cell.Row, COL_GROSS).Value
                    Me.Cells(cell.Row, COL_CREDIT_MIRROR).Value = 0

                    ' ---- ID STAMPING (PNV-XXXXX) AND SYNC ----
                    If Me.Cells(cell.Row, COL_TIN).Value <> "" And _
                       Me.Cells(cell.Row, COL_NAME).Value <> "" And _
                       Me.Cells(cell.Row, COL_COA).Value <> "" Then
                        If Me.Cells(cell.Row, COL_REF).Value = "" Then
                            newRef = modEngine.GetNextPJNumber()       ' Uses sequence from AA13
                            newRef = "PNV-" & Right(newRef, 5)         ' Change prefix to PNV
                            Me.Cells(cell.Row, COL_REF).Value = newRef
                            Me.Cells(cell.Row, COL_FULL_ID).Value = "PNVINV-" & currentTimestamp & "-" & Right(newRef, 5)
                        End If

                        ' ---- SYNC TO CDJ (New or Updated) ----
                        Call modEngine.SyncPJNonVatToCDJ(Me, cell.Row)
                    End If
                End If
            End If
NextLookup:
        Next cell

CleanupLookup:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If"""

if old_s11_wc in s11_text:
    s11_text = s11_text.replace(old_s11_wc, new_s11_wc, 1)
    print("[Sheet11] Updated Worksheet_Change to include Column B and support continuous updates.")
else:
    print("[Sheet11] WARNING: old_s11_wc not found")

with open('Sheet11_v833_hardened.cls', 'w', encoding='utf-8') as f:
    f.write(s11_text)

print("All hardened modules created successfully.")
