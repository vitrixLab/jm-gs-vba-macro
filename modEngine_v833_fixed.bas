Attribute VB_Name = "modEngine"
Rem Attribute VBA_ModuleType=VBAModule
'Option VBASupport 1
Option Explicit

' ============================================================
' modEngine
' Central memory-first engine for PJ?CDJ, SJ?CRJ, and GL auto-aggregation
' Version: 7.9.4 (fix: GL_AUTO_TOL moved to module top)
' ============================================================

' ============================================================
' MODULE-LEVEL CONSTANTS  (MUST be at top, before any Sub/Function)
' ============================================================
Private Const GL_AUTO_TOL As Double = 0.01

' ============================================================
' COLUMN CONSTANTS — PJ
' ============================================================
Private Const PJ_COL_TIN As Integer = 4
Private Const PJ_COL_NAME As Integer = 5
Private Const PJ_COL_ADDRESS As Integer = 6
Private Const PJ_COL_COA As Integer = 7
Private Const PJ_COL_REF As Integer = 8
Private Const PJ_COL_NET As Integer = 10
Private Const PJ_COL_VAT As Integer = 13
Private Const PJ_COL_GROSS As Integer = 14
Private Const PJ_COL_FULLID As Integer = 17

' ============================================================
' COLUMN CONSTANTS — SUPPLIERS DATA
' ============================================================
Private Const SUP_COL_TIN As Integer = 3
Private Const SUP_COL_NAME As Integer = 4
Private Const SUP_COL_ADDRESS As Integer = 5
Private Const SUP_COL_COA As Integer = 6

' ============================================================
' COLUMN CONSTANTS — CDJ
' ============================================================
Private Const CDJ_COL_DATE As Integer = 3
Private Const CDJ_COL_PARTICULARS As Integer = 5
Private Const CDJ_COL_CASH As Integer = 6
Private Const CDJ_COL_VAT As Integer = 7
Private Const CDJ_COL_WHTAX As Integer = 8
Private Const CDJ_COL_REF As Integer = 25
Private Const CDJ_COL_REF_PNV As Integer = 26
Private Const CDJ_COL_DAY As Integer = 4
Private Const CDJ_COL_NOTE As Integer = 24
Private Const CDJ_COL_SEQ As Integer = 27
Private Const CDJ_COL_NEXTROW As Integer = 27
Private Const CDJ_EXPENSE_START As Integer = 14
Private Const CDJ_EXPENSE_END As Integer = 19
Private Const CDJ_DATA_START As Integer = 15
Private Const CDJ_COL_DEBIT As Integer = 20
Private Const CDJ_ROW_SEQ As Integer = 13
Private Const CDJ_ROW_NEXTROW As Integer = 14

' ============================================================
' COLUMN CONSTANTS — CRJ
' ============================================================
Private Const CRJ_COL_DATE As Integer = 4
Private Const CRJ_COL_CUSTOMER As Integer = 6
Private Const CRJ_COL_INVOICE As Integer = 7
Private Const CRJ_COL_CASH As Integer = 8
Private Const CRJ_COL_OUTPUTVAT As Integer = 10
Private Const CRJ_COL_EXEMPT As Integer = 11
Private Const CRJ_COL_SALES As Integer = 12
Private Const CRJ_COL_NOTE As Integer = 15
Private Const CRJ_COL_REF As Integer = 16
Private Const CRJ_COL_SEQ As Integer = 17
Private Const CRJ_COL_NEXTROW As Integer = 17
Private Const CRJ_ROW_SEQ As Integer = 8
Private Const CRJ_ROW_NEXTROW As Integer = 9
Private Const CRJ_DATA_START As Integer = 10

' ============================================================
' COLUMN CONSTANTS — SJ
' ============================================================
Private Const SJ_COL_DATE As Integer = 3
Private Const SJ_COL_CUSTOMER As Integer = 6
Private Const SJ_COL_AMOUNT As Integer = 10
Private Const SJ_COL_DISCOUNT As Integer = 11
Private Const SJ_COL_OUTPUTVAT As Integer = 13
Private Const SJ_COL_NETSALES As Integer = 15
Private Const SJ_COL_INVOICE As Integer = 9
Private Const SJ_COL_REFERENCE As Integer = 21

' ============================================================
' GLOBAL STATE — CDJ / PJ
' ============================================================
Private dictTIN As Object
Private dictName As Object
Private dictRef As Object
Private arrHeaders As Variant
Private arrCols As Variant
Private gNextSeq As Long
Private gNextRow As Long
Private gInitialized As Boolean
Private gSyncSheet As String
Private gHeaderCount As Integer

' ============================================================
' GLOBAL STATE — CRJ
' ============================================================
Private dictCRJRef As Object
Private gNextCRJRow As Long
Private gNextCRJSeq As Long
Private gInitializedCRJ As Boolean

' ============================================================
' INIT
' ============================================================
Public Sub InitializeEngine(Optional targetSheet As String = "CDJ")
    On Error Resume Next
    gSyncSheet = targetSheet
    Set dictTIN = CreateObject("Scripting.Dictionary")
    Set dictName = CreateObject("Scripting.Dictionary")
    Set dictRef = CreateObject("Scripting.Dictionary")
    Call LoadSuppliers
    Call LoadCDJHeaders
    Call LoadCDJReferences
    Call LoadSequence
    Call LoadNextRow
    gInitialized = True
    On Error GoTo 0
End Sub

Public Sub InitializeCRJ()
    If gInitializedCRJ Then Exit Sub
    Set dictCRJRef = CreateObject("Scripting.Dictionary")
    Call LoadCRJReferences
    Call LoadCRJNextRow
    Call LoadCRJSequence
    gInitializedCRJ = True
End Sub

' ============================================================
' LOADERS
' ============================================================
Private Sub LoadSuppliers()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim tin As String, name As String, address As String, coa As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, SUP_COL_TIN).End(xlUp).Row
    For i = 2 To lastRow
        tin = Trim(ws.Cells(i, SUP_COL_TIN).Value)
        name = Trim(ws.Cells(i, SUP_COL_NAME).Value)
        address = Trim(ws.Cells(i, SUP_COL_ADDRESS).Value)
        coa = Trim(ws.Cells(i, SUP_COL_COA).Value)
        If tin <> "" Then dictTIN(tin) = Array(name, address, coa)
        If name <> "" Then dictName(name) = Array(tin, address, coa)
    Next i
End Sub

Private Sub LoadCDJHeaders()
    Dim ws As Worksheet
    Dim col As Integer, header As String, count As Integer
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    count = 0
    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END
        header = Trim(Trim(ws.Cells(13, col).Value) & " " & Trim(ws.Cells(14, col).Value))
        If header <> "" Then count = count + 1
    Next col
    If count = 0 Then
        arrHeaders = Array("Clinic Supplies", "Rent", "Fuel", "Utilities", "Professional", "Sundry")
        arrCols = Array(14, 15, 16, 17, 18, 19)
        gHeaderCount = 7
        Exit Sub
    End If
    ReDim arrHeaders(1 To count)
    ReDim arrCols(1 To count)
    count = 0
    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END
        header = Trim(Trim(ws.Cells(13, col).Value) & " " & Trim(ws.Cells(14, col).Value))
        If header <> "" Then
            count = count + 1
            arrHeaders(count) = LCase(header)
            arrCols(count) = col
        End If
    Next col
    gHeaderCount = count
End Sub

Private Sub LoadCDJReferences()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, ref As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
    For i = CDJ_DATA_START To lastRow
        ref = Trim(ws.Cells(i, CDJ_COL_REF).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_PNV).Value)
        If ref <> "" Then dictRef(ref) = i
    Next i
End Sub

Private Sub LoadSequence()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then gNextSeq = 0: Exit Sub
    val = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value
    If IsNumeric(val) Then gNextSeq = CLng(val) Else gNextSeq = 0
End Sub

Private Sub LoadNextRow()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then gNextRow = CDJ_DATA_START: Exit Sub
    val = ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value
    If IsNumeric(val) Then
        gNextRow = CLng(val)
    Else
        gNextRow = CDJ_DATA_START
        Do While ws.Cells(gNextRow, CDJ_COL_REF).Value <> ""
            gNextRow = gNextRow + 1
        Loop
    End If
End Sub

Private Sub LoadCRJReferences()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, inv As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    For i = CRJ_DATA_START To lastRow
        inv = Trim(ws.Cells(i, CRJ_COL_INVOICE).Value)
        If inv <> "" Then dictCRJRef(inv) = i
    Next i
End Sub

Private Sub LoadCRJNextRow()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then gNextCRJRow = CRJ_DATA_START: Exit Sub
    val = ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value
    If IsNumeric(val) Then
        gNextCRJRow = CLng(val)
    Else
        gNextCRJRow = CRJ_DATA_START
        Do While ws.Cells(gNextCRJRow, CRJ_COL_INVOICE).Value <> ""
            gNextCRJRow = gNextCRJRow + 1
        Loop
    End If
End Sub

Private Sub LoadCRJSequence()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then gNextCRJSeq = 0: Exit Sub
    val = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value
    If IsNumeric(val) Then gNextCRJSeq = CLng(val) Else gNextCRJSeq = 0
End Sub

' ============================================================
' SEQUENCE
' ============================================================
Public Function GetNextPJNumber() As String
    Dim ws As Worksheet
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    gNextSeq = gNextSeq + 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    If Not ws Is Nothing Then ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = gNextSeq
    On Error GoTo 0
    GetNextPJNumber = "PJ-" & Format(gNextSeq, "00000")
End Function

Public Function GetNextInvoiceNumber() As Long
    Dim ws As Worksheet
    If Not gInitializedCRJ Then Call InitializeCRJ
    gNextCRJSeq = gNextCRJSeq + 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    If Not ws Is Nothing Then ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = gNextCRJSeq
    On Error GoTo 0
    GetNextInvoiceNumber = gNextCRJSeq
End Function

Public Function GetNextSeq() As Long
    Dim s As String
    s = GetNextPJNumber()
    If Len(s) > 3 Then GetNextSeq = CLng(Mid$(s, 4)) Else GetNextSeq = 0
End Function

' ============================================================
' SUPPLIER LOOKUPS
' ============================================================
Public Function GetSupplierByTIN(tin As String) As Variant
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    tin = Trim(tin)
    If tin = "" Then GetSupplierByTIN = False: Exit Function
    If dictTIN.Exists(tin) Then GetSupplierByTIN = dictTIN(tin) Else GetSupplierByTIN = False
End Function

Public Function GetSupplierByName(name As String) As Variant
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    name = Trim(name)
    If name = "" Then GetSupplierByName = False: Exit Function
    If dictName.Exists(name) Then GetSupplierByName = dictName(name) Else GetSupplierByName = False
End Function

Public Sub AddSupplierToCache(tin As String, name As String, address As String, coa As String)
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    tin = Trim(tin): name = Trim(name)
    If tin <> "" Then dictTIN(tin) = Array(name, address, coa)
    If name <> "" Then dictName(name) = Array(tin, address, coa)
End Sub

' ============================================================
' EXPENSE COLUMN MAPPING
' ============================================================
Public Function FindExpenseColumn(description As String) As Integer
    Dim descLower As String, i As Integer
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
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
    ' --- Pass 2: COA alias table (PJ COA values -> CDJ column) ---
    ' CDJ column map: 14=Clinic Materials, 15=Rent, 16=Gas/Fuel, 17=Comm/Light/Water, 18=Prof Fees, 19=Sundry
    Select Case True
        Case InStr(descLower, "clinic material") > 0 Or InStr(descLower, "clinic supplies") > 0 _
          Or InStr(descLower, "pantry") > 0 Or InStr(descLower, "medical supplies") > 0
            FindExpenseColumn = CDJ_EXPENSE_START       ' col 14
        Case InStr(descLower, "rent") > 0 Or InStr(descLower, "rental") > 0 _
          Or InStr(descLower, "lease") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 1   ' col 15
        Case InStr(descLower, "fuel") > 0 Or InStr(descLower, "gas") > 0 _
          Or InStr(descLower, "oil") > 0 Or InStr(descLower, "toll") > 0 _
          Or InStr(descLower, "transport") > 0 Or InStr(descLower, "delivery") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 2   ' col 16
        Case InStr(descLower, "communication") > 0 Or InStr(descLower, "light") > 0 _
          Or InStr(descLower, "water") > 0 Or InStr(descLower, "utilities") > 0 _
          Or InStr(descLower, "electric") > 0 Or InStr(descLower, "internet") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 3   ' col 17
        Case InStr(descLower, "professional") > 0 Or InStr(descLower, "consulting") > 0 _
          Or InStr(descLower, "clinician") > 0 Or InStr(descLower, "legal") > 0 _
          Or InStr(descLower, "accounting") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 4   ' col 18
        Case Else
            ' Unmapped COA -> Sundry column (col 19) so entry is always written
            FindExpenseColumn = CDJ_EXPENSE_END         ' col 19 = Sundry
    End Select
End Function

' ============================================================
' WITHHOLDING TAX
' ============================================================
Public Function CalculateWithholdingTax(expenseType As String, amount As Double) As Double
    Dim expLower As String
    expLower = LCase(Trim(expenseType))
    Select Case True
        Case InStr(expLower, "rent") > 0
            CalculateWithholdingTax = amount * 0.05
        Case InStr(expLower, "professional") > 0 Or InStr(expLower, "consulting") > 0
            CalculateWithholdingTax = amount * 0.1
        Case InStr(expLower, "repair") > 0 Or InStr(expLower, "maintenance") > 0
            CalculateWithholdingTax = amount * 0.02
        Case Else
            CalculateWithholdingTax = 0
    End Select
End Function

' ============================================================
' REFERENCE CACHE
' ============================================================
Public Function FindCDJRow(ref As String) As Long
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    ref = Trim(ref)
    If ref = "" Then FindCDJRow = 0: Exit Function
    If dictRef.Exists(ref) Then FindCDJRow = dictRef(ref) Else FindCDJRow = 0
End Function

Public Sub UpdateCDJReference(ref As String, rowNum As Long)
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    ref = Trim(ref)
    If ref <> "" And rowNum > 0 Then dictRef(ref) = rowNum
End Sub

Public Function FindCRJRow(invoiceNum As String) As Long
    If Not gInitializedCRJ Then Call InitializeCRJ
    invoiceNum = Trim(invoiceNum)
    If invoiceNum = "" Then FindCRJRow = 0: Exit Function
    If dictCRJRef.Exists(invoiceNum) Then FindCRJRow = dictCRJRef(invoiceNum) Else FindCRJRow = 0
End Function

Public Sub UpdateCRJReference(invoiceNum As String, rowNum As Long)
    If Not gInitializedCRJ Then Call InitializeCRJ
    invoiceNum = Trim(invoiceNum)
    If invoiceNum <> "" And rowNum > 0 Then dictCRJRef(invoiceNum) = rowNum
End Sub

' ============================================================
' SYNC PJ -> CDJ
' ============================================================
Public Sub SyncPJToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim grossAmount As Double, netAmount As Double, vatAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
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
End Sub

' ============================================================
' SYNC SJ -> CRJ
' ============================================================
Public Sub SyncSJToCRJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim invoiceNum As Variant, targetRow As Long
    Dim referenceInvoiceNum As Variant
    Dim customerName As String
    Dim dateVal As Variant
    Dim amount As Double, discount As Double, outputVAT As Double, netSales As Double
    Dim cashInBank As Double
    On Error GoTo ErrorHandler
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    invoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_INVOICE).Value
    If invoiceNum = "" Or invoiceNum = "TOTAL" Then Exit Sub
    customerName = Trim(sourceSheet.Cells(sourceRow, SJ_COL_CUSTOMER).Value)
    dateVal = sourceSheet.Cells(sourceRow, SJ_COL_DATE).Value
    referenceInvoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_REFERENCE).Value
    amount = sourceSheet.Cells(sourceRow, SJ_COL_AMOUNT).Value
    discount = sourceSheet.Cells(sourceRow, SJ_COL_DISCOUNT).Value
    outputVAT = sourceSheet.Cells(sourceRow, SJ_COL_OUTPUTVAT).Value
    netSales = sourceSheet.Cells(sourceRow, SJ_COL_NETSALES).Value
    cashInBank = amount - discount
    If customerName = "" Or amount = 0 Or Not IsNumeric(amount) Then Exit Sub
    targetRow = FindCRJRow(CStr(invoiceNum))
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CRJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CRJ_COL_CUSTOMER).Value = customerName
            .Cells(targetRow, CRJ_COL_INVOICE).Value = invoiceNum
            .Cells(targetRow, CRJ_COL_REF).Value = referenceInvoiceNum
            .Cells(targetRow, CRJ_COL_CASH).Value = cashInBank
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).Value = outputVAT
            .Cells(targetRow, CRJ_COL_EXEMPT).Value = 0
            .Cells(targetRow, CRJ_COL_SALES).Value = netSales
            .Cells(targetRow, CRJ_COL_NOTE).Value = "Updated from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now
        End With
    Else
        targetRow = gNextCRJRow
        If targetRow < CRJ_DATA_START Then
            targetRow = CRJ_DATA_START
            Do While wsTarget.Cells(targetRow, CRJ_COL_INVOICE).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextCRJRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CRJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CRJ_COL_CUSTOMER).Value = customerName
            .Cells(targetRow, CRJ_COL_INVOICE).Value = invoiceNum
            .Cells(targetRow, CRJ_COL_REF).Value = referenceInvoiceNum
            .Cells(targetRow, CRJ_COL_CASH).Value = cashInBank
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).Value = outputVAT
            .Cells(targetRow, CRJ_COL_EXEMPT).Value = 0
            .Cells(targetRow, CRJ_COL_SALES).Value = netSales
            .Cells(targetRow, CRJ_COL_NOTE).Value = "Inserted from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now
            .Cells(targetRow, CRJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_EXEMPT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_SALES).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCRJReference CStr(invoiceNum), targetRow
        gNextCRJRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = gNextCRJRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub

' ============================================================
' STATUS
' ============================================================
Public Function EngineStatus() As String
    If gInitialized Then
        EngineStatus = "Initialized | Suppliers: " & dictTIN.count & _
                       " | CDJ Refs: " & dictRef.count & _
                       " | Headers: " & gHeaderCount & _
                       " | Next PJ Seq: " & gNextSeq & _
                       " | Next CDJ Row: " & gNextRow & _
                       " | Target: " & gSyncSheet
    Else
        EngineStatus = "NOT INITIALIZED"
    End If
    If gInitializedCRJ Then
        EngineStatus = EngineStatus & " || CRJ Refs: " & dictCRJRef.count & _
                       " | Next CRJ Seq: " & gNextCRJSeq & _
                       " | Next CRJ Row: " & gNextCRJRow
    End If
End Function

Public Sub ShowCDJStatus()
    Dim ws As Worksheet, msg As String, lastRow As Long, i As Long
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Target sheet '" & gSyncSheet & "' not found!", vbCritical: Exit Sub
    msg = "===== CDJ STATUS REPORT =====" & vbCrLf & vbCrLf
    msg = msg & "Target Sheet: " & gSyncSheet & vbCrLf
    msg = msg & "Data starts at row: " & CDJ_DATA_START & vbCrLf
    msg = msg & "Current row counter (Z14): " & gNextRow & vbCrLf
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If lastRow < CDJ_DATA_START Then
        msg = msg & "Last used row (with reference): None yet" & vbCrLf
    Else
        msg = msg & "Last used row: " & lastRow & vbCrLf
    End If
    msg = msg & "References in cache: " & dictRef.count & vbCrLf
    msg = msg & "Sequence (Z13): " & IIf(IsNumeric(ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value), ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value, "empty") & vbCrLf
    msg = msg & "Next PJ number: " & Format(gNextSeq + 1, "PJ-00000") & vbCrLf
    msg = msg & vbCrLf & "Expense headers (row 14, N-T):" & vbCrLf
    For i = 1 To gHeaderCount
        msg = msg & "  Col " & arrCols(i) & ": " & arrHeaders(i) & vbCrLf
    Next i
    MsgBox msg, vbInformation, "CDJ Status"
End Sub

Public Sub ShowCRJStatus()
    Dim ws As Worksheet, msg As String, lastRow As Long
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CRJ sheet not found!", vbCritical: Exit Sub
    msg = "===== CRJ STATUS REPORT =====" & vbCrLf & vbCrLf
    msg = msg & "Data starts at row: " & CRJ_DATA_START & vbCrLf
    msg = msg & "Current row counter (Q9): " & gNextCRJRow & vbCrLf
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    If lastRow < CRJ_DATA_START Then
        msg = msg & "Last used row (with invoice): None yet" & vbCrLf
    Else
        msg = msg & "Last used row: " & lastRow & vbCrLf
    End If
    msg = msg & "Invoice references in cache: " & dictCRJRef.count & vbCrLf
    msg = msg & "Invoice sequence (Q8): " & gNextCRJSeq & vbCrLf
    msg = msg & "Next invoice number will be: " & Format(gNextCRJSeq + 1, "00000") & vbCrLf
    MsgBox msg, vbInformation, "CRJ Status"
End Sub

' ============================================================
' REPAIR / RESET
' ============================================================
Public Sub RepairEngine()
    Dim ws As Worksheet, lastRow As Long, seqVal As Variant
    gSyncSheet = "CDJ"
    Call InitializeEngine(gSyncSheet)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If lastRow < CDJ_DATA_START Then lastRow = CDJ_DATA_START - 1
        gNextRow = lastRow + 1
        ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        seqVal = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value
        If IsNumeric(seqVal) Then gNextSeq = CLng(seqVal) Else gNextSeq = 0: ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = 0
        Call LoadCDJReferences
    End If
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
        If lastRow < CRJ_DATA_START Then lastRow = CRJ_DATA_START - 1
        gNextCRJRow = lastRow + 1
        ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = gNextCRJRow
        seqVal = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value
        If IsNumeric(seqVal) Then
            gNextCRJSeq = CLng(seqVal)
        Else
            gNextCRJSeq = 0
            ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = 0
        End If
        Call LoadCRJReferences
    End If
    MsgBox "Engine repaired!" & vbCrLf & vbCrLf & EngineStatus, vbInformation
End Sub

Public Sub ResetCRJ()
    Dim ws As Worksheet, lastRow As Long, confirm As VbMsgBoxResult
    confirm = MsgBox("This will DELETE ALL data in the CRJ sheet and reset counters." & vbCrLf & _
                     "Are you sure?", vbYesNo + vbExclamation, "Reset CRJ")
    If confirm <> vbYes Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CRJ sheet not found!", vbCritical: Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    If lastRow >= CRJ_DATA_START Then ws.Rows(CRJ_DATA_START & ":" & lastRow).ClearContents
    ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = 0
    gNextCRJSeq = 0
    ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = CRJ_DATA_START
    gNextCRJRow = CRJ_DATA_START
    If Not gInitializedCRJ Then Call InitializeCRJ
    dictCRJRef.RemoveAll
    gInitializedCRJ = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "CRJ has been reset." & vbCrLf & _
           "Sequence (Q8) = 0, Next row (Q9) = " & CRJ_DATA_START, vbInformation
End Sub

Public Sub ResetCDJ()
    Dim ws As Worksheet, lastRow As Long, confirm As VbMsgBoxResult
    confirm = MsgBox("This will DELETE ALL data in the CDJ sheet and reset counters." & vbCrLf & _
                     "Are you sure?", vbYesNo + vbExclamation, "Reset CDJ")
    If confirm <> vbYes Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CDJ sheet not found!", vbCritical: Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If lastRow >= CDJ_DATA_START Then ws.Rows(CDJ_DATA_START & ":" & lastRow).ClearContents
    ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = 0
    gNextSeq = 0
    ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = CDJ_DATA_START
    gNextRow = CDJ_DATA_START
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
    dictRef.RemoveAll
    Call LoadCDJReferences
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "CDJ has been reset." & vbCrLf & _
           "Sequence (Z13) = 0, Next row (Z14) = " & CDJ_DATA_START, vbInformation
End Sub

' ============================================================
' EWT (extended)
' ============================================================
Public Function CalculateEWT(expenseType As String, amount As Double) As Double
    Dim expLower As String
    expLower = LCase(Trim(expenseType))
    Select Case True
        Case InStr(expLower, "rent") > 0
            CalculateEWT = amount * 0.05
        Case InStr(expLower, "professional") > 0 Or InStr(expLower, "consulting") > 0
            CalculateEWT = amount * 0.1
        Case InStr(expLower, "repair") > 0 Or InStr(expLower, "maintenance") > 0
            CalculateEWT = amount * 0.02
        Case InStr(expLower, "supply") > 0 Or InStr(expLower, "office") > 0
            CalculateEWT = amount * 0.01
        Case Else
            CalculateEWT = 0
    End Select
End Function

' ============================================================
' SYNC PJ NON-VAT -> CDJ
' ============================================================
Public Sub SyncPJNonVatToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim netAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(gSyncSheet)
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
End Sub

' ============================================================
' GL AUTO-AGGREGATION  (no Const here — it lives at the top)
' ============================================================
Private Function GLN(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, "'", "")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    GLN = s
End Function

Private Function GLAcct(ByVal s As String) As String
    Dim k As String
    k = GLN(s)
    Select Case k
        Case "cash in bank", "cash": GLAcct = "Cash in Bank"
        Case "petty cash fund", "petty cash": GLAcct = "Petty Cash Fund"
        Case "account receivables", "accounts receivable", "account receivable": GLAcct = "Account Receivables"
        Case "advances to employee", "advances to employees": GLAcct = "Advances to Employees"
        Case "input vat", "input tax": GLAcct = "Input VAT"
        Case "excess cwt over it", "cwt/it", "cwt": GLAcct = "Excess CWT Over IT"
        Case "leasehold improvements": GLAcct = "Leasehold Improvements"
        Case "dental equipment", "medical equipment": GLAcct = "Dental Equipment"
        Case "due to clinicians": GLAcct = "Due To Clinicians"
        Case "due to clinicians- visiting", "due to clinicians visiting": GLAcct = "Due to Clinicians- Visiting"
        Case "vat payable", "output vat payable", "output vat": GLAcct = "VAT Payable"
        Case "ewt payable", "withholding tax", "withholding tax payable", "ewt": GLAcct = "EWT Payable"
        Case "government contributions (ee)", "ee government contributions": GLAcct = "Government Contributions (EE)"
        Case "government loans (ee)", "ee government loans": GLAcct = "Government Loans (EE)"
        Case "income tax payable": GLAcct = "Income Tax Payable"
        Case "paid-up capital", "paid up capital": GLAcct = "Paid-Up Capital"
        Case "retained earnings (deficit)", "retained earnings": GLAcct = "Retained Earnings (Deficit)"
        Case "sales", "exempt sales": GLAcct = "Sales"
        Case "rent", "rental": GLAcct = "Rent"
        Case "clinician's fee-corporators", "clinicians fee-corporators": GLAcct = "Clinician's Fee-Corporators"
        Case "clinician's fee-visiting", "clinicians fee-visiting": GLAcct = "Clinician's Fee-Visiting"
        Case "clinic material and supplies", "clinic materials and supplies", "clinic supplies": GLAcct = "Clinic Material and Supplies"
        Case "common use service area": GLAcct = "Common Use Service Area"
        Case "light and water expense", "light and water": GLAcct = "Light and Water Expense"
        Case "dep'n expense-leasehold improvements", "depn expense-leasehold improvements": GLAcct = "Dep'n Expense-Leasehold Improvements"
        Case "dep'n expense-dental equipment", "depn expense-dental equipment": GLAcct = "Dep'n Expense-Dental Equipment"
        Case "salaries and wages", "salary and wages": GLAcct = "Salaries and Wages"
        Case "de minimis": GLAcct = "De Minimis"
        Case "government contributions er share", "government contributions er": GLAcct = "Government Contributions ER Share"
        Case "professional fees", "professional fee": GLAcct = "Professional Fees"
        Case "repair and maintenance", "repairs and maintenance": GLAcct = "Repair and Maintenance"
        Case "pantry supplies": GLAcct = "Pantry Supplies"
        Case "office supplies", "supplies": GLAcct = "Office Supplies"
        Case "meals, gifts & entertainment expense", "meals and entertainment": GLAcct = "Meals, Gifts & Entertainment Expense"
        Case "communication": GLAcct = "Communication "
        Case "printing and duplication": GLAcct = "Printing and Duplication"
        Case "taxes and licences", "taxes and licenses": GLAcct = "Taxes and Licences"
        Case "delivery": GLAcct = "Delivery"
        Case "software subscription": GLAcct = "Software Subscription"
        Case "miscellaneous", "charges", "bank charge": GLAcct = "Miscellaneous"
        Case "transportation and travel", "transporation and travel": GLAcct = "Transportation and Travel"
        Case "advertisement": GLAcct = "Advertisement"
        Case "marketing expense": GLAcct = "Marketing Expense"
        Case "gas, oil, parking, toll fees", "fuel and oil": GLAcct = "Gas, Oil, Parking, Toll Fees"
        Case "other penalties and charges": GLAcct = "Other Penalties and Charges"
        Case "insurance": GLAcct = "Insurance"
    End Select
End Function

Private Function GLMonth(ByVal v As Variant) As Long
    Dim k As String
    k = GLN(CStr(v))
    Select Case Left$(k, 3)
        Case "jan": GLMonth = 1
        Case "feb": GLMonth = 2
        Case "mar": GLMonth = 3
        Case "apr": GLMonth = 4
        Case "may": GLMonth = 5
        Case "jun": GLMonth = 6
        Case "jul": GLMonth = 7
        Case "aug": GLMonth = 8
        Case "sep": GLMonth = 9
        Case "oct": GLMonth = 10
        Case "nov": GLMonth = 11
        Case "dec": GLMonth = 12
        Case Else: If IsDate(v) Then GLMonth = month(CDate(v))
    End Select
End Function

Private Sub GLAdd(ByRef d As Object, ByVal a As String, ByVal m As Long, ByVal dr As Double, ByVal cr As Double)
    If a = "" Or m < 1 Or m > 12 Then Exit Sub
    Dim k As String: k = a & "|" & m
    Dim x As Variant
    If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
    x(0) = x(0) + dr
    x(1) = x(1) + cr
    d(k) = x
End Sub

Private Sub GLAudit(ByVal aw As Worksheet, ByRef r As Long, ByVal src As String, ByVal detail As String, ByVal amt As Double)
    aw.Cells(r, 1).Value = src
    aw.Cells(r, 2).Value = detail
    aw.Cells(r, 3).Value = amt
    r = r + 1
End Sub

Private Sub GLScanCDJ(ByVal d As Object, ByVal aw As Worksheet, ByRef ar As Long)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("CDJ")
    Dim r As Long, m As Long, a As String, v As Double
    For r = 15 To w.Cells(w.Rows.count, 3).End(xlUp).Row
        If GLN(w.Cells(r, 5).Value) <> "total" Then
            m = GLMonth(w.Cells(r, 3).Value)
            If m > 0 Then
                If IsNumeric(w.Cells(r, 6).Value) Then
                    v = CDbl(w.Cells(r, 6).Value)
                    If v < 0 Then
                        GLAdd d, "Cash in Bank", m, 0, -v
                    Else
                        GLAdd d, "Cash in Bank", m, v, 0
                    End If
                End If
                If IsNumeric(w.Cells(r, 7).Value) Then GLAdd d, "Input VAT", m, CDbl(w.Cells(r, 7).Value), 0
                If IsNumeric(w.Cells(r, 8).Value) Then GLAdd d, "EWT Payable", m, 0, CDbl(w.Cells(r, 8).Value)
                If IsNumeric(w.Cells(r, 9).Value) Then GLAdd d, "Petty Cash Fund", m, CDbl(w.Cells(r, 9).Value), 0
                If IsNumeric(w.Cells(r, 10).Value) Then GLAdd d, "Government Contributions (EE)", m, 0, CDbl(w.Cells(r, 10).Value)
                If IsNumeric(w.Cells(r, 11).Value) Then GLAdd d, "Government Loans (EE)", m, 0, CDbl(w.Cells(r, 11).Value)
                If IsNumeric(w.Cells(r, 12).Value) Then GLAdd d, "De Minimis", m, CDbl(w.Cells(r, 12).Value), 0
                If IsNumeric(w.Cells(r, 13).Value) Then GLAdd d, "Salaries and Wages", m, CDbl(w.Cells(r, 13).Value), 0
                If IsNumeric(w.Cells(r, 14).Value) Then GLAdd d, "Clinic Material and Supplies", m, CDbl(w.Cells(r, 14).Value), 0
                If IsNumeric(w.Cells(r, 15).Value) Then GLAdd d, "Rent", m, CDbl(w.Cells(r, 15).Value), 0
                If IsNumeric(w.Cells(r, 16).Value) Then GLAdd d, "Gas, Oil, Parking, Toll Fees", m, CDbl(w.Cells(r, 16).Value), 0
                If IsNumeric(w.Cells(r, 17).Value) Then GLAdd d, "Communication ", m, CDbl(w.Cells(r, 17).Value), 0
                If IsNumeric(w.Cells(r, 18).Value) Then GLAdd d, "Professional Fees", m, CDbl(w.Cells(r, 18).Value), 0
                a = GLAcct(w.Cells(r, 19).Value)
                If a <> "" Then
                    If IsNumeric(w.Cells(r, 20).Value) Then GLAdd d, a, m, CDbl(w.Cells(r, 20).Value), 0
                    If IsNumeric(w.Cells(r, 21).Value) Then GLAdd d, a, m, 0, CDbl(w.Cells(r, 21).Value)
                ElseIf Len(Trim$(CStr(w.Cells(r, 19).Value))) > 0 Then
                    GLAudit aw, ar, "CDJ", CStr(w.Cells(r, 19).Value), 0
                End If
            End If
        End If
    Next r
End Sub

Private Sub GLScanCRJ(ByVal d As Object)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("CRJ")
    Dim ends As New Collection, r As Long, m As Long, st As Long, en As Long, rr As Long
    For r = 10 To w.Cells(w.Rows.count, 4).End(xlUp).Row
        If GLN(w.Cells(r, 6).Value) = "total" And GLMonth(w.Cells(r, 4).Value) > 0 Then ends.Add r
    Next r
    st = 10
    For m = 1 To 12
        If m <= ends.count Then en = ends(m) - 1 Else en = w.Cells(w.Rows.count, 6).End(xlUp).Row
        For rr = st To en
            If GLN(w.Cells(rr, 6).Value) <> "total" Then
                If IsNumeric(w.Cells(rr, 8).Value) Then GLAdd d, "Cash in Bank", m, CDbl(w.Cells(rr, 8).Value), 0
                If IsNumeric(w.Cells(rr, 9).Value) Then GLAdd d, "Excess CWT Over IT", m, CDbl(w.Cells(rr, 9).Value), 0
                If IsNumeric(w.Cells(rr, 10).Value) Then GLAdd d, "VAT Payable", m, 0, CDbl(w.Cells(rr, 10).Value)
                If IsNumeric(w.Cells(rr, 11).Value) Then GLAdd d, "Sales", m, 0, CDbl(w.Cells(rr, 11).Value)
                If IsNumeric(w.Cells(rr, 12).Value) Then GLAdd d, "Sales", m, 0, CDbl(w.Cells(rr, 12).Value)
            End If
        Next rr
        If m <= ends.count Then st = ends(m) + 1 Else Exit For
    Next m
End Sub

Private Sub GLScanGJ(ByVal d As Object, ByVal aw As Worksheet, ByRef ar As Long)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("GJ")
    Dim r As Long, m As Long, a As String, s As String, dr As Double, cr As Double
    For r = 10 To w.Cells(w.Rows.count, 5).End(xlUp).Row
        If GLMonth(w.Cells(r, 2).Value) > 0 Then m = GLMonth(w.Cells(r, 2).Value)
        s = Trim$(CStr(w.Cells(r, 5).Value))
        If m > 0 And s <> "" And GLN(s) <> "total" Then
            a = GLAcct(s)
            dr = 0: cr = 0
            If IsNumeric(w.Cells(r, 6).Value) Then dr = CDbl(w.Cells(r, 6).Value)
            If IsNumeric(w.Cells(r, 7).Value) Then cr = CDbl(w.Cells(r, 7).Value)
            If a <> "" Then
                GLAdd d, a, m, dr, cr
            ElseIf dr <> 0 Or cr <> 0 Then
                GLAudit aw, ar, "GJ", s, dr + cr
            End If
        End If
    Next r
End Sub

Private Function GLAccounts() As Variant
    GLAccounts = Array( _
        "Cash in Bank", "Petty Cash Fund", "Account Receivables", "Advances to Employees", _
        "Input VAT", "Excess CWT Over IT", "Leasehold Improvements", "Dental Equipment", _
        "Due To Clinicians", "Due to Clinicians- Visiting", "VAT Payable", "EWT Payable", _
        "Government Contributions (EE)", "Government Loans (EE)", "Income Tax Payable", _
        "Paid-Up Capital", "Retained Earnings (Deficit)", "Sales", "Rent", _
        "Clinician's Fee-Corporators", "Clinician's Fee-Visiting", "Clinic Material and Supplies", _
        "Common Use Service Area", "Light and Water Expense", _
        "Dep'n Expense-Leasehold Improvements", "Dep'n Expense-Dental Equipment", _
        "Salaries and Wages", "De Minimis", "Government Contributions ER Share", _
        "Professional Fees", "Repair and Maintenance", "Pantry Supplies", "Office Supplies", _
        "Meals, Gifts & Entertainment Expense", "Communication ", "Printing and Duplication", _
        "Taxes and Licences", "Delivery", "Software Subscription", "Miscellaneous", _
        "Transportation and Travel", "Advertisement", "Marketing Expense", _
        "Gas, Oil, Parking, Toll Fees", "Other Penalties and Charges", "Insurance")
End Function

' ============================================================
' REFRESH GL — public entry point
'   reportYear: 0 = auto (detect from source sheets, else system year)
'               any other value = force that year
' ============================================================
Private Sub Legacy_RefreshGL(Optional ByVal reportYear As Long = 0)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim aw As Worksheet, gl As Worksheet
    Dim ar As Long, r As Long, i As Long, m As Long
    Dim bal As Double, x As Variant, a As Variant
    Dim yr As Long

    On Error GoTo EH

    ' ---- Resolve the reporting year ----
    If reportYear <= 0 Then
        yr = DetectGLYear()
    Else
        yr = reportYear
    End If
    If yr < 1900 Or yr > 9999 Then yr = year(Date)

    ' ---- GL_AUDIT sheet ----
    On Error Resume Next
    Set aw = ThisWorkbook.Sheets("GL_AUDIT")
    On Error GoTo EH
    If aw Is Nothing Then
        Set aw = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        On Error Resume Next
        aw.name = "GL_AUDIT"
        On Error GoTo EH
    End If

    aw.Cells.Clear
    aw.Range("A1:C1").Value = Array("Source", "Unmapped Detail", "Amount")
    ar = 2

    ' ---- Scan sources ----
    GLScanCDJ d, aw, ar
    GLScanCRJ d
    GLScanGJ d, aw, ar

    ' ---- GL sheet must exist ----
    On Error Resume Next
    Set gl = ThisWorkbook.Sheets("GL")
    On Error GoTo EH
    If gl Is Nothing Then
        MsgBox "Sheet named 'GL' not found in this workbook." & vbCrLf & _
               "Existing sheets:" & vbCrLf & _
               Join(SheetNamesList(), vbCrLf), vbCritical, "RefreshGL"
        Exit Sub
    End If

    gl.Range("B12:I1000").ClearContents
    gl.Range("B11:I11").Value = Array("Date", "Day", "Reference", "Explanation", _
                                      "Account Title", "Debit", "Credit", "Ending Balance")

    ' ---- Write every account × month ----
    a = GLAccounts()
    r = 13
    For i = LBound(a) To UBound(a)
        bal = 0
        For m = 1 To 12
            gl.Cells(r, 2).Value = MonthName(m, True)
            gl.Cells(r, 3).Value = Day(DateSerial(yr, m + 1, 0))
            gl.Cells(r, 5).Value = "Ending Balance"
            gl.Cells(r, 6).Value = a(i)
            x = Array(0#, 0#)
            If d.Exists(CStr(a(i)) & "|" & m) Then x = d(CStr(a(i)) & "|" & m)
            gl.Cells(r, 7).Value = x(0)
            gl.Cells(r, 8).Value = x(1)
            bal = bal + x(0) - x(1)
            gl.Cells(r, 9).Value = bal
            r = r + 1
        Next m
    Next i

    gl.Range("G13:I" & r - 1).NumberFormat = "#,##0.00;[Red](#,##0.00);-"
    aw.Columns("A:C").AutoFit
    MsgBox "GL refreshed for " & yr & ": 46 accounts x 12 months.", vbInformation
    Exit Sub

EH:
    MsgBox "RefreshGL error " & Err.Number & ": " & Err.description, vbCritical, "RefreshGL Error"
End Sub

Private Function SheetNamesList() As String()
    Dim arr() As String, i As Long
    ReDim arr(1 To ThisWorkbook.Sheets.count)
    For i = 1 To ThisWorkbook.Sheets.count
        arr(i) = "  • " & ThisWorkbook.Sheets(i).name
    Next i
    SheetNamesList = arr
End Function

' ============================================================
' Detect the GL reporting year
'   1. Look for a year in CDJ/CRJ/GJ dates (max date wins)
'   2. Fall back to the current system year
' ============================================================
Private Function DetectGLYear() As Long
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim v As Variant
    Dim maxDate As Date
    Dim found As Boolean
    Dim sheetNames As Variant
    Dim sn As Variant

    sheetNames = Array("CDJ", "CRJ", "GJ")
    For Each sn In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sn))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Select Case CStr(sn)
                Case "CDJ": lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
                Case "CRJ": lastRow = ws.Cells(ws.Rows.count, 4).End(xlUp).Row
                Case "GJ":  lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
            End Select
            For i = 8 To lastRow
                Select Case CStr(sn)
                    Case "CDJ": v = ws.Cells(i, 3).Value
                    Case "CRJ": v = ws.Cells(i, 4).Value
                    Case "GJ":  v = ws.Cells(i, 2).Value
                End Select
                If IsDate(v) Then
                    If Not found Or CDate(v) > maxDate Then
                        maxDate = CDate(v)
                        found = True
                    End If
                End If
            Next i
        End If
        Set ws = Nothing
    Next sn

    If found Then
        DetectGLYear = year(maxDate)
    Else
        DetectGLYear = year(Date)
    End If
End Function

' ============================================================
' GL HELPERS — public API
' ============================================================
Private Function Legacy_CalculateMonthlyNet(ByVal accountCode As String, ByVal month As Integer, _
                                    ByVal year As Integer, Optional ByVal sheetName As String = "") As Double
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim aw As Worksheet, ar As Long, x As Variant
    On Error Resume Next
    Set aw = ThisWorkbook.Sheets("GL_AUDIT")
    On Error GoTo 0
    If aw Is Nothing Then
        Set aw = ThisWorkbook.Worksheets.Add
        aw.name = "GL_AUDIT"
        aw.Range("A1:C1").Value = Array("Source", "Unmapped Detail", "Amount")
    End If
    ar = aw.Cells(aw.Rows.count, 1).End(xlUp).Row + 1
    GLScanCDJ d, aw, ar
    GLScanCRJ d
    GLScanGJ d, aw, ar
    x = Array(0#, 0#)
    If d.Exists(GLAcct(accountCode) & "|" & month) Then x = d(GLAcct(accountCode) & "|" & month)
    CalculateMonthlyNet = x(0) - x(1)
End Function

Private Function Legacy_CalculateEndingBalanceFull(ByVal accountCode As String, ByVal year As Integer, _
                                            ByVal startBalance As Double, Optional ByVal targetSheet As String = "") As Variant
    Dim z(1 To 12) As Double, m As Long
    z(1) = startBalance + CalculateMonthlyNet(accountCode, 1, year)
    For m = 2 To 12
        z(m) = z(m - 1) + CalculateMonthlyNet(accountCode, m, year)
    Next m
    CalculateEndingBalanceFull = z
End Function

Private Function Legacy_ValidateGLConsistency(Optional ByVal targetSheet As String = "GL") As Boolean
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets(targetSheet)
    Dim r As Long, td As Double, tc As Double, last As Long
    last = w.Cells(w.Rows.count, 6).End(xlUp).Row
    For r = 13 To last
        If IsNumeric(w.Cells(r, 7).Value) Then td = td + w.Cells(r, 7).Value
        If IsNumeric(w.Cells(r, 8).Value) Then tc = tc + w.Cells(r, 8).Value
    Next r
    ValidateGLConsistency = Abs(td - tc) <= GL_AUTO_TOL
End Function

Private Sub Legacy_RefreshAllGL(ByVal targetSheet As String, Optional ByVal startBalances As Variant)
    RefreshGL
End Sub

Public Sub InitializeCOALookup()
End Sub

Public Function GetAccountFromCOA(ByVal coaCode As String) As String
    GetAccountFromCOA = GLAcct(coaCode)
    If GetAccountFromCOA = "" Then GetAccountFromCOA = coaCode
End Function

Public Function GetAccountNameFromTIN(ByVal tin As String) As String
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("SUPPLIERS DATA")
    Dim r As Long, last As Long
    last = w.Cells(w.Rows.count, 3).End(xlUp).Row
    For r = 2 To last
        If Trim$(CStr(w.Cells(r, 3).Value)) = Trim$(tin) Then
            GetAccountNameFromTIN = CStr(w.Cells(r, 4).Value)
            Exit Function
        End If
    Next r
    GetAccountNameFromTIN = tin
End Function

Public Function GetFinancialMonth(ByVal dt As Date) As Integer
    GetFinancialMonth = month(dt)
End Function

Public Function GetFinancialYear(ByVal dt As Date) As Integer
    GetFinancialYear = year(dt)
End Function


