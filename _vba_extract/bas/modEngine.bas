Attribute VB_Name = "modEngine"

Option Explicit



' ============================================================

' ENGINE MODULE – modEngine

' PURPOSE: Central memory-first engine for PJ?CDJ and SJ?CRJ automation

' ============================================================



' ============================================================

' CONSTANTS – Column Mappings

' ============================================================



' --- PJ Sheet ---

Private Const PJ_COL_TIN As Integer = 4

Private Const PJ_COL_NAME As Integer = 5

Private Const PJ_COL_ADDRESS As Integer = 6

Private Const PJ_COL_COA As Integer = 7

Private Const PJ_COL_REF As Integer = 8

Private Const PJ_COL_NET As Integer = 10

Private Const PJ_COL_VAT As Integer = 13

Private Const PJ_COL_GROSS As Integer = 14

Private Const PJ_COL_FULLID As Integer = 17



' --- SUPPLIERS DATA ---

Private Const SUP_COL_TIN As Integer = 3

Private Const SUP_COL_NAME As Integer = 4

Private Const SUP_COL_ADDRESS As Integer = 5

Private Const SUP_COL_COA As Integer = 6



' --- CDJ Sheet ---

Private Const CDJ_COL_DATE As Integer = 3          ' C

Private Const CDJ_COL_PARTICULARS As Integer = 5   ' E

Private Const CDJ_COL_CASH As Integer = 6          ' F

Private Const CDJ_COL_VAT As Integer = 7           ' G

Private Const CDJ_COL_WHTAX As Integer = 8         ' H

Private Const CDJ_COL_CURRENCY As Integer = 25          ' Y

Private Const CDJ_COL_DAY As Integer = 4           ' D (day-of-month)
Private Const CDJ_COL_SUNDRY As Integer = 19         ' S (Sundry account description)

Private Const CDJ_COL_NOTE As Integer = 24         ' X

Private Const CDJ_COL_SEQ As Integer = 27          ' AA (sequence counter)

Private Const CDJ_COL_NEXTROW As Integer = 27      ' AA (next row counter)

Private Const CDJ_EXPENSE_START As Integer = 14    ' N

Private Const CDJ_EXPENSE_END As Integer = 19      ' S (T is the Debit total column, not an expense)

Private Const CDJ_DATA_START As Integer = 15       ' First data row

Private Const CDJ_COL_DEBIT As Integer = 20         ' T (Debit) = net + expanded WHT = gross



Private Const CDJ_ROW_SEQ As Integer = 13

Private Const CDJ_ROW_NEXTROW As Integer = 14



' --- CRJ Sheet (for SJ?CRJ sync) ---

Private Const CRJ_COL_DATE As Integer = 4           ' D

Private Const CRJ_COL_CUSTOMER As Integer = 6       ' F

Private Const CRJ_COL_INVOICE As Integer = 7        ' G

Private Const CRJ_COL_CASH As Integer = 8           ' H

Private Const CRJ_COL_OUTPUTVAT As Integer = 10     ' J

Private Const CRJ_COL_EXEMPT As Integer = 11        ' K

Private Const CRJ_COL_SALES As Integer = 12         ' L

Private Const CRJ_COL_NOTE As Integer = 15          ' O

Private Const CRJ_COL_REF As Integer = 16           ' P (full ID / reference)

Private Const CRJ_COL_SEQ As Integer = 17           ' Q – sequence counter

Private Const CRJ_COL_NEXTROW As Integer = 17       ' Q – next row counter

Private Const CRJ_ROW_SEQ As Integer = 8            ' Row 8 = sequence

Private Const CRJ_ROW_NEXTROW As Integer = 9        ' Row 9 = next row

Private Const CRJ_DATA_START As Integer = 10



' --- SJ Sheet (used by SyncSJToCRJ) ---

Private Const SJ_COL_DATE As Integer = 3

Private Const SJ_COL_CUSTOMER As Integer = 6

Private Const SJ_COL_AMOUNT As Integer = 10

Private Const SJ_COL_DISCOUNT As Integer = 11

Private Const SJ_COL_OUTPUTVAT As Integer = 13

Private Const SJ_COL_NETSALES As Integer = 15

Private Const SJ_COL_INVOICE As Integer = 9         ' Column I – short invoice number (SJ-00304)

Private Const SJ_COL_REFERENCE As Integer = 21      ' Column U – full ID (SJINV-...)



' ============================================================

' GLOBAL CACHE – CDJ

' ============================================================

Private dictTIN As Object

Private dictName As Object

Private dictRef As Object            ' CDJ reference ? row

Private arrHeaders As Variant

Private arrCols As Variant

Private gNextSeq As Long             ' PJ sequence

Private gNextRow As Long             ' Next CDJ row

Private gInitialized As Boolean

Private gSyncSheet As String

Private gHeaderCount As Integer



' GLOBAL CACHE – CRJ

Private dictCRJRef As Object         ' Invoice ? row

Private gNextCRJRow As Long          ' Next CRJ row

Private gNextCRJSeq As Long          ' CRJ invoice sequence (Q8)

Private gInitializedCRJ As Boolean



' ============================================================

' INITIALIZE ENGINE (PJ?CDJ part)

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



' INITIALIZE CRJ CACHE

Public Sub InitializeCRJ()

    If gInitializedCRJ Then Exit Sub



    Set dictCRJRef = CreateObject("Scripting.Dictionary")

    Call LoadCRJReferences

    Call LoadCRJNextRow

    Call LoadCRJSequence

    gInitializedCRJ = True

End Sub



' ============================================================

' LOAD SUPPLIERS (unchanged)

' ============================================================

Private Sub LoadSuppliers()

    Dim ws As Worksheet

    Dim lastRow As Long, i As Long

    Dim tin As String, name As String, address As String, coa As String



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    lastRow = ws.Cells(ws.Rows.count, SUP_COL_TIN).End(xlUp).row

    For i = 2 To lastRow

        tin = Trim(ws.Cells(i, SUP_COL_TIN).value)

        name = Trim(ws.Cells(i, SUP_COL_NAME).value)

        address = Trim(ws.Cells(i, SUP_COL_ADDRESS).value)

        coa = Trim(ws.Cells(i, SUP_COL_COA).value)

        If tin <> "" Then dictTIN(tin) = Array(name, address, coa)

        If name <> "" Then dictName(name) = Array(tin, address, coa)

    Next i

End Sub



' ============================================================

' LOAD CDJ HEADERS (unchanged)

' ============================================================

Private Sub LoadCDJHeaders()

    Dim ws As Worksheet

    Dim col As Integer, header As String, count As Integer



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(gSyncSheet)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    count = 0

    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END

        header = Trim(Trim(ws.Cells(13, col).value) & " " & Trim(ws.Cells(14, col).value))

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

        header = Trim(Trim(ws.Cells(13, col).value) & " " & Trim(ws.Cells(14, col).value))

        If header <> "" Then

            count = count + 1

            arrHeaders(count) = LCase(header)

            arrCols(count) = col

        End If

    Next col

    gHeaderCount = count

End Sub



' ============================================================

' LOAD CDJ REFERENCES (unchanged)

' ============================================================

Private Sub LoadCDJReferences()

    Dim ws As Worksheet

    Dim lastRow As Long, i As Long

    Dim ref As String



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(gSyncSheet)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).row

    If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).row

    For i = CDJ_DATA_START To lastRow

        ref = Trim(ws.Cells(i, CDJ_COL_REF).value)

        If ref <> "" Then dictRef(ref) = i

        ref = Trim(ws.Cells(i, CDJ_COL_REF_PNV).value)

        If ref <> "" Then dictRef(ref) = i

    Next i

End Sub



' ============================================================

' LOAD SEQUENCE FROM CDJ!Z13 (unchanged)

' ============================================================

Private Sub LoadSequence()

    Dim ws As Worksheet

    Dim val As Variant



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(gSyncSheet)

    On Error GoTo 0

    If ws Is Nothing Then

        gNextSeq = 0

        Exit Sub

    End If



    val = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).value

    If IsNumeric(val) Then

        gNextSeq = CLng(val)

    Else

        gNextSeq = 0

    End If

End Sub



' ============================================================

' LOAD NEXT ROW COUNTER FROM CDJ!Z14 (unchanged)

' ============================================================

Private Sub LoadNextRow()

    Dim ws As Worksheet

    Dim val As Variant



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(gSyncSheet)

    On Error GoTo 0

    If ws Is Nothing Then

        gNextRow = CDJ_DATA_START

        Exit Sub

    End If



    val = ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).value

    If IsNumeric(val) Then

        gNextRow = CLng(val)

    Else

        gNextRow = CDJ_DATA_START

        Do While ws.Cells(gNextRow, CDJ_COL_REF).value <> ""

            gNextRow = gNextRow + 1

        Loop

    End If

End Sub



' LOAD CRJ REFERENCES

Private Sub LoadCRJReferences()

    Dim ws As Worksheet

    Dim lastRow As Long, i As Long

    Dim inv As String



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CRJ")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).row

    For i = CRJ_DATA_START To lastRow

        inv = Trim(ws.Cells(i, CRJ_COL_INVOICE).value)

        If inv <> "" Then dictCRJRef(inv) = i

    Next i

End Sub



' LOAD NEXT ROW COUNTER FROM CRJ!Q9

Private Sub LoadCRJNextRow()

    Dim ws As Worksheet

    Dim val As Variant



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CRJ")

    On Error GoTo 0

    If ws Is Nothing Then

        gNextCRJRow = CRJ_DATA_START

        Exit Sub

    End If



    val = ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).value

    If IsNumeric(val) Then

        gNextCRJRow = CLng(val)

    Else

        gNextCRJRow = CRJ_DATA_START

        Do While ws.Cells(gNextCRJRow, CRJ_COL_INVOICE).value <> ""

            gNextCRJRow = gNextCRJRow + 1

        Loop

    End If

End Sub



' LOAD CRJ SEQUENCE FROM Q8

Private Sub LoadCRJSequence()

    Dim ws As Worksheet

    Dim val As Variant



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CRJ")

    On Error GoTo 0

    If ws Is Nothing Then

        gNextCRJSeq = 0

        Exit Sub

    End If



    val = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).value

    If IsNumeric(val) Then

        gNextCRJSeq = CLng(val)

    Else

        gNextCRJSeq = 0

    End If

End Sub



' ============================================================

' GET NEXT PJ NUMBER (unchanged)

' ============================================================

Public Function GetNextPJNumber() As String

    Dim ws As Worksheet

    If Not gInitialized Then Call InitializeEngine(gSyncSheet)



    gNextSeq = gNextSeq + 1

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(gSyncSheet)

    If Not ws Is Nothing Then ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).value = gNextSeq

    On Error GoTo 0



    GetNextPJNumber = "PJ-" & Format(gNextSeq, "00000")

End Function



' GET NEXT INVOICE NUMBER (for SJ?CRJ)

Public Function GetNextInvoiceNumber() As Long

    Dim ws As Worksheet

    If Not gInitializedCRJ Then Call InitializeCRJ



    gNextCRJSeq = gNextCRJSeq + 1

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CRJ")

    If Not ws Is Nothing Then ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).value = gNextCRJSeq

    On Error GoTo 0



    GetNextInvoiceNumber = gNextCRJSeq

End Function



' ============================================================

' SUPPLIER LOOKUPS (unchanged)

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

' EXPENSE COLUMN MAPPING (unchanged)

' ============================================================

Public Function FindExpenseColumn(description As String) As Integer

    Dim descLower As String, i As Integer

    If Not gInitialized Then Call InitializeEngine(gSyncSheet)

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

End Function



' ============================================================

' WITHHOLDING TAX (unchanged)

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

' CDJ REFERENCE CACHE (unchanged)

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



' CRJ REFERENCE CACHE

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

' SYNC PJ ? CDJ (unchanged)

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



    ref = Trim(sourceSheet.Cells(sourceRow, PJ_COL_REF).value)

    If ref = "" Or ref = "TOTAL" Then Exit Sub



    supplierName = Trim(sourceSheet.Cells(sourceRow, PJ_COL_NAME).value)

    description = Trim(sourceSheet.Cells(sourceRow, PJ_COL_COA).value)

    dateVal = sourceSheet.Cells(sourceRow, 2).value

    grossAmount = sourceSheet.Cells(sourceRow, PJ_COL_GROSS).value

    netAmount = sourceSheet.Cells(sourceRow, PJ_COL_NET).value

    vatAmount = sourceSheet.Cells(sourceRow, PJ_COL_VAT).value



    If grossAmount = 0 Or Not IsNumeric(grossAmount) Then Exit Sub



    expenseCol = FindExpenseColumn(description)

    If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

        whtax = CalculateWithholdingTax(Trim(Trim(wsTarget.Cells(13, expenseCol).value) & " " & Trim(wsTarget.Cells(14, expenseCol).value)), netAmount)

    Else

        whtax = 0

    End If

    targetRow = FindCDJRow(ref)



    If targetRow > 0 Then

        ' UPDATE

        With wsTarget

            .Cells(targetRow, CDJ_COL_DATE).value = dateVal

.Cells(targetRow, CDJ_COL_PARTICULARS).value = supplierName

.Cells(targetRow, CDJ_COL_SUNDRY).value = description

.Cells(targetRow, CDJ_COL_CASH).value = grossAmount * -1

.Cells(targetRow, CDJ_COL_VAT).value = vatAmount

.Cells(targetRow, CDJ_COL_WHTAX).value = whtax

.Cells(targetRow, CDJ_COL_REF).value = ref

For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END

    .Cells(targetRow, c).value = ""

Next c

If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

    .Cells(targetRow, expenseCol).value = netAmount

    .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount

End If

.Cells(targetRow, CDJ_COL_NOTE).value = "Updated from PJ Row " & sourceRow & " | " & Now

        End With

    Else

        ' INSERT

        targetRow = gNextRow

        If targetRow < CDJ_DATA_START Then

            targetRow = CDJ_DATA_START

            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).value <> ""

                targetRow = targetRow + 1

            Loop

            gNextRow = targetRow

        End If



        With wsTarget

            .Cells(targetRow, CDJ_COL_DATE).value = dateVal

.Cells(targetRow, CDJ_COL_PARTICULARS).value = supplierName

.Cells(targetRow, CDJ_COL_SUNDRY).value = description

.Cells(targetRow, CDJ_COL_CASH).value = grossAmount * -1

.Cells(targetRow, CDJ_COL_VAT).value = vatAmount

.Cells(targetRow, CDJ_COL_WHTAX).value = whtax

.Cells(targetRow, CDJ_COL_REF).value = ref

If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then

    .Cells(targetRow, expenseCol).value = netAmount

    .Cells(targetRow, CDJ_COL_DEBIT).value = grossAmount

End If

.Cells(targetRow, CDJ_COL_NOTE).value = "Inserted from PJ Row " & sourceRow & " | " & Now

            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        End With



        UpdateCDJReference ref, targetRow

        gNextRow = targetRow + 1

        On Error Resume Next

        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).value = gNextRow

        On Error GoTo 0

    End If



    Exit Sub

ErrorHandler:

    ' Silent fail

End Sub



' SYNC SJ ? CRJ (FIXED)

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



    invoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_INVOICE).value

    If invoiceNum = "" Or invoiceNum = "TOTAL" Then Exit Sub



    customerName = Trim(sourceSheet.Cells(sourceRow, SJ_COL_CUSTOMER).value)

    dateVal = sourceSheet.Cells(sourceRow, SJ_COL_DATE).value

    referenceInvoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_REFERENCE).value

    amount = sourceSheet.Cells(sourceRow, SJ_COL_AMOUNT).value

    discount = sourceSheet.Cells(sourceRow, SJ_COL_DISCOUNT).value

    outputVAT = sourceSheet.Cells(sourceRow, SJ_COL_OUTPUTVAT).value

    netSales = sourceSheet.Cells(sourceRow, SJ_COL_NETSALES).value



    cashInBank = amount - discount



    If customerName = "" Or amount = 0 Or Not IsNumeric(amount) Then Exit Sub



    targetRow = FindCRJRow(CStr(invoiceNum))



    If targetRow > 0 Then

        ' UPDATE

        With wsTarget

            .Cells(targetRow, CRJ_COL_DATE).value = dateVal

            .Cells(targetRow, CRJ_COL_CUSTOMER).value = customerName

            .Cells(targetRow, CRJ_COL_INVOICE).value = invoiceNum

            .Cells(targetRow, CRJ_COL_REF).value = referenceInvoiceNum

            .Cells(targetRow, CRJ_COL_CASH).value = cashInBank

            .Cells(targetRow, CRJ_COL_OUTPUTVAT).value = outputVAT

            .Cells(targetRow, CRJ_COL_EXEMPT).value = 0

            .Cells(targetRow, CRJ_COL_SALES).value = netSales

            .Cells(targetRow, CRJ_COL_NOTE).value = "Updated from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now

        End With

    Else

        ' INSERT (FIXED: added CRJ_COL_REF)

        targetRow = gNextCRJRow

        If targetRow < CRJ_DATA_START Then

            targetRow = CRJ_DATA_START

            Do While wsTarget.Cells(targetRow, CRJ_COL_INVOICE).value <> ""

                targetRow = targetRow + 1

            Loop

            gNextCRJRow = targetRow

        End If



        With wsTarget

            .Cells(targetRow, CRJ_COL_DATE).value = dateVal

            .Cells(targetRow, CRJ_COL_CUSTOMER).value = customerName

            .Cells(targetRow, CRJ_COL_INVOICE).value = invoiceNum

            .Cells(targetRow, CRJ_COL_REF).value = referenceInvoiceNum   ' ? FIXED: was missing

            .Cells(targetRow, CRJ_COL_CASH).value = cashInBank

            .Cells(targetRow, CRJ_COL_OUTPUTVAT).value = outputVAT

            .Cells(targetRow, CRJ_COL_EXEMPT).value = 0

            .Cells(targetRow, CRJ_COL_SALES).value = netSales

            .Cells(targetRow, CRJ_COL_NOTE).value = "Inserted from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now

            .Cells(targetRow, CRJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, CRJ_COL_OUTPUTVAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, CRJ_COL_EXEMPT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, CRJ_COL_SALES).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        End With



        UpdateCRJReference CStr(invoiceNum), targetRow

        gNextCRJRow = targetRow + 1

        On Error Resume Next

        wsTarget.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).value = gNextCRJRow

        On Error GoTo 0

    End If



    Exit Sub

ErrorHandler:

    ' Silent fail

End Sub



' ============================================================

' ENGINE STATUS (enhanced)

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



' ============================================================

' CHECK CDJ STATUS (unchanged)

' ============================================================

Public Sub ShowCDJStatus()

    Dim ws As Worksheet

    Dim msg As String

    Dim lastRow As Long

    Dim i As Long



    If Not gInitialized Then Call InitializeEngine(gSyncSheet)



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CDJ")

    On Error GoTo 0



    If ws Is Nothing Then

        MsgBox "Target sheet '" & gSyncSheet & "' not found!", vbCritical

        Exit Sub

    End If



    msg = "===== CDJ STATUS REPORT =====" & vbCrLf & vbCrLf

    msg = msg & "Target Sheet: " & gSyncSheet & vbCrLf

    msg = msg & "Data starts at row: " & CDJ_DATA_START & vbCrLf

    msg = msg & "Current row counter (Z14): " & gNextRow & vbCrLf



    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).row

    If lastRow < CDJ_DATA_START Then

        msg = msg & "Last used row (with reference): None yet" & vbCrLf

    Else

        msg = msg & "Last used row: " & lastRow & vbCrLf

    End If



    msg = msg & "References in cache: " & dictRef.count & vbCrLf

    msg = msg & "Sequence (Z13): " & IIf(IsNumeric(ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).value), ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).value, "empty") & vbCrLf

    msg = msg & "Next PJ number: " & Format(gNextSeq + 1, "PJ-00000") & vbCrLf



    msg = msg & vbCrLf & "Expense headers (row 14, N-T):" & vbCrLf

    For i = 1 To gHeaderCount

        msg = msg & "  Col " & arrCols(i) & ": " & arrHeaders(i) & vbCrLf

    Next i



    MsgBox msg, vbInformation, "CDJ Status"

End Sub



' CHECK CRJ STATUS (FIXED: labels now say Q8/Q9)

Public Sub ShowCRJStatus()

    Dim ws As Worksheet

    Dim msg As String

    Dim lastRow As Long



    If Not gInitializedCRJ Then Call InitializeCRJ



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CRJ")

    On Error GoTo 0



    If ws Is Nothing Then

        MsgBox "CRJ sheet not found!", vbCritical

        Exit Sub

    End If



    msg = "===== CRJ STATUS REPORT =====" & vbCrLf & vbCrLf

    msg = msg & "Data starts at row: " & CRJ_DATA_START & vbCrLf

    msg = msg & "Current row counter (Q9): " & gNextCRJRow & vbCrLf



    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).row

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

' REPAIR ENGINE (CDJ + CRJ)

' ============================================================

Public Sub RepairEngine()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim seqVal As Variant



    ' --- Repair CDJ ---

    gSyncSheet = "CDJ"

    Call InitializeEngine(gSyncSheet)



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CDJ")

    On Error GoTo 0



    If Not ws Is Nothing Then

        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).row

        If lastRow < CDJ_DATA_START Then lastRow = CDJ_DATA_START - 1

        gNextRow = lastRow + 1

        ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).value = gNextRow



        seqVal = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).value

        If IsNumeric(seqVal) Then gNextSeq = CLng(seqVal) Else gNextSeq = 0: ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).value = 0

        Call LoadCDJReferences

    End If



    ' --- Repair CRJ ---

    If Not gInitializedCRJ Then Call InitializeCRJ



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CRJ")

    On Error GoTo 0



    If Not ws Is Nothing Then

        lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).row

        If lastRow < CRJ_DATA_START Then lastRow = CRJ_DATA_START - 1

        gNextCRJRow = lastRow + 1

        ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).value = gNextCRJRow



        seqVal = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).value

        If IsNumeric(seqVal) Then

            gNextCRJSeq = CLng(seqVal)

        Else

            gNextCRJSeq = 0

            ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).value = 0

        End If

        Call LoadCRJReferences

    End If



    MsgBox "Engine repaired!" & vbCrLf & vbCrLf & EngineStatus, vbInformation

End Sub



' ============================================================

' RESET CRJ – Clear all data and counters

' ============================================================

Public Sub ResetCRJ()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim confirm As VbMsgBoxResult



    confirm = MsgBox("This will DELETE ALL data in the CRJ sheet and reset counters." & vbCrLf & _

                     "Are you sure?", vbYesNo + vbExclamation, "Reset CRJ")

    If confirm <> vbYes Then Exit Sub



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CRJ")

    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox "CRJ sheet not found!", vbCritical

        Exit Sub

    End If



    Application.ScreenUpdating = False

    Application.EnableEvents = False



    ' 1. Clear all data rows (from CRJ_DATA_START down)

    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).row

    If lastRow >= CRJ_DATA_START Then

        ws.Rows(CRJ_DATA_START & ":" & lastRow).ClearContents

        ' Also clear formats if you want fully blank rows – optional:

        ' ws.Rows(CRJ_DATA_START & ":" & lastRow).Clear

    End If



    ' 2. Reset sequence counter (Q8) to 0

    ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).value = 0

    gNextCRJSeq = 0



    ' 3. Reset next row counter (Q9) to CRJ_DATA_START

    ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).value = CRJ_DATA_START

    gNextCRJRow = CRJ_DATA_START



    ' 4. Re-initialize cache (clear all stored invoice references)

    If Not gInitializedCRJ Then Call InitializeCRJ

    dictCRJRef.RemoveAll

    ' Reload (will find no invoices)

    ' (No need to reload as we already set counters; but to be safe:)

    gInitializedCRJ = True   ' already true, but ensures dict is empty



    Application.EnableEvents = True

    Application.ScreenUpdating = True



    MsgBox "CRJ has been reset." & vbCrLf & _

           "Sequence (Q8) = 0, Next row (Q9) = " & CRJ_DATA_START, vbInformation

End Sub



' ============================================================

' RESET CDJ – Clear all data and counters

' ============================================================

Public Sub ResetCDJ()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim confirm As VbMsgBoxResult



    confirm = MsgBox("This will DELETE ALL data in the CDJ sheet and reset counters." & vbCrLf & _

                     "Are you sure?", vbYesNo + vbExclamation, "Reset CDJ")

    If confirm <> vbYes Then Exit Sub



    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("CDJ")

    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox "CDJ sheet not found!", vbCritical

        Exit Sub

    End If



    Application.ScreenUpdating = False

    Application.EnableEvents = False



    ' 1. Clear all data rows (from CDJ_DATA_START down)

    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).row

    If lastRow >= CDJ_DATA_START Then

        ws.Rows(CDJ_DATA_START & ":" & lastRow).ClearContents

    End If



    ' 2. Reset sequence counter (Z13) to 0

    ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).value = 0

    gNextSeq = 0



    ' 3. Reset next row counter (Z14) to CDJ_DATA_START

    ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).value = CDJ_DATA_START

    gNextRow = CDJ_DATA_START



    ' 4. Re-initialize CDJ cache (clear stored references)

    If Not gInitialized Then Call InitializeEngine(gSyncSheet)

    dictRef.RemoveAll

    ' reload references (will be empty)

    Call LoadCDJReferences



    Application.EnableEvents = True

    Application.ScreenUpdating = True



    MsgBox "CDJ has been reset." & vbCrLf & _

           "Sequence (Z13) = 0, Next row (Z14) = " & CDJ_DATA_START, vbInformation

End Sub



' ============================================================

' CALCULATE EWT (unchanged)

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

' SYNC PJ NON-VAT ? CDJ (unchanged)

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



    ref = Trim(sourceSheet.Cells(sourceRow, PJ_COL_REF).value)

    If ref = "" Or ref = "TOTAL" Then Exit Sub



    supplierName = Trim(sourceSheet.Cells(sourceRow, PJ_COL_NAME).value)

    description = Trim(sourceSheet.Cells(sourceRow, PJ_COL_COA).value)

    dateVal = sourceSheet.Cells(sourceRow, 2).value

    netAmount = sourceSheet.Cells(sourceRow, 14).value



    If netAmount = 0 Or Not IsNumeric(netAmount) Then Exit Sub



    whtax = CalculateEWT(description, netAmount)

    expenseCol = FindExpenseColumn(description)

    targetRow = FindCDJRow(ref)



    If targetRow > 0 Then

        With wsTarget

            .Cells(targetRow, CDJ_COL_DATE).value = dateVal

.Cells(targetRow, CDJ_COL_PARTICULARS).value = supplierName

.Cells(targetRow, CDJ_COL_SUNDRY).value = description

.Cells(targetRow, CDJ_COL_CASH).value = netAmount * -1

.Cells(targetRow, CDJ_COL_VAT).value = 0

.Cells(targetRow, CDJ_COL_WHTAX).value = whtax

.Cells(targetRow, CDJ_COL_REF_PNV).value = ref

.Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value

For c = CDJ_EXPENSE_START To CDJ_EXPENSE_END

    .Cells(targetRow, c).value = ""

Next c

.Cells(targetRow, expenseCol).value = netAmount

.Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax

.Cells(targetRow, CDJ_COL_NOTE).value = "Updated Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now

        End With

    Else

        targetRow = gNextRow

        If targetRow < CDJ_DATA_START Then

            targetRow = CDJ_DATA_START

            Do While wsTarget.Cells(targetRow, CDJ_COL_REF).value <> "" Or _
                     wsTarget.Cells(targetRow, CDJ_COL_REF_PNV).value <> ""

                targetRow = targetRow + 1

            Loop

            gNextRow = targetRow

        End If



        With wsTarget

            .Cells(targetRow, CDJ_COL_DATE).value = dateVal

.Cells(targetRow, CDJ_COL_PARTICULARS).value = supplierName

.Cells(targetRow, CDJ_COL_SUNDRY).value = description

.Cells(targetRow, CDJ_COL_CASH).value = netAmount * -1

.Cells(targetRow, CDJ_COL_VAT).value = 0

.Cells(targetRow, CDJ_COL_WHTAX).value = whtax

.Cells(targetRow, CDJ_COL_REF_PNV).value = ref

.Cells(targetRow, CDJ_COL_DAY).value = sourceSheet.Cells(sourceRow, 2).value

.Cells(targetRow, expenseCol).value = netAmount

.Cells(targetRow, CDJ_COL_DEBIT).value = netAmount + whtax

.Cells(targetRow, CDJ_COL_NOTE).value = "Inserted Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now

            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        End With



        UpdateCDJReference ref, targetRow

        gNextRow = targetRow + 1

        On Error Resume Next

        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).value = gNextRow

        On Error GoTo 0

    End If



    Exit Sub

ErrorHandler:

    ' Silent fail

End Sub





