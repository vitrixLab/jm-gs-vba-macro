Attribute VB_Name = "Sheet8"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' ============================================================
' SJ — Worksheet module
' Delegates CRJ sync + invoice numbering to modEngine
' ============================================================

Private Const START_ROW As Long = 10

' --- Column layout (must match modEngine.SJ_COL_* constants) ---
Private Const COL_DATE As Integer = 3           ' C
Private Const COL_CUSTOMER As Integer = 6       ' F
Private Const COL_INVOICE As Integer = 9        ' I
Private Const COL_AMOUNT As Integer = 10        ' J
Private Const COL_DISCOUNT As Integer = 11      ' K
Private Const COL_OUTPUTVAT As Integer = 13     ' M
Private Const COL_NETSALES As Integer = 15      ' O
Private Const COL_REFERENCE As Integer = 21     ' U

' ============================================================
' Sync one SJ row to CRJ (via modEngine)
' ============================================================
Private Sub SyncToCRJ(ByVal sourceRow As Long)
    modEngine.SyncSJToCRJ Me, sourceRow
End Sub

' ============================================================
' MAIN TRIGGER
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range, cell As Range

    ' Watch customer (F), amount (J), discount (K) edits
    Set rng = Intersect(Target, Union(Me.Range("F:F"), Me.Range("J:K")))
    If rng Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup

    For Each cell In rng
        If cell.row >= START_ROW _
           And UCase$(Trim(CStr(Me.Cells(cell.row, 1).value))) <> "TOTAL" Then

            ' ---- Number formats ----
            Me.Cells(cell.row, COL_AMOUNT).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            Me.Cells(cell.row, COL_DISCOUNT).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            Me.Cells(cell.row, COL_OUTPUTVAT).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            Me.Cells(cell.row, COL_NETSALES).NumberFormat = _
                "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

            ' ---- Formulas (J and K are the amount columns) ----
            ' Output VAT = (Amount - Discount) * 0.107142857
            Me.Cells(cell.row, COL_OUTPUTVAT).Formula = _
                "=IF(J" & cell.row & "-K" & cell.row & "=0,0," & _
                "(J" & cell.row & "-K" & cell.row & ")*0.107142857)"

            ' Net Sales = Amount - Discount - Output VAT
            Me.Cells(cell.row, COL_NETSALES).Formula = _
                "=J" & cell.row & "-K" & cell.row & "-M" & cell.row

            ' ---- Auto-assign invoice number (via engine) ----
            If Me.Cells(cell.row, COL_INVOICE).value = "" _
               And Me.Cells(cell.row, COL_DATE).value <> "" _
               And Me.Cells(cell.row, COL_AMOUNT).value <> "" Then
                Me.Cells(cell.row, COL_INVOICE).value = modEngine.GetNextInvoiceNumber()
            End If
        End If
    Next cell

    ' ---- CRJ SYNCHRONIZATION ----
    For Each cell In rng
        If cell.row >= START_ROW _
           And Me.Cells(cell.row, COL_INVOICE).value <> "" _
           And Me.Cells(cell.row, COL_AMOUNT).value <> "" Then
            Call SyncToCRJ(cell.row)
        End If
    Next cell

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

