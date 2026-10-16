Attribute VB_Name = "Sheet9"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit

' GATE2: these constants were previously undeclared. Without Option Explicit
' they silently resolved to Empty, every column guard failed, and
' Me.Cells(c.Row, COL_REF) evaluated as Cells(row, 0) and raised an error.
' Values match the verified PJ Non-Vat2 layout (header row 6, data from row 8).
Private Const START_ROW As Long = 8
Private Const COL_TIN As Long = 4
Private Const COL_NAME As Long = 5
Private Const COL_ADDR As Long = 6
Private Const COL_DESC As Long = 7
Private Const COL_COA As Long = 7
Private Const COL_REF As Long = 8
Private Const COL_AMOUNT As Long = 9
Private Const COL_DEBIT As Long = 10
Private Const COL_DEBIT_MIRROR As Long = 9
Private Const COL_DISCOUNT As Long = 11
Private Const COL_EWT As Long = 12
Private Const COL_VAT As Long = 13
Private Const COL_NET As Long = 14
Private Const COL_FULL_ID As Long = 17

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range
    Dim c As Range

    Dim currentTimestamp As String
    Dim tinVal As String
    Dim nameVal As String
    Dim refVal As String
    Dim coaVal As String
    Dim amountVal As Variant
    Dim seqNum As Long

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    On Error GoTo ErrorHandler

    ' ============================================================
    ' SAVE EXCEL APPLICATION STATE
    ' ============================================================

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation

    ' ============================================================
    ' LIMIT EVENT TO USER-EDITABLE PJ NON-VAT INPUT COLUMNS
    '
    ' D = TIN
    ' E = Supplier / Vendor Name
    ' F = Address
    ' G = Description
    ' I = Non-VAT Amount
    ' ============================================================

    Set rng = Intersect( _
        Target, _
        Union( _
            Me.Columns("D"), _
            Me.Columns("E"), _
            Me.Columns("F"), _
            Me.Columns("G"), _
            Me.Columns("I") _
        ) _
    )

    If rng Is Nothing Then Exit Sub

    ' ============================================================
    ' PROTECT AGAINST RE-ENTRY
    ' ============================================================

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    currentTimestamp = Format$(Now, "yyyymmdd-HHnnss")

    ' ============================================================
    ' PASS 1
    '
    ' SUPPLIER LOOKUP
    ' NON-VAT AMOUNT CALCULATION
    ' REFERENCE GENERATION
    ' ============================================================

    For Each c In rng.Cells

        If c.Row >= START_ROW Then

            ' ====================================================
            ' SUPPLIER LOOKUP
            ' ====================================================

            If c.Column = COL_TIN Or c.Column = COL_NAME Then

                tinVal = Trim$(CStr( _
                    Me.Cells(c.Row, COL_TIN).Value _
                ))

                nameVal = Trim$(CStr( _
                    Me.Cells(c.Row, COL_NAME).Value _
                ))

                ' ------------------------------------------------
                ' TOTAL ROW
                ' ------------------------------------------------

                If StrComp(tinVal, "TOTAL", vbTextCompare) = 0 _
                   Or StrComp(nameVal, "TOTAL", vbTextCompare) = 0 Then

                    ' IMPORTANT:
                    ' CalculateTotals receives a Long row number.
                    ' The helper declaration must therefore be:
                    '
                    ' Private Sub CalculateTotals(ByVal rowNum As Long)

                    CalculateTotals CLng(c.Row)

                ' ------------------------------------------------
                ' LOOKUP BY TIN
                ' ------------------------------------------------

                ElseIf Len(tinVal) > 0 Then

                    LookupByTIN_NonVat CLng(c.Row), tinVal

                ' ------------------------------------------------
                ' LOOKUP BY NAME
                ' ------------------------------------------------

                ElseIf Len(nameVal) > 0 Then

                    LookupByName_NonVat CLng(c.Row), nameVal

                ' ------------------------------------------------
                ' BOTH EMPTY
                ' ------------------------------------------------

                Else

                    ClearSupplierRow CLng(c.Row)

                End If

            End If

            ' ====================================================
            ' NON-VAT AMOUNT
            ' ====================================================

            If c.Column = COL_AMOUNT Then

                amountVal = Me.Cells(c.Row, COL_AMOUNT).Value

                tinVal = UCase$(Trim$(CStr( _
                    Me.Cells(c.Row, COL_TIN).Value _
                )))

                nameVal = UCase$(Trim$(CStr( _
                    Me.Cells(c.Row, COL_NAME).Value _
                )))

                ' ------------------------------------------------
                ' DO NOT PROCESS TOTAL ROWS
                ' ------------------------------------------------

                If tinVal <> "TOTAL" And nameVal <> "TOTAL" Then

                    ' ------------------------------------------------
                    ' VALID AMOUNT
                    ' ------------------------------------------------

                    If Len(Trim$(CStr(amountVal))) > 0 Then

                        ' --------------------------------------------
                        ' DEBIT AMOUNT
                        ' --------------------------------------------

                        Me.Cells(c.Row, COL_DEBIT).Value = amountVal

                        ' --------------------------------------------
                        ' DEBIT MIRROR
                        ' --------------------------------------------

                        Me.Cells(c.Row, COL_DEBIT_MIRROR).Value = _
                            amountVal

                        ' --------------------------------------------
                        ' NON-VAT
                        '
                        ' VAT = ZERO
                        ' EWT = ZERO
                        ' --------------------------------------------

                        Me.Cells(c.Row, 13).Value = 0
                        Me.Cells(c.Row, 12).Value = 0

                        ' --------------------------------------------
                        ' REFERENCE GENERATION
                        ' --------------------------------------------

                        refVal = Trim$(CStr( _
                            Me.Cells(c.Row, COL_REF).Value _
                        ))

                        coaVal = Trim$(CStr( _
                            Me.Cells(c.Row, COL_COA).Value _
                        ))

                        ' GATE2: reference generation is disabled on this sheet.
                        ' "PJ Non-Vat2" mirrors PJ and must not consume the shared
                        ' PJ reference sequence, nor mint refs that are never posted.
                        If False Then

                            seqNum = CLng(GetNextSequence())

                            ' ----------------------------------------
                            ' PJ NON-VAT REFERENCE
                            ' ----------------------------------------

                            Me.Cells(c.Row, COL_REF).Value = _
                                "PNV-" & Format$(seqNum, "00000")

                            ' ----------------------------------------
                            ' FULL UNIQUE TRANSACTION ID
                            ' ----------------------------------------

                            Me.Cells(c.Row, COL_FULL_ID).Value = _
                                "PNVINV-" & currentTimestamp & "-" & _
                                Format$(seqNum, "00000")

                        End If

                    Else

                        ' ------------------------------------------------
                        ' AMOUNT CLEARED
                        ' ------------------------------------------------

                        Me.Cells(c.Row, COL_DEBIT).ClearContents

                        Me.Cells(c.Row, COL_DEBIT_MIRROR).ClearContents

                        Me.Cells(c.Row, 12).Value = 0
                        Me.Cells(c.Row, 13).Value = 0

                    End If

                End If

            End If

        End If

    Next c

    ' ============================================================
    ' PASS 2
    '
    ' SYNCHRONIZE VALID PNV RECORDS TO CDJ
    ' ============================================================

    For Each c In rng.Cells

        If c.Row >= START_ROW Then

            refVal = Trim$(CStr( _
                Me.Cells(c.Row, COL_REF).Value _
            ))

            ' ----------------------------------------------------
            ' ONLY REAL PNV RECORDS
            ' ----------------------------------------------------

            If Len(refVal) > 0 _
               And StrComp(refVal, "TOTAL", vbTextCompare) <> 0 Then

                SyncToCDJ_NonVat CLng(c.Row)

            End If

        End If

    Next c

CleanExit:

    ' ============================================================
    ' RESTORE ORIGINAL EXCEL APPLICATION STATE
    ' ============================================================

    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = oldCalculation

    Exit Sub

ErrorHandler:

    MsgBox _
        "PJ Non-VAT Worksheet_Change error:" & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.description, _
        vbExclamation, _
        "PNV Automation"

    Resume CleanExit

End Sub




