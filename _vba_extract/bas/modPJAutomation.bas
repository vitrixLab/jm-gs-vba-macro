Attribute VB_Name = "modPJAutomation"

Option Explicit



' ============================================================

' GLOBAL VARIABLES (Cache for speed)

' ============================================================

Private dictTIN As Object          ' Scripting.Dictionary

Private dictName As Object         ' Scripting.Dictionary

Private gMaxSeq As Long            ' Highest PJ-XXXXX sequence ever assigned

Private gInitialized As Boolean



' ============================================================

' INITIALIZE CACHE (Call once on Workbook Open)

' ============================================================

Public Sub InitializePJAutomation()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim i As Long

    Dim tin As String

    Dim name As String

    Dim address As String

    Dim coa As String

    

    On Error Resume Next

    Set dictTIN = CreateObject("Scripting.Dictionary")

    Set dictName = CreateObject("Scripting.Dictionary")

    On Error GoTo 0

    

    If dictTIN Is Nothing Then

        MsgBox "Failed to create dictionaries. Please enable Microsoft Scripting Runtime.", vbCritical

        Exit Sub

    End If

    

    ' Load SUPPLIERS DATA

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    If ws Is Nothing Then

        MsgBox "SUPPLIERS DATA sheet not found!", vbCritical

        Exit Sub

    End If

    On Error GoTo 0

    

    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row  ' Column B = NO

    If lastRow < 2 Then lastRow = 2

    

    ' Load into dictionaries (normalized keys)

    For i = 2 To lastRow

        tin = NormalizeString(ws.Cells(i, 3).value)   ' C = TIN#

        name = NormalizeString(ws.Cells(i, 4).value)  ' D = PARTICULARS

        address = ws.Cells(i, 5).value                ' E = ADDRESS

        coa = ws.Cells(i, 6).value                    ' F = CHART OF ACCOUNT

        

        If tin <> "" Then

            ' Store as array: Name, Address, COA

            dictTIN(tin) = Array(name, address, coa)

        End If

        If name <> "" Then

            dictName(name) = Array(tin, address, coa)

        End If

    Next i

    

    ' Compute max sequence from PJ column H

    gMaxSeq = GetMaxSequenceFromPJ()

    gInitialized = True

End Sub



' ============================================================

' GET MAX SEQUENCE FROM PJ (Runs only once)

' ============================================================

Private Function GetMaxSequenceFromPJ() As Long

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim i As Long

    Dim cellVal As String

    Dim dashPos As Long

    Dim numPart As String

    Dim maxSeq As Long

    

    Set ws = ThisWorkbook.Sheets("PJ")

    If ws Is Nothing Then

        GetMaxSequenceFromPJ = 0

        Exit Function

    End If

    

    lastRow = ws.Cells(ws.Rows.count, 8).End(xlUp).row  ' Column H

    maxSeq = 0

    

    For i = 8 To lastRow

        cellVal = ws.Cells(i, 8).value

        If cellVal <> "" And cellVal <> "TOTAL" Then

            dashPos = InStrRev(cellVal, "-")

            If dashPos > 0 Then

                numPart = Mid(cellVal, dashPos + 1)

                If IsNumeric(numPart) Then

                    If CLng(numPart) > maxSeq Then

                        maxSeq = CLng(numPart)

                    End If

                End If

            End If

        End If

    Next i

    

    GetMaxSequenceFromPJ = maxSeq

End Function



' ============================================================

' NORMALIZE STRING (Fast)

' ============================================================

Private Function NormalizeString(ByVal text As String) As String

    If text = "" Then

        NormalizeString = ""

        Exit Function

    End If

    NormalizeString = Application.WorksheetFunction.Trim(text)

End Function



' ============================================================

' ADD SUPPLIER TO CACHE (Call when adding new supplier)

' ============================================================

Public Sub AddSupplierToCache(tin As String, name As String, address As String, coa As String)

    If Not gInitialized Then Exit Sub

    tin = NormalizeString(tin)

    name = NormalizeString(name)

    If tin <> "" Then

        dictTIN(tin) = Array(name, address, coa)

    End If

    If name <> "" Then

        dictName(name) = Array(tin, address, coa)

    End If

End Sub



' ============================================================

' GET SUPPLIER BY TIN (Fast)

' ============================================================

Public Function GetSupplierByTIN(tin As String) As Variant

    If Not gInitialized Then

        InitializePJAutomation

        If Not gInitialized Then

            GetSupplierByTIN = False

            Exit Function

        End If

    End If

    tin = NormalizeString(tin)

    If dictTIN.Exists(tin) Then

        GetSupplierByTIN = dictTIN(tin)

    Else

        GetSupplierByTIN = False

    End If

End Function



' ============================================================

' GET SUPPLIER BY NAME (Fast)

' ============================================================

Public Function GetSupplierByName(name As String) As Variant

    If Not gInitialized Then

        InitializePJAutomation

        If Not gInitialized Then

            GetSupplierByName = False

            Exit Function

        End If

    End If

    name = NormalizeString(name)

    If dictName.Exists(name) Then

        GetSupplierByName = dictName(name)

    Else

        GetSupplierByName = False

    End If

End Function



' ============================================================

' GET NEXT SEQUENCE NUMBER (Fast)

' ============================================================

Public Function GetNextSeq() As Long

    If Not gInitialized Then

        InitializePJAutomation

        If Not gInitialized Then

            GetNextSeq = 1

            Exit Function

        End If

    End If

    gMaxSeq = gMaxSeq + 1

    GetNextSeq = gMaxSeq

End Function





