Attribute VB_Name = "Sheet10"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Sub CheckDuplicates()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim i As Long

    Dim j As Long

    Dim name1 As String, add1 As String

    Dim name2 As String, add2 As String

    Dim found As Boolean

    

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row

    found = False

    

    For i = 2 To lastRow

        name1 = ws.Cells(i, "D").value

        add1 = ws.Cells(i, "E").value

        If name1 <> "" And add1 <> "" And add1 <> "—" Then

            For j = i + 1 To lastRow

                name2 = ws.Cells(j, "D").value

                add2 = ws.Cells(j, "E").value

                If name1 = name2 And add1 = add2 Then

                    Debug.Print "DUPLICATE: Row " & i & " and Row " & j & " - " & name1 & " | " & add1

                    found = True

                End If

            Next j

        End If

    Next i

    

    If Not found Then

        MsgBox "No duplicates found.", vbInformation

    Else

        MsgBox "Duplicates found! Check Immediate Window (Ctrl+G).", vbExclamation

    End If

End Sub





