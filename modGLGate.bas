Attribute VB_Name = "modGLGate"
Option Explicit

' v8.0 fail-closed gate. Exact workbook mapping is handled by modGLWorkbookMap.

Public Sub RunV8Gate()
    If ValidateGLConsistency(2026) Then
        If BuildV8CalcSheet(2026) Then
            MsgBox "v8.0 gate PASS: posting population reconciles and 46x12 matrix exists.",vbInformation
        Else
            MsgBox "v8.0 gate HOLD: matrix size is not 552.",vbExclamation
        End If
    Else
        MsgBox "v8.0 gate HOLD: review GL_AUDIT for unmapped or invalid activity.",vbExclamation
    End If
End Sub
