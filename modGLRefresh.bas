Attribute VB_Name = "modGLRefresh"
Option Explicit

' v8.3.3 compatibility wrapper.
' The old RefreshGL matrix writer is retired from the execution path.
' RefreshGL now delegates to the row-scanning engine so the visible GL,
' GL_V8_CALC, and GL_AUDIT all come from the same scanned postings.

Public Sub RefreshGLIntoSheet(Optional ByVal yearNumber As Long = 2026)
    ScanGL yearNumber
End Sub

Public Sub RefreshGL(Optional ByVal yearNumber As Long = 2026)
    ScanGL yearNumber
End Sub

Public Sub RefreshAllGL(Optional ByVal yearNumber As Long = 2026)
    ScanGL yearNumber
End Sub
