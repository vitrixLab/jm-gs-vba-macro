Attribute VB_Name = "ThisWorkbook"

Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Option Explicit



Private Sub Workbook_Open()

    ' Initialize the engine – this loads all suppliers, CDJ references, headers, and sequence

    ' into memory. The target sheet is "CDJ" (you can change to "TEST" for testing).

    Call modEngine.InitializeEngine("CDJ")

    

    ' Optional: show a message to confirm

    ' MsgBox "Engine initialized. Status: " & modEngine.EngineStatus, vbInformation

End Sub





