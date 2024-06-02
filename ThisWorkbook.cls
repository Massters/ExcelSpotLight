VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public myApp As AppEvents

'Dim myApp As New AppEvents

Private Sub Workbook_Open()
    Set myApp = New AppEvents
    Set myApp.App = Application
End Sub
