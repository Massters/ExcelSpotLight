Attribute VB_Name = "ControlSpotLight"
Public bEnableSpotlight As Boolean

' ������������رվ۹��
Public Sub ToggleSpotlight()
    If bEnableSpotlight = True Then
        bEnableSpotlight = False
        ThisWorkbook.myApp.SpotlightOff
    Else
        bEnableSpotlight = True
    End If
End Sub
