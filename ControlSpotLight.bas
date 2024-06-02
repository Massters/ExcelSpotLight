Attribute VB_Name = "ControlSpotLight"
Public bEnableSpotlight As Boolean

' 函数：开启或关闭聚光灯
Public Sub ToggleSpotlight()
    If bEnableSpotlight = True Then
        bEnableSpotlight = False
        ThisWorkbook.myApp.SpotlightOff
    Else
        bEnableSpotlight = True
    End If
End Sub
