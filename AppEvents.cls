VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Dim lastTarget As Range
Dim dict As Object


Private Sub Class_Initialize()
    Set dict = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set dict = Nothing
    Set lastTarget = Nothing
End Sub

Public Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If bEnableSpotlight Then
        Dim cell As Range
        Dim lastColor As Long
        Dim rng As Range
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        ' 恢复之前选中行和列的颜色
        If Not lastTarget Is Nothing Then
            lastTarget.EntireRow.Interior.Color = xlNone
            lastTarget.EntireColumn.Interior.Color = xlNone
            
            Dim rowIntersect As Range
            Dim colIntersect As Range
            
            Set rowIntersect = Intersect(lastTarget.EntireRow, lastTarget.Parent.UsedRange)
            Set colIntersect = Intersect(lastTarget.EntireColumn, lastTarget.Parent.UsedRange)
            
            If Not rowIntersect Is Nothing Then
                For Each cell In rowIntersect
                    If dict.Exists(cell.Address) Then
                        cell.Interior.Color = dict(cell.Address)
                    End If
                Next cell
            End If
            
            If Not colIntersect Is Nothing Then
                For Each cell In colIntersect
                    If dict.Exists(cell.Address) Then
                        cell.Interior.Color = dict(cell.Address)
                    End If
                Next cell
            End If
        End If
    
        Set lastTarget = Target
        dict.RemoveAll
    
        ' 仅存储有颜色的单元格
        Set rng = Intersect(Target.EntireRow, Sh.UsedRange)
        If Not rng Is Nothing Then
            For Each cell In rng
                lastColor = cell.Interior.Color
                If lastColor <> 16777215 Then
                    dict(cell.Address) = lastColor
                End If
            Next cell
        End If
    
        Set rng = Intersect(Target.EntireColumn, Sh.UsedRange)
        If Not rng Is Nothing Then
            For Each cell In rng
                lastColor = cell.Interior.Color
                If lastColor <> 16777215 Then
                    dict(cell.Address) = lastColor
                End If
            Next cell
        End If
    
        ' 应用新颜色
        Target.EntireRow.Interior.ColorIndex = 36
        Target.EntireColumn.Interior.ColorIndex = 36
    End If
        Application.EnableEvents = True
        Application.ScreenUpdating = True
End Sub

Public Sub SpotlightOff()
    '恢复颜色
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 恢复之前选中行和列的颜色
    If Not lastTarget Is Nothing Then
        lastTarget.EntireRow.Interior.Color = xlNone
        lastTarget.EntireColumn.Interior.Color = xlNone
        
        Dim rowIntersect As Range
        Dim colIntersect As Range
        
        Set rowIntersect = Intersect(lastTarget.EntireRow, lastTarget.Parent.UsedRange)
        Set colIntersect = Intersect(lastTarget.EntireColumn, lastTarget.Parent.UsedRange)
        
        If Not rowIntersect Is Nothing Then
            For Each cell In rowIntersect
                If dict.Exists(cell.Address) Then
                    cell.Interior.Color = dict(cell.Address)
                End If
            Next cell
        End If
        
        If Not colIntersect Is Nothing Then
            For Each cell In colIntersect
                If dict.Exists(cell.Address) Then
                    cell.Interior.Color = dict(cell.Address)
                End If
            Next cell
        End If
    End If

    'Set lastTarget = Target
    dict.RemoveAll
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
