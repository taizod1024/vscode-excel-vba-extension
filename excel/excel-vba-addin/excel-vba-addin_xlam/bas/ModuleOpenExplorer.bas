Attribute VB_Name = "ModuleOpenExplorer"
Option Explicit

Sub OpenExplorer_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWindow Is Nothing)
End Sub

Sub OpenExplorer_onAction(constrol As IRibbonControl)
    OpenExplorer
End Sub

Sub OpenExplorer()
    Dim folderPath As String    
    folderPath = ThisWorkbook.Path
    shell "explorer.exe " & Chr(34) & folderPath & Chr(34), vbNormalFocus
End Sub