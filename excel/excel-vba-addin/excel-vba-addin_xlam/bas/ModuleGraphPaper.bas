Attribute VB_Name = "ModuleGraphPaper"
Option Explicit

Sub GraphPaper_onAction(control As IRibbonControl)

    GraphPaper

End Sub

Sub GraphPaper_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = Not (ActiveWindow Is Nothing)

End Sub

Sub GraphPaper()

    Dim defaultFontName As String
    Dim defaultFontSize As Integer
    
    ' SnapToGrid On
    If Not Application.CommandBars.GetPressedMso("SnapToGrid") Then  
        Application.CommandBars.ExecuteMso "SnapToGrid"
    End If
    
    ' Get default font name and size from Normal style
    defaultFontName = Application.StandardFont
    defaultFontSize = Application.StandardFontSize
    
    ' Set font to default
    ActiveSheet.Cells.Font.Name = defaultFontName
    
    ' Set font size to default
    ActiveSheet.Cells.Font.Size = defaultFontSize
    
    ' Set column width to 2
    ActiveSheet.Columns.ColumnWidth = 2
    
    MsgBox "Graph paper format applied.", vbInformation

End Sub