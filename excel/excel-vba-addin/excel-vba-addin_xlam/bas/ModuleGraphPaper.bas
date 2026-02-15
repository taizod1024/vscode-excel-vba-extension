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
    
    ' Get default font name and size from Normal style
    defaultFontName = ActiveWorkbook.Styles("Normal").Font.Name
    defaultFontSize = ActiveWorkbook.Styles("Normal").Font.Size
    
    ' Set font to default
    ActiveSheet.Cells.Font.Name = defaultFontName
    
    ' Set font size to default
    ActiveSheet.Cells.Font.Size = defaultFontSize
    
    ' Set column width to 2
    ActiveSheet.Columns.ColumnWidth = 2
    
    MsgBox "Graph paper format applied.", vbInformation

End Sub