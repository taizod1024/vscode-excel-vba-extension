Attribute VB_Name = "ModuleTableInsert"
Option Explicit

Sub TableInsert_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWindow Is Nothing)
End Sub

Sub TableInsertColumnsLeft_onAction(control As IRibbonControl)
    On Error Resume Next
    Dim colCount As Long
    colCount = Selection.Columns.Count
    Selection.EntireColumn.Insert CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Offset(0, colCount).EntireColumn.Select
    On Error GoTo 0
End Sub

Sub TableInsertColumnsRight_onAction(control As IRibbonControl)
    On Error Resume Next
    Dim rng As Range
    Set rng = Selection.Offset(0, Selection.Columns.Count)
    rng.EntireColumn.Insert CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.EntireColumn.Select
    On Error GoTo 0
End Sub

Sub TableInsertRowsAbove_onAction(control As IRibbonControl)
    On Error Resume Next
    Dim rowCount As Long
    rowCount = Selection.Rows.Count
    Selection.EntireRow.Insert CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Offset(rowCount, 0).EntireRow.Select
    On Error GoTo 0
End Sub

Sub TableInsertRowsBelow_onAction(control As IRibbonControl)
    On Error Resume Next
    Dim rng As Range
    Set rng = Selection.Offset(Selection.Rows.Count, 0)
    rng.EntireRow.Insert CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.EntireRow.Select
    On Error GoTo 0
End Sub

Sub TableDeleteColumns_onAction(control As IRibbonControl)
    On Error Resume Next
    Selection.EntireColumn.Delete
    Selection.EntireColumn.Select
    On Error GoTo 0
End Sub

Sub TableDeleteRows_onAction(control As IRibbonControl)
    On Error Resume Next
    Selection.EntireRow.Delete
    Selection.EntireRow.Select
    On Error GoTo 0
End Sub

Sub CopyToNewSheet_onAction(control As IRibbonControl)
    On Error Resume Next
    Dim srcRange As Range
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim sheetNum As Long
    Dim wb As Workbook
    Dim i As Long
    Dim srcSheet As Worksheet
    
    Set srcRange = Selection
    Set srcSheet = srcRange.Worksheet
    Set wb = ActiveWorkbook
    
    ' Find next available Sheet*.png name
    sheetNum = 1
    Do
        sheetName = "Sheet" & sheetNum & ".png"
        Dim exists As Boolean
        exists = False
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            If ws.Name = sheetName Then
                exists = True
                Exit For
            End If
        Next ws
        If Not exists Then Exit Do
        sheetNum = sheetNum + 1
    Loop
    
    ' Create new sheet at the end
    Set newSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    newSheet.Name = sheetName
    
    ' Copy column widths from source range to new sheet
    For i = 1 To srcRange.Columns.Count
        newSheet.Columns(i).ColumnWidth = srcSheet.Columns(srcRange.Column + i - 1).ColumnWidth
    Next i
    
    ' Copy row heights from source range to new sheet
    For i = 1 To srcRange.Rows.Count
        newSheet.Rows(i).RowHeight = srcSheet.Rows(srcRange.Row + i - 1).RowHeight
    Next i
    
    ' Copy selection to new sheet (including shapes)
    srcRange.Copy
    newSheet.Activate
    newSheet.Range("A1").Select
    newSheet.Paste
    Application.CutCopyMode = False
    
    ' Select A1
    newSheet.Range("A1").Select
    
    On Error GoTo 0
End Sub