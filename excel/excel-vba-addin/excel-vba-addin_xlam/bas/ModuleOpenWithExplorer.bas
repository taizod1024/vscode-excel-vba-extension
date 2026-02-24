Attribute VB_Name = "ModuleOpenWithExplorer"
Option Explicit

' 定数定義
Const EXPLORER_COMMAND As String = "explorer.exe"

Sub OpenWithExplorer_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWindow Is Nothing)
End Sub

Sub OpenWithExplorer_onAction(control As IRibbonControl)
    OpenWithExplorer
End Sub

''' ================================================================================
''' サブルーチン: OpenWithExplorer
''' 説明: Explorerを起動（アクティブなワークブックのフォルダで）
''' 説明: Webから開いている場合はRecentフォルダから対応するファイルを探す
''' パラメータ: なし
''' 戻り値: なし
''' ================================================================================
Sub OpenWithExplorer()
    Dim folderPath As String
    Dim bookPath As String
    
    On Error GoTo ErrorHandler
    
    ' アクティブなワークブックが存在するか確認
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook open.", vbInformation
        Exit Sub
    End If
    
    ' ActiveWorkbook.FullName の値を取得
    bookPath = ActiveWorkbook.FullName
    
    ' Webから開いている場合はRecentフォルダから対応するファイルを探す
    bookPath = ResolveWebBookPath(bookPath, ActiveWorkbook.Name)
    If bookPath = "" Then
        Exit Sub
    End If
    
    ' ワークブックのパスからフォルダを取得
    folderPath = GetParentFolder(bookPath)
    
    If folderPath = "" Then
        MsgBox "Workbook not saved.", vbInformation
        Exit Sub
    End If
    
    ' Explorer でフォルダを開く
    Shell EXPLORER_COMMAND & " """ & folderPath & """", vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Failed to open Explorer: " & Err.description, vbExclamation
End Sub