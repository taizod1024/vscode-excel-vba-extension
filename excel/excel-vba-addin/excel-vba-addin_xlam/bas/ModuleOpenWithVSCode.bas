Attribute VB_Name = "ModuleOpenWithVSCode"
Option Explicit

' ================================================================================
' モジュール: ModuleOpenWithVSCode
' 説明: VS Code起動機能
' ================================================================================

' 定数定義
Const VSCODE_COMMAND As String = "code.cmd"

''' ================================================================================
''' 関数: OpenWithVSCode_getEnabled (リボンコールバック)
''' 説明: リボンボタンの有効/無効を制御
''' パラメータ: なし
''' 戻り値: Boolean - true で有効
''' ================================================================================
Sub OpenWithVSCode_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWindow Is Nothing)
End Sub

''' ================================================================================
''' サブルーチン: OpenWithVSCode_onAction (リボンコールバック)
''' 説明: リボンボタンから呼ばれるコールバック
''' 戻り値: なし
''' ================================================================================
Sub OpenWithVSCode_onAction(control As IRibbonControl)
    OpenWithVSCode
End Sub

''' ================================================================================
''' サブルーチン: OpenWithVSCode
''' 説明: VS Codeを起動（アクティブなワークブックのフォルダで）
''' 説明: Webから開いている場合はRecentフォルダから対応するファイルを探す
''' パラメータ: なし
''' 戻り値: なし
''' ================================================================================
Sub OpenWithVSCode()
    Dim command As String
    Dim bookFolderPath As String
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
    bookFolderPath = GetParentFolder(bookPath)
    
    If bookFolderPath = "" Then
        MsgBox "Workbook not saved.", vbInformation
        Exit Sub
    End If
    
    ' VS Code でフォルダを開く
    command = VSCODE_COMMAND & " """ & bookFolderPath & """ """ & bookPath & """"
    Shell command, vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Failed to open VS Code: " & Err.description, vbExclamation
End Sub