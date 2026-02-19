Attribute VB_Name = "ModuleSaveVBA"
Option Explicit

' ================================================================================
' モジュール: ModuleSaveVBA
' 説明: VBA セーブ機能（リボンボタンコールバック）
' ================================================================================

Sub SaveVBA_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWorkbook Is Nothing)
End Sub

''' ================================================================================
''' サブルーチン: SaveVBA_onAction (リボンコールバック)
''' 説明: リボンボタンから呼ばれるコールバック
''' 戻り値: なし
''' ================================================================================
Sub SaveVBA_onAction(control As IRibbonControl)
    SaveVBA
End Sub

''' ================================================================================
''' サブルーチン: SaveVBA
''' 説明: VBA を Excel ワークブックにセーブ
''' パラメータ: なし
''' 戻り値: なし
''' ================================================================================
Sub SaveVBA()
    Dim shell As Object
    Dim bookPath As String
    Dim extensionPath As String
    Dim scriptPath As String
    Dim command As String
    
    On Error GoTo ErrorHandler
    
    ' カーソルを砂時計に変更
    Application.Cursor = xlWait
    
    ' ワークブックの確認と初期化
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook open.", vbInformation
        Application.Cursor = xlDefault
        Exit Sub
    End If
    
    bookPath = ActiveWorkbook.FullName
    
    ' クラウドファイルの場合は Recent フォルダから検索
    If Left(bookPath, 7) = "http://" Or Left(bookPath, 8) = "https://" Then
        bookPath = GetRecentFilePath(ActiveWorkbook.Name & ".url")
        If bookPath = "" Then
            MsgBox "Recent file not found: " & ActiveWorkbook.Name & ".url", vbExclamation
            Application.Cursor = xlDefault
            Exit Sub
        End If
    End If
    
    ' Azure拡張機能のパスを取得
    extensionPath = GetExtensionPath()
    If extensionPath = "" Then
        MsgBox "Excel VBA Extension not found.", vbExclamation
        Application.Cursor = xlDefault
        Exit Sub
    End If
    
    scriptPath = extensionPath & "\bin\Save-VBA.ps1"
    If Dir(scriptPath) = "" Then
        MsgBox "PowerShell script not found: " & scriptPath, vbExclamation
        Application.Cursor = xlDefault
        Exit Sub
    End If
    
    ' 出力パスの構築
    Dim fileExt As String
    fileExt = GetActualFileExtension(bookPath)
    Dim baseName As String
    baseName = GetActualFileNameWithoutExt(bookPath)
    
    Dim saveSourcePath As String
    saveSourcePath = GetParentFolder(bookPath) & "\" & baseName & "_" & fileExt & "\bas"
    
    ' PowerShell スクリプト実行
    Set shell = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File """ & _
              scriptPath & """ """ & bookPath & """ """ & saveSourcePath & """"
    shell.Run command, 0, True
    
    ' 完了通知ダイアログを表示
    MsgBox "VBA saved successfully." & vbCrLf & "Folder: " & saveSourcePath, vbInformation, "Save Completed"
    
    ' カーソルを通常状態に戻す
    Application.Cursor = xlDefault
    
    Exit Sub
    
ErrorHandler:
    ' カーソルを通常状態に戻す
    Application.Cursor = xlDefault
    MsgBox "Failed to save VBA: " & Err.description, vbExclamation
End Sub