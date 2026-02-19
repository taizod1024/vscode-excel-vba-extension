Attribute VB_Name = "ModuleLoadCSV"
Option Explicit

' ================================================================================
' モジュール: ModuleLoadCSV
' 説明: CSV ロード機能（リボンボタンコールバック）
' ================================================================================

Sub LoadCSV_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWorkbook Is Nothing)
End Sub

''' ================================================================================
''' サブルーチン: LoadCSV_onAction (リボンコールバック)
''' 説明: リボンボタンから呼ばれるコールバック
''' 戻り値: なし
''' ================================================================================
Sub LoadCSV_onAction(control As IRibbonControl)
    LoadCSV
End Sub

''' ================================================================================
''' サブルーチン: LoadCSV
''' 説明: CSV ファイルを Excel シートにロード
''' パラメータ: なし
''' 戻り値: なし
''' ================================================================================
Sub LoadCSV()
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
    
    scriptPath = extensionPath & "\bin\Load-CSV.ps1"
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
    
    Dim csvPath As String
    csvPath = GetParentFolder(bookPath) & "\" & baseName & "_" & fileExt & "\csv~"
    
    ' PowerShell スクリプト実行
    Set shell = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File """ & _
              scriptPath & """ """ & bookPath & """ """ & csvPath & """"
    
    Dim exitCode As Long
    exitCode = shell.Run(command, 0, True)
    
    ' 実行失敗時はエラー表示
    If exitCode <> 0 Then
        MsgBox "Error: PowerShell execution failed (Exit code: " & exitCode & ")", vbExclamation
        Application.Cursor = xlDefault
        Exit Sub
    End If
    
    ' 成功時：csv~ を csv にリネーム
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim finalPath As String
    finalPath = Left(csvPath, Len(csvPath) - 1)  ' csv~ から ~ を削除
    
    If fso.FolderExists(finalPath) Then
        fso.DeleteFolder finalPath
    End If
    fso.MoveFolder csvPath, finalPath
    
    ' 出力フォルダをエクスプローラで開く
    ' 完了通知ダイアログを表示
    MsgBox "CSV loaded successfully." & vbCrLf & "Folder: " & finalPath, vbInformation, "Load Completed"
    
    OpenFolderInExplorer finalPath
    
    ' カーソルを通常状態に戻す
    Application.Cursor = xlDefault
    
    Exit Sub
    
ErrorHandler:
    ' カーソルを通常状態に戻す
    Application.Cursor = xlDefault
    MsgBox "Failed to load CSV: " & Err.description, vbExclamation
End Sub