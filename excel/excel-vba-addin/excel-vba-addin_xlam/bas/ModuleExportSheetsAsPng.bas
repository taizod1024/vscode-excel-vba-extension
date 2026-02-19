Attribute VB_Name = "ModuleExportSheetsAsPng"
Option Explicit

' ================================================================================
' モジュール: ModuleExportSheetsAsPng
' 説明: PNG エクスポート機能（リボンボタンコールバック）
' ================================================================================

Sub ExportSheetsAsPng_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWindow Is Nothing)
End Sub

''' ================================================================================
''' サブルーチン: ExportSheetsAsPng_onAction (リボンコールバック)
''' 説明: リボンボタンから呼ばれるコールバック
''' 戻り値: なし
''' ================================================================================
Sub ExportSheetsAsPng_onAction(control As IRibbonControl)
    ExportSheetsAsPng
End Sub

''' ================================================================================
''' サブルーチン: ExportSheetsAsPng
''' 説明: .png で終わるシートを PNG にエクスポート
''' パラメータ: なし
''' 戻り値: なし
''' ================================================================================
Sub ExportSheetsAsPng()
    Dim shell As Object
    Dim fso As Object
    Dim bookPath As String
    Dim extensionPath As String
    Dim scriptPath As String
    Dim imageOutputPath As String
    Dim fileExt As String
    Dim command As String
    
    On Error GoTo ErrorHandler
    
    ' カーソルを砂時計に変更
    Application.Cursor = xlWait
    
    ' ワークブックの確認と初期化
    If ActiveWorkbook Is Nothing Then
        MsgBox "No workbook open.", vbInformation
        Exit Sub
    End If
    
    bookPath = ActiveWorkbook.FullName
    
    ' クラウドファイルの場合は Recent フォルダから検索
    If Left(bookPath, 7) = "http://" Or Left(bookPath, 8) = "https://" Then
        bookPath = GetRecentFilePath(ActiveWorkbook.Name & ".url")
        If bookPath = "" Then
            MsgBox "Recent file not found: " & ActiveWorkbook.Name & ".url", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Azure拡張機能のパスを取得
    extensionPath = GetExtensionPath()
    If extensionPath = "" Then
        MsgBox "Excel VBA Extension not found.", vbExclamation
        Exit Sub
    End If
    
    scriptPath = extensionPath & "\bin\Export-SheetAsPng.ps1"
    If Dir(scriptPath) = "" Then
        MsgBox "PowerShell script not found: " & scriptPath, vbExclamation
        Exit Sub
    End If
    
    ' 出力パスの構築
    fileExt = GetActualFileExtension(bookPath)
    imageOutputPath = GetParentFolder(bookPath) & "\" & GetActualFileNameWithoutExt(bookPath) & _
                      "_" & fileExt & "\png"
    
    ' PowerShell スクリプト実行
    Set shell = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File """ & _
              scriptPath & """ """ & bookPath & """ """ & imageOutputPath & """"
    shell.Run command, 0, True
    
    ' 出力フォルダをエクスプローラで開く
    ' 完了通知ダイアログを表示
    MsgBox "PNG export completed." & vbCrLf & "Folder: " & imageOutputPath, vbInformation, "Export Completed"
    
    OpenFolderInExplorer imageOutputPath
    
    ' カーソルを通常状態に戻す
    Application.Cursor = xlDefault
    
    Exit Sub
    
ErrorHandler:
    ' カーソルを通常状態に戻す
    Application.Cursor = xlDefault
    MsgBox "Failed to export sheets as PNG: " & Err.description, vbExclamation
End Sub