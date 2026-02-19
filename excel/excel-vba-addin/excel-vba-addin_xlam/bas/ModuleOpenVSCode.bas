Attribute VB_Name = "ModuleOpenVSCode"
Option Explicit

' ================================================================================
' モジュール: ModuleOpenVSCode
' 説明: VS Code起動機能
' ================================================================================

' 定数定義
Const VSCODE_COMMAND As String = "code "

''' ================================================================================
''' 関数: OpenVSCode_getEnabled (リボンコールバック)
''' 説明: リボンボタンの有効/無効を制御
''' パラメータ: なし
''' 戻り値: Boolean - true で有効
''' ================================================================================
Sub OpenVSCode_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Not (ActiveWindow Is Nothing)
End Sub

''' ================================================================================
''' サブルーチン: OpenVSCode_onAction (リボンコールバック)
''' 説明: リボンボタンから呼ばれるコールバック
''' 戻り値: なし
''' ================================================================================
Sub OpenVSCode_onAction(control As IRibbonControl)
    OpenVSCode
End Sub

''' ================================================================================
''' サブルーチン: OpenVSCode
''' 説明: VS Codeを起動（アクティブなワークブックのフォルダで）
''' 説明: Webから開いている場合はRecentフォルダから対応するファイルを探す
''' パラメータ: なし
''' 戻り値: なし
''' ================================================================================
Sub OpenVSCode()
    Dim shell As Object
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
    
    ' Webから開いている場合（URLの場合）は、Recentフォルダから.urlを探す
    If Left(bookPath, 7) = "http://" Or Left(bookPath, 8) = "https://" Then
        Dim originalUrl As String
        originalUrl = bookPath
        bookPath = GetRecentFilePath(ActiveWorkbook.Name & ".url")
        If bookPath = "" Then
            ' .urlファイルを作成する
            bookPath = CreateRecentUrlFile(ActiveWorkbook.Name & ".url", originalUrl)
            If bookPath = "" Then
                MsgBox "Failed to create recent file: " & ActiveWorkbook.Name & ".url", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    ' ワークブックのパスからフォルダを取得
    bookFolderPath = GetParentFolder(bookPath)
    
    If bookFolderPath = "" Then
        MsgBox "Workbook not saved.", vbInformation
        Exit Sub
    End If
    
    ' VS Code でフォルダを開く
    Set shell = CreateObject("WScript.Shell")
    command = VSCODE_COMMAND & """" & bookFolderPath & """" & " """ & bookPath & """"
    shell.Run command, 0, False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Failed to open VS Code: " & Err.description, vbExclamation
End Sub

''' ================================================================================
''' 関数: CreateRecentUrlFile
''' 説明: Recentフォルダに.urlファイルを作成
''' パラメータ:
'''   fileName As String - 作成するファイル名
'''   urlContent As String - URLの内容
''' 戻り値: String - 作成したファイルのパス、失敗した場合は空文字列
''' ================================================================================
Private Function CreateRecentUrlFile(fileName As String, urlContent As String) As String
    Dim shell As Object
    Dim userAppDataPath As String
    Dim recentPath As String
    Dim fileSystemObj As Object
    Dim filePath As String
    Dim fileHandle As Object
    
    On Error GoTo ErrorHandler
    
    ' APPDATAフォルダのパスを取得
    Set shell = CreateObject("WScript.Shell")
    userAppDataPath = shell.ExpandEnvironmentStrings("%APPDATA%")
    
    ' Recentフォルダのパスを構築
    recentPath = userAppDataPath & "\Microsoft\Office\Recent"
    filePath = recentPath & "\" & fileName
    
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
    
    ' Recentフォルダが存在しなければ作成
    If Not fileSystemObj.FolderExists(recentPath) Then
        fileSystemObj.CreateFolder recentPath
    End If
    
    ' .urlファイルを作成
    Set fileHandle = fileSystemObj.CreateTextFile(filePath, True)
    fileHandle.WriteLine "[InternetShortcut]"
    fileHandle.WriteLine "URL=" & urlContent
    fileHandle.Close
    
    CreateRecentUrlFile = filePath
    
    Exit Function
    
ErrorHandler:
    CreateRecentUrlFile = ""
End Function