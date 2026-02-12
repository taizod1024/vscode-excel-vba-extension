Attribute VB_Name = "ModuleOpenVSCode"
Option Explicit

' ================================================================================
' モジュール: ModuleOpenVSCode
' 説明: VS Code起動機能
' ================================================================================

' 定数定義
Const ENV_USERPROFILE As String = "USERPROFILE"
Const VSCODE_EXTENSIONS_PATH As String = ".vscode\extensions\"
Const EXTENSION_PREFIX As String = "taizod1024.excel-vba-"
Const VSCODE_COMMAND As String = "code "

''' ================================================================================
''' 関数: OpenVSCode_getEnabled (リボンコールバック)
''' 説明: リボンボタンの有効/無効を制御
''' パラメータ: なし
''' 戻り値: Boolean - true で有効
''' ================================================================================
Sub OpenVSCode_getEnabled(control As IRibbonControl, ByRef enabled)
    enabled = True
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
''' パラメータ: なし
''' 戻り値: なし
''' ================================================================================
Sub OpenVSCode()
    Dim shell As Object
    Dim command As String
    Dim workbookFolderPath As String
    
    On Error GoTo ErrorHandler
    
    ' アクティブなワークブックが存在するか確認
    If ActiveWorkbook Is Nothing Then
        MsgBox "NO WORKBOOK", vbInformation
        Exit Sub
    End If
    
    ' ワークブックのパスからフォルダを取得
    workbookFolderPath = GetParentFolder(ActiveWorkbook.FullName)
    
    If workbookFolderPath = "" Then
        MsgBox "WORKBOOK NOT SAVED", vbInformation
        Exit Sub
    End If
    
    ' VS Code でフォルダを開く
    Set shell = CreateObject("WScript.Shell")
    command = VSCODE_COMMAND & """" & workbookFolderPath & """" & " """ & ActiveWorkbook.FullName & """"
    shell.Run command, 0, False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "VSCODE NOT OPEN: " & Err.description, vbExclamation
End Sub

''' ================================================================================
''' 関数: GetExtensionPath
''' 説明: VSCode拡張機能のパスを取得 (%USERPROFILE%\.vscode\extensions\taizod1024.excel-vba-*)
''' 説明: 存在しない場合はエラーメッセージを表示
''' パラメータ: なし
''' 戻り値: String - 拡張機能のパス、見つからない場合は空文字列
''' ================================================================================
Function GetExtensionPath() As String
    Dim userProfile As String
    Dim extensionsPath As String
    Dim foundPath As String
    
    On Error GoTo ErrorHandler
    
    ' %USERPROFILE% から .vscode\extensions パスを構築
    userProfile = Environ(ENV_USERPROFILE)
    
    If userProfile = "" Then
        MsgBox ENV_USERPROFILE & " 環境変数が取得できません", vbExclamation
        GetExtensionPath = ""
        Exit Function
    End If
    
    extensionsPath = userProfile & "\" & VSCODE_EXTENSIONS_PATH
    
    ' taizod1024.excel-vba-* パターンのフォルダを検索
    foundPath = FindExtensionFolder(extensionsPath)
    
    If foundPath = "" Then
        MsgBox "VSCode拡張機能フォルダが見つかりません:" & vbCrLf & _
               extensionsPath & EXTENSION_PREFIX & "*", vbExclamation
        GetExtensionPath = ""
    Else
        GetExtensionPath = foundPath
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "拡張機能パス取得エラー: " & Err.Description, vbExclamation
    GetExtensionPath = ""
End Function

''' ================================================================================
''' 関数: FindExtensionFolder
''' 説明: taizod1024.excel-vba-* パターンのフォルダを検索
''' パラメータ:
'''   extensionsPath As String - 拡張機能フォルダの親パス
''' 戻り値: String - 見つかったフォルダのパス、見つからない場合は空文字列
''' ================================================================================
Private Function FindExtensionFolder(extensionsPath As String) As String
    Dim fileSystemObj As Object
    Dim extensionsFolder As Object
    Dim folder As Object
    Dim folderName As String
    
    On Error GoTo ErrorHandler
    
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
    
    ' 拡張機能フォルダが存在するか確認
    If Not fileSystemObj.FolderExists(extensionsPath) Then
        FindExtensionFolder = ""
        Exit Function
    End If
    
    Set extensionsFolder = fileSystemObj.GetFolder(extensionsPath)
    
    ' フォルダを列挙して taizod1024.excel-vba-* パターンを検索
    For Each folder In extensionsFolder.SubFolders
        folderName = folder.Name
        
        ' EXTENSION_PREFIX で始まるかチェック
        If Left(folderName, Len(EXTENSION_PREFIX)) = EXTENSION_PREFIX Then
            FindExtensionFolder = folder.Path
            Exit Function
        End If
    Next folder
    
    ' 見つからなかった場合
    FindExtensionFolder = ""
    
    Exit Function
    
ErrorHandler:
    FindExtensionFolder = ""
End Function

''' ================================================================================
''' 関数: GetParentFolder
''' 説明: ファイルまたはフォルダのパスから親フォルダを取得
''' パラメータ:
'''   filePath As String - ファイルまたはフォルダのパス
''' 戻り値: String - 親フォルダのパス、取得できない場合は空文字列
''' ================================================================================
Private Function GetParentFolder(filePath As String) As String
    Dim lastSeparatorPos As Integer
    
    If filePath = "" Then
        GetParentFolder = ""
        Exit Function
    End If
    
    ' 最後のバックスラッシュを検索
    lastSeparatorPos = InStrRev(filePath, "\")
    
    If lastSeparatorPos > 0 Then
        GetParentFolder = Left(filePath, lastSeparatorPos - 1)
    Else
        GetParentFolder = ""
    End If
End Function