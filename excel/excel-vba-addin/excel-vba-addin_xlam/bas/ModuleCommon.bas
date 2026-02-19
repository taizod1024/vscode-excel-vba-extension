Attribute VB_Name = "ModuleCommon"
Option Explicit

' ================================================================================
' モジュール: ModuleCommon
' 説明: 共通ユーティリティ関数
' ================================================================================

' 定数定義
Public Const ENV_USERPROFILE As String = "USERPROFILE"
Public Const VSCODE_EXTENSIONS_PATH As String = ".vscode\extensions\"
Public Const EXTENSION_PREFIX As String = "taizod1024.excel-vba-"

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
        MsgBox "Environment variable not found: " & ENV_USERPROFILE, vbExclamation
        GetExtensionPath = ""
        Exit Function
    End If
    
    extensionsPath = userProfile & "\" & VSCODE_EXTENSIONS_PATH
    
    ' taizod1024.excel-vba-* パターンのフォルダを検索
    foundPath = FindExtensionFolder(extensionsPath)
    
    If foundPath = "" Then
        MsgBox "Extension folder not found: " & vbCrLf & _
               extensionsPath & EXTENSION_PREFIX & "*", vbExclamation
        GetExtensionPath = ""
    Else
        GetExtensionPath = foundPath
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Failed to retrieve extension path: " & Err.description, vbExclamation
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
''' 関数: GetRecentFilePath
''' 説明: %APPDATA%\Microsoft\Office\Recentフォルダからファイルを検索
''' パラメータ:
'''   fileName As String - 検索するファイル名
''' 戻り値: String - 見つかったファイルのパス、見つからない場合は空文字列
''' ================================================================================
Function GetRecentFilePath(fileName As String) As String
    Dim shell As Object
    Dim userAppDataPath As String
    Dim recentPath As String
    Dim fileSystemObj As Object
    Dim recentFolder As Object
    Dim file As Object
    Dim foundPath As String
    
    On Error GoTo ErrorHandler
    
    ' APPDATAフォルダのパスを取得
    Set shell = CreateObject("WScript.Shell")
    userAppDataPath = shell.ExpandEnvironmentStrings("%APPDATA%")
    
    ' Recentフォルダのパスを構築
    recentPath = userAppDataPath & "\Microsoft\Office\Recent"
    
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
    
    ' Recentフォルダが存在するか確認
    If Not fileSystemObj.FolderExists(recentPath) Then
        GetRecentFilePath = ""
        Exit Function
    End If
    
    Set recentFolder = fileSystemObj.GetFolder(recentPath)
    
    ' ファイルを列挙して同じ名前のファイルを検索
    For Each file In recentFolder.Files
        If LCase(file.Name) = LCase(fileName) Then
            GetRecentFilePath = file.Path
            Exit Function
        End If
    Next file
    
    ' 見つからなかった場合
    GetRecentFilePath = ""
    
    Exit Function
    
ErrorHandler:
    GetRecentFilePath = ""
End Function

''' ================================================================================
''' 関数: GetParentFolder
''' 説明: ファイルまたはフォルダのパスから親フォルダを取得
''' パラメータ:
'''   filePath As String - ファイルまたはフォルダのパス
''' 戻り値: String - 親フォルダのパス、取得できない場合は空文字列
''' ================================================================================
Function GetParentFolder(filePath As String) As String
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

''' ================================================================================
''' 関数: GetFileNameWithoutExt
''' 説明: ファイルパスからファイル名（拡張子なし）を取得
''' パラメータ: filePath - ファイルの完全パス
''' 戻り値: String - ファイル名（拡張子なし）
''' ================================================================================
Function GetFileNameWithoutExt(filePath As String) As String
    Dim fileName As String
    Dim lastSlash As Long
    Dim lastDot As Long
    
    lastSlash = InStrRev(filePath, "\")
    If lastSlash > 0 Then
        fileName = Mid(filePath, lastSlash + 1)
    Else
        fileName = filePath
    End If
    
    lastDot = InStrRev(fileName, ".")
    If lastDot > 0 Then
        GetFileNameWithoutExt = Left(fileName, lastDot - 1)
    Else
        GetFileNameWithoutExt = fileName
    End If
End Function

''' ================================================================================
''' 関数: GetActualFileExtension
''' 説明: ファイルパスから実際の拡張子を取得（.url の場合は除去）
''' パラメータ: filePath - ファイルの完全パス
''' 戻り値: String - ファイルの拡張子（ドット抜き）
''' ================================================================================
Function GetActualFileExtension(filePath As String) As String
    Dim actualPath As String
    
    ' .url の場合は除去して実際の拡張子を取得
    If Right(filePath, 4) = ".url" Then
        actualPath = Left(filePath, Len(filePath) - 4)
    Else
        actualPath = filePath
    End If
    
    ' ドットの位置から拡張子を抽出
    GetActualFileExtension = Mid(actualPath, InStrRev(actualPath, ".") + 1)
End Function

''' ================================================================================
''' 関数: GetActualFileNameWithoutExt
''' 説明: ファイルパスからファイル名（拡張子なし）を取得（.url の場合は除去）
''' パラメータ: filePath - ファイルの完全パス
''' 戻り値: String - ファイル名（拡張子なし）
''' ================================================================================
Function GetActualFileNameWithoutExt(filePath As String) As String
    Dim actualPath As String
    
    ' .url の場合は除去
    If Right(filePath, 4) = ".url" Then
        actualPath = Left(filePath, Len(filePath) - 4)
    Else
        actualPath = filePath
    End If
    
    ' ファイル名から拡張子を除去
    GetActualFileNameWithoutExt = GetFileNameWithoutExt(actualPath)
End Function

''' ================================================================================
''' サブルーチン: OpenFolderInExplorer
''' 説明: フォルダをエクスプローラで開く（既に開いている場合はアクティベート、
'''        フォルダが存在しない場合は親フォルダを開く）
''' パラメータ:
'''   folderPath As String - 開くフォルダのパス
''' 戻り値: なし
''' ================================================================================
Sub OpenFolderInExplorer(folderPath As String)
    Dim shell As Object
    Dim fso As Object
    Dim parentFolder As String
    Dim windows As Object
    Dim window As Object
    Dim found As Boolean
    
    On Error GoTo ErrorHandler
    
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    found = False
    
    ' 指定されたフォルダが存在する場合
    If fso.FolderExists(folderPath) Then
        ' Shell.Application で既に開いているエクスプローラを検索
        Set windows = CreateObject("Shell.Application").Windows
        On Error Resume Next
        For Each window In windows
            If window.Document.Folder.Self.Path = folderPath Then
                ' 既に開いている場合はアクティベート
                window.Activate
                found = True
                Exit For
            End If
        Next window
        On Error GoTo ErrorHandler
        
        ' 開いている場合は処理終了
        If found Then Exit Sub
        
        ' 見つからないので新規で開く
        shell.Run "explorer """ & folderPath & """", 1, False
    Else
        ' 親フォルダが存在する場合は親フォルダを開く
        parentFolder = GetParentFolder(folderPath)
        If parentFolder <> "" Then
            ' 親フォルダも同様にチェック
            Set windows = CreateObject("Shell.Application").Windows
            On Error Resume Next
            For Each window In windows
                If window.Document.Folder.Self.Path = parentFolder Then
                    window.Activate
                    found = True
                    Exit For
                End If
            Next window
            On Error GoTo ErrorHandler
            
            ' 見つからないので新規で開く
            If Not found Then
                shell.Run "explorer """ & parentFolder & """", 1, False
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' エラー時は通常のexplorerを開く
    If fso.FolderExists(folderPath) Then
        shell.Run "explorer """ & folderPath & """", 1, False
    Else
        parentFolder = GetParentFolder(folderPath)
        If parentFolder <> "" Then
            shell.Run "explorer """ & parentFolder & """", 1, False
        End If
    End If
End Sub