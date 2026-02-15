Attribute VB_Name = "ModuleSnapToGrid"
Option Explicit

Sub SnapToGrid_onAction(constrol As IRibbonControl)

    SnapToGrid

End Sub

Sub SnapToGrid_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = Not (ActiveWindow Is Nothing)
    
End Sub

Sub SnapToGrid()

    Dim oBefore As Object   ' セレクションの保存用
    Dim lColMax As Long     ' シートの最大列数
    Dim lRowMax As Long     ' シートの最大行数
    Dim dblXMax As Double   ' X軸の最大値
    Dim dblYMax As Double   ' Y軸の最大値
    Dim dblWork As Double   ' X軸・Y軸計算用
    Dim lIdxCol As Long     ' 列列挙用
    Dim lIdxRow As Long     ' 行列挙用
    Dim dblXArr() As Double ' X軸の配列
    Dim dblYArr() As Double ' Y軸の配列
    Dim bTgt As Boolean     ' 処理中の図形が操作対象かどうかのフラグ
    Dim bChg As Boolean     ' 処理中の図形を操作したどうかのフラグ
    Dim lCnt As Long        ' 図形の個数
    Dim lShp As Long        ' 操作対象の図形の個数
    Dim lChg As Long        ' 修正した図形の個数
    Dim oShp As Shape       ' 操作対象の図形
    Dim oShpRng As ShapeRange ' 操作対象の図形全体
    Dim sInfo As String     ' メッセージ
    
    ' SnapToGrid On
    If Not Application.CommandBars.GetPressedMso("SnapToGrid") Then  
        Application.CommandBars.ExecuteMso "SnapToGrid"
    End If
    
    ' 実行前チェック
    If TypeName(Selection) <> "Range" Then
        Set oShpRng = Selection.ShapeRange
        ' 実行確認
        Select Case MsgBox("This operation cannot be undone." + vbCrLf + "Snap " & oShpRng.Count & " selected shape(s) to grid?", _
            vbOKCancel + vbExclamation)
        Case vbOK
        Case vbCancel
            Exit Sub
        End Select
    Else
        ' 実行確認
        Select Case MsgBox("This operation cannot be undone." + vbCrLf + "Snap all shapes to grid?", _
            vbOKCancel + vbExclamation)
        Case vbOK
        Case vbCancel
            Exit Sub
        End Select
        ' 選択して図形があることを確認
        If ActiveSheet.Shapes.Count = 0 Then
            MsgBox "No shapes found.", vbInformation
            Exit Sub
        End If
        ActiveSheet.Shapes.SelectAll
        Set oShpRng = Selection.ShapeRange
    End If
    
    ' 画面描画停止
    Application.ScreenUpdating = False
    
    ' セルの列数・行数の取得
    Set oBefore = Selection
    Cells.Select
    lColMax = Cells.Columns.Count
    lRowMax = Cells.Rows.Count
    If TypeOf oBefore Is Range Then
        oBefore.Select
    Else
        ActiveWindow.VisibleRange(1, 1).Select
    End If
    
    ' 画面描画再開
    Application.ScreenUpdating = True
    
    ' X軸の最大値・Y軸の最大値の取得
    dblXMax = 0
    dblYMax = 0
    For lCnt = 1 To oShpRng.Count
        Set oShp = oShpRng.Item(lCnt)
        dblWork = oShp.Left + oShp.Width
        If dblXMax < dblWork Then dblXMax = dblWork
        dblWork = oShp.Top + oShp.Height
        If dblYMax < dblWork Then dblYMax = dblWork
    Next
        
    ' 列単位に横位置の取得・行単位に縦位置の取得
    lIdxCol = 0
    Do
        lIdxCol = lIdxCol + 1
        ReDim Preserve dblXArr(lIdxCol) As Double
        dblWork = ActiveSheet.Cells(1, lIdxCol).Left
        dblXArr(lIdxCol) = dblWork
    Loop While dblWork < dblXMax And lIdxCol < lColMax
    
    lIdxRow = 0
    Do
        lIdxRow = lIdxRow + 1
        ReDim Preserve dblYArr(lIdxRow) As Double
        dblWork = ActiveSheet.Cells(lIdxRow, 1).Top
        dblYArr(lIdxRow) = dblWork
    Loop While dblWork < dblYMax And lIdxRow < lRowMax
    
    ' すべての図形に対して近いグリッド線に寄せる
    lShp = 0
    lChg = 0
    For lCnt = 1 To oShpRng.Count
        
        Set oShp = oShpRng.Item(lCnt)
        
        ' メッセージ
        Application.StatusBar = "Adjusting " + CStr(lCnt) + " of " + CStr(oShpRng.Count) + " shapes ..."
        DoEvents
        
        bChg = False
        ' 操作対象かどうかを判断
        If oShp.Connector Then
            ' コネクタなら...
            ' 両方とも繋がっていないこと
            bTgt = _
                Not oShp.ConnectorFormat.BeginConnected And _
                Not oShp.ConnectorFormat.EndConnected
        Else
            ' コネクタでないなら...
            ' コメントではないこと
            bTgt = _
               oShp.Type <> msoComment
        End If
        ' 操作対象なら...
        If bTgt Then
            lShp = lShp + 1
            With oShp
            
                ' Leftの調整
                For lIdxCol = 1 To UBound(dblXArr) - 1
                    If dblXArr(lIdxCol) = .Left Then
                        Exit For
                    ElseIf dblXArr(lIdxCol) < .Left And .Left < dblXArr(lIdxCol + 1) Then
                        bChg = True
                        If .Left < (dblXArr(lIdxCol) + dblXArr(lIdxCol + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Left = dblXArr(lIdxCol)
                        Else
                            .LockAspectRatio = False
                            .Left = dblXArr(lIdxCol + 1)
                        End If
                        Exit For
                    End If
                Next
                dblWork = .Left
                
                ' Widthの調整
                For lIdxCol = lIdxCol To UBound(dblXArr) - 1
                    If dblXArr(lIdxCol) = .Left + .Width Then
                        Exit For
                    ElseIf dblXArr(lIdxCol) < .Left + .Width And .Left + .Width < dblXArr(lIdxCol + 1) Then
                        bChg = True
                        If .Left + .Width < (dblXArr(lIdxCol) + dblXArr(lIdxCol + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Width = dblXArr(lIdxCol) - .Left
                        Else
                            .LockAspectRatio = False
                            .Width = dblXArr(lIdxCol + 1) - .Left
                        End If
                        Exit For
                    End If
                Next

                ' Leftの再調整
                ' 何故かWidthを調整したときに、Leftが微妙にずれる場合があるため...
                .Left = dblWork

                ' Topの調整
                For lIdxRow = 1 To UBound(dblYArr) - 1
                    If dblYArr(lIdxRow) = .Top Then
                        Exit For
                    ElseIf dblYArr(lIdxRow) < .Top And .Top < dblYArr(lIdxRow + 1) Then
                        bChg = True
                        If .Top < (dblYArr(lIdxRow) + dblYArr(lIdxRow + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Top = dblYArr(lIdxRow)
                        Else
                            .LockAspectRatio = False
                            .Top = dblYArr(lIdxRow + 1)
                        End If
                        Exit For
                    End If
                Next
                dblWork = .Top
                
                ' Heightの調整
                For lIdxRow = lIdxRow To UBound(dblYArr) - 1
                    If dblYArr(lIdxRow) - .Top = .Height Then
                        Exit For
                    ElseIf dblYArr(lIdxRow) - .Top < .Height And .Height < dblYArr(lIdxRow + 1) - .Top Then
                        bChg = True
                        If .Top + .Height < (dblYArr(lIdxRow) + dblYArr(lIdxRow + 1)) / 2 Then
                            .LockAspectRatio = False
                            .Height = dblYArr(lIdxRow) - .Top
                        Else
                            .LockAspectRatio = False
                            .Height = dblYArr(lIdxRow + 1) - .Top
                        End If
                        Exit For
                    End If
                Next
                
                ' Topの再調整
                ' 何故かHeightを調整したときに、Leftが微妙にずれる場合があるため...
                .Top = dblWork
                
                ' 変更されていたら...
                If bChg Then
                    If lChg = 0 Then
                        ActiveSheet.Cells(lIdxRow, lIdxCol).Activate
                    End If
                    ActiveSheet.Shapes(.Name).Select False
                    sInfo = "Left: " & CStr(.Left) & " Top: " & CStr(.Top) & " Width: " & CStr(.Width) & " Height: " & CStr(.Height)
                    lChg = lChg + 1
                End If
            End With
        End If
    Next
    
    ' メッセージ
    Application.StatusBar = ""
    
    ' 最後のメッセージ
    If lShp = 0 Then
        MsgBox "No shapes found.", vbInformation
    ElseIf lChg <> 0 Then
        MsgBox CStr(lChg) & " shape(s) snapped to grid." & vbCrLf & sInfo, vbInformation
    Else
        MsgBox "All shapes are already snapped to grid.", vbInformation
    End If
    
End Sub