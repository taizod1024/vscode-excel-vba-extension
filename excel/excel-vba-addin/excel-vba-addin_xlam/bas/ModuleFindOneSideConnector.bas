Attribute VB_Name = "ModuleFindOneSideConnector"
Option Explicit

Sub FindOneSideConnector_onAction(constrol As IRibbonControl)

    FindOneSideConnector

End Sub

Sub FindOneSideConnector_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = Not (ActiveWindow Is Nothing)

End Sub

Sub FindOneSideConnector()

    Dim shp As Shape    ' shape
    Dim flg As Boolean  ' flag
    
    ' Loop through all shapes in the sheet
    For Each shp In ActiveSheet.Shapes
        
        If shp.Connector Then
        
            flg = False
            
            ' Check for one-side connector
            If shp.ConnectorFormat.BeginConnected And Not shp.ConnectorFormat.EndConnected Then flg = True
            If shp.ConnectorFormat.EndConnected And Not shp.ConnectorFormat.BeginConnected Then flg = True
            
            ' If one-side connector found, select it and exit
            If flg Then
                shp.Select
                Exit Sub
            End If
            
        End If
        
    Next
    
    ' No one-side connectors found
    MsgBox "No one-side connectors found.", vbInformation
    Exit Sub
    
End Sub