Attribute VB_Name = "ModuleSample"
Option Explicit

Sub Sample_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = True

End Sub

Sub Sample_onAction(control As IRibbonControl)

    Sample

End Sub

Sub Sample

    Msgbox "Sample"

End Sub