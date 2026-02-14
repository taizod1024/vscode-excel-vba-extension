Attribute VB_Name = "ModuleSampleAddin"
Option Explicit

Sub SampleAddin_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = True

End Sub

Sub SampleAddin_onAction(control As IRibbonControl)

    SampleAddin

End Sub

Sub SampleAddin

    Msgbox "SampleAddin"

End Sub