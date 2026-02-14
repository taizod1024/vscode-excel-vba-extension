Attribute VB_Name = "ModuleSampleMacro"
Option Explicit

Sub SampleMacro_getEnabled(control As IRibbonControl, ByRef enabled)

    enabled = True

End Sub

Sub SampleMacro_onAction(control As IRibbonControl)

    SampleMacro

End Sub

Sub SampleMacro

    Msgbox "SampleMacro"

End Sub