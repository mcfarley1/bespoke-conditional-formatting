Attribute VB_Name = "Module1"
Sub ClearFormat()
Attribute ClearFormat.VB_ProcData.VB_Invoke_Func = "g\n14"

With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With


End Sub
