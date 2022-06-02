Attribute VB_Name = "Module7"
Sub RainbowSuperBrite()

Dim MinNum As Double
Dim MaxNum As Double
Dim Blau As Integer
Dim Grun As Integer
Dim Rot As Integer
Dim Radian As Double

Set rng = Selection
MinNum = Application.WorksheetFunction.Min(rng)
MaxNum = Application.WorksheetFunction.Max(rng)

For Each cell In rng

    If IsNumeric(cell.Value) = False Or cell.Value = "" Then
    
        'Do nothing
        
    Else

    
        If MaxNum - MinNum = 0 Then
            Radian = 4.5
        Else
            Radian = (cell.Value - MinNum) * 4.5 / (MaxNum - MinNum)
        End If
        
        Select Case Radian
            Case 0 To 1
                Rot = 255
                Grun = Radian * 255
                Blau = 0
            Case 1 To 2
                Grun = 255
                Rot = (-Radian + 2) * 255
                Blau = 0
            Case 2 To 3
                Grun = 255
                Blau = (Radian - 2) * 255
                Rot = 0
            Case 3 To 4
                Blau = 255
                Grun = (-Radian + 4) * 255
                Rot = 0
            Case 4 To 5
                Blau = 255
                Rot = (Radian - 4) * 255
                Grun = 0
            Case 5 To 6
                Rot = 255
                Blau = (-Radian + 6) * 255 / 2
                Grun = 0
        End Select
        
        cell.Interior.Color = RGB(Rot, Grun, Blau)
        
    End If
    
Next cell


End Sub

