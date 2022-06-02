Attribute VB_Name = "Module2"
Sub BubblegumUnicorn()
Attribute BubblegumUnicorn.VB_ProcData.VB_Invoke_Func = " \n14"

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
            Radian = 5.4
        Else
            Radian = (cell.Value - MinNum) * 5.4 / (MaxNum - MinNum)
        End If
        
        Select Case Radian
            Case 0 To 1
                Rot = 255
                Blau = (-Radian + 1) * 63.75 + 191.25
                Grun = 191.25
            Case 1 To 2
                Rot = 255
                Grun = (Radian - 1) * 63.75 + 191.25
                Blau = 191.25
            Case 2 To 3
                Grun = 255
                Rot = (-Radian + 3) * 63.75 + 191.25
                Blau = 191.25
            Case 3 To 4
                Grun = 255
                Blau = (Radian - 3) * 63.75 + 191.25
                Rot = 191.25
            Case 4 To 5
                Blau = 255
                Grun = (-Radian + 5) * 63.75 + 191.25
                Rot = 191.25
            Case 5 To 6
                Blau = 255
                Rot = (Radian - 5) * 63.75 + 191.25
                Grun = 191.25
        End Select
        
        cell.Interior.Color = RGB(Rot, Grun, Blau)
        
    End If
    
Next cell


End Sub
