Attribute VB_Name = "Module8"
Sub JungleSunrise()
Attribute JungleSunrise.VB_ProcData.VB_Invoke_Func = "m\n14"

Dim rng As Range
Dim MinNum As Double
Dim MaxNum As Double
Dim Blau As Integer
Dim Grun As Integer
Dim Rot As Integer
Dim Radian As Double
Dim RotFreq As Double
Dim RotPhase As Double
Dim GrunFreq As Double
Dim GrunPhase As Double
Dim BlauFreq As Double
Dim BlauPhase As Double
Dim RotSat As Boolean
Dim RotOff As Boolean
Dim GrunSat As Boolean
Dim GrunOff As Boolean
Dim BlauSat As Boolean
Dim BlauOff As Boolean

RotFreq = 50
GrunFreq = 50
BlauFreq = 100

RotPhase = 150
GrunPhase = 0
BlauPhase = 100

RotSat = False
GrunSat = False
BlauSat = False
RotOff = False
GrunOff = False
BlauOff = False


Set rng = Selection
MinNum = Application.WorksheetFunction.Min(rng)
MaxNum = Application.WorksheetFunction.Max(rng)

For Each cell In rng

    If IsNumeric(cell.Value) = False Or cell.Value = "" Then
    
        'Do nothing
        
    Else
    
        If MaxNum - MinNum = 0 Then
            Radian = 0.5
        Else
            Radian = (cell.Value - MinNum) * 0.5 / (MaxNum - MinNum)
        End If
        
        If RotSat Then
            Rot = 255
        ElseIf RotOff Then
            Rot = 0
        Else
            Rot = (Sin((Radian / (RotFreq / 100) + (RotPhase / 100)) * 3.1415926) + 1) * 255 / 2
        End If
        
        If GrunSat Then
            Grun = 255
        ElseIf GrunOff Then
            Grun = 0
        Else
            Grun = (Sin((Radian / (GrunFreq / 100) + (GrunPhase / 100)) * 3.1415926) + 1) * 255 / 2
        End If
        
        If BlauSat Then
            Blau = 255
        ElseIf BlauOff Then
            Blau = 0
        Else
            Blau = (Sin((Radian / (BlauFreq / 100) + (BlauPhase / 100)) * 3.1415926) + 1) * 255 / 2
        End If
        
        cell.Interior.Color = RGB(Rot, Grun, Blau)
    
    End If
    
Next cell


End Sub



