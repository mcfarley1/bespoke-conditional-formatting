VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CondFormatMaker 
   Caption         =   "Conditional Format Parameters"
   ClientHeight    =   4314
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   12018
   OleObjectBlob   =   "CondFormatMaker_Studio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CondFormatMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ScrollBar1_Change()

End Sub

Private Sub BFSpin_Change()

BlueFreq.Value = BFSpin.Value

End Sub

Private Sub BlueFreq_Change()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    If BlueFreq = "" Then
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        BlueFreq = Sheet2.Range("K2")
    End If
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub BlueOff_Click()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub BluePhase_Change()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    If BluePhase = "" Then
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        BluePhase = Sheet2.Range("K3")
    End If
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub BlueSat_Click()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub BPSpin_Change()

BluePhase.Value = BPSpin.Value

End Sub

Private Sub Cancel_Click()

Unload Me

End Sub

Private Sub Clear_Click()
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub GFSpin_Change()

GreenFreq.Value = GFSpin.Value

End Sub

Private Sub GPSpin_Change()

GreenPhase.Value = GPSpin.Value

End Sub

Private Sub GreenFreq_Change()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    If GreenFreq = "" Then
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        GreenFreq = Sheet2.Range("H2")
    End If
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub GreenOff_Click()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub GreenPhase_Change()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    If GreenPhase = "" Then
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        GreenPhase = Sheet2.Range("H3")
    End If
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub GreenSat_Click()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub RedFreq_Change()

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


If IsNumeric(RedFreq) = False Then
    If RedFreq = "" Then
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        RedFreq = Sheet2.Range("E2")
    End If
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub RedOff_Click()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub RedPhase_Change()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    If RedPhase = "" Then
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        RedPhase = Sheet2.Range("E3")
    End If
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub RedSat_Click()

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


If IsNumeric(RedFreq) = False Then
    RedFreq = Sheet2.Range("E2")
End If

If IsNumeric(GreenFreq) = False Then
    GreenFreq = Sheet2.Range("H2")
End If

If IsNumeric(BlueFreq) = False Then
    BlueFreq = Sheet2.Range("K2")
End If

If IsNumeric(RedPhase) = False Then
    RedPhase = Sheet2.Range("E3")
End If

If IsNumeric(GreenPhase) = False Then
    GreenPhase = Sheet2.Range("H3")
End If

If IsNumeric(BluePhase) = False Then
    BluePhase = Sheet2.Range("K3")
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Private Sub RFSpin_Change()

RedFreq.Value = RFSpin.Value

End Sub

Private Sub RPSpin_Change()

RedPhase.Value = RPSpin.Value

End Sub

Private Sub Spiral_Click()
       

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


If IsNumeric(RedFreq) = False Then
    If RedFreq = "" Then
        MsgBox ("Do not forget to enter a non-zero number for the red frequency divider.")
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        RedFreq = Sheet2.Range("E2")
        Exit Sub
    End If
End If

If IsNumeric(GreenFreq) = False Then
    If GreenFreq = "" Then
        MsgBox ("Do not forget to enter a non-zero number for the green frequency divider.")
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        GreenFreq = Sheet2.Range("H2")
        Exit Sub
    End If
End If

If IsNumeric(BlueFreq) = False Then
    If BlueFreq = "" Then
        MsgBox ("Do not forget to enter a non-zero number for the blue frequency divider.")
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        BlueFreq = Sheet2.Range("K2")
        Exit Sub
    End If
End If

If IsNumeric(RedPhase) = False Then
    If RedPhase = "" Then
        MsgBox ("Do not forget to enter a number for the red phase shift.")
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        RedPhase = Sheet2.Range("E3")
        Exit Sub
    End If
End If

If IsNumeric(GreenPhase) = False Then
    If GreenPhase = "" Then
        MsgBox ("Do not forget to enter a number for the green phase shift.")
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        GreenPhase = Sheet2.Range("H3")
        Exit Sub
    End If
End If

If IsNumeric(BluePhase) = False Then
    If BluePhase = "" Then
        MsgBox ("Do not forget to enter a number for the blue phase shift.")
        Exit Sub
    Else
        MsgBox ("Only numbers are accepted as frequency dividers and phase shifts.")
        BluePhase = Sheet2.Range("K3")
        Exit Sub
    End If
End If

If RedFreq = 0 Then
    MsgBox ("The frequency of red cannot be set to '0'.")
    RedFreq = Sheet2.Range("E2")
    Exit Sub
End If

If GreenFreq = 0 Then
    MsgBox ("The frequency of green cannot be set to '0'.")
    GreenFreq = Sheet2.Range("H2")
    Exit Sub
End If

If BlueFreq = 0 Then
    MsgBox ("The frequency of blue cannot be set to '0'.")
    BlueFreq = Sheet2.Range("K2")
    Exit Sub
End If


RotFreq = RedFreq.Value
GrunFreq = GreenFreq.Value
BlauFreq = BlueFreq.Value

RotPhase = RedPhase.Value
GrunPhase = GreenPhase.Value
BlauPhase = BluePhase.Value

RotSat = RedSat.Value
GrunSat = GreenSat.Value
BlauSat = BlueSat.Value

RotOff = RedOff.Value
GrunOff = GreenOff.Value
BlauOff = BlueOff.Value


If RotSat And RotOff Then
    MsgBox ("Red cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If GrunSat And GrunOff Then
    MsgBox ("Green cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If

If BlauSat And BlauOff Then
    MsgBox ("Blue cannot be both always on and always off at the same time.  Please set one to 'No'.")
    Exit Sub
End If


Sheet2.Unprotect

Sheet2.Range("E2").Value = RotFreq
Sheet2.Range("H2").Value = GrunFreq
Sheet2.Range("K2").Value = BlauFreq

Sheet2.Range("E3").Value = RotPhase
Sheet2.Range("H3").Value = GrunPhase
Sheet2.Range("K3").Value = BlauPhase

Sheet2.Range("E4").Value = RotSat
Sheet2.Range("H4").Value = GrunSat
Sheet2.Range("K4").Value = BlauSat

Sheet2.Range("E5").Value = RotOff
Sheet2.Range("H5").Value = GrunOff
Sheet2.Range("K5").Value = BlauOff

Set Rng = Selection
MinNum = Application.WorksheetFunction.Min(Rng)
MaxNum = Application.WorksheetFunction.Max(Rng)

For Each cell In Rng

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

Sheet2.Protect


End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

RedFreq.Value = Sheet2.Range("E2").Value
GreenFreq.Value = Sheet2.Range("H2").Value
BlueFreq.Value = Sheet2.Range("K2").Value

RedPhase.Value = Sheet2.Range("E3").Value
GreenPhase.Value = Sheet2.Range("H3").Value
BluePhase.Value = Sheet2.Range("K3").Value

RedSat.Value = Sheet2.Range("E4").Value
GreenSat.Value = Sheet2.Range("H4").Value
BlueSat.Value = Sheet2.Range("K4").Value

RedOff.Value = Sheet2.Range("E5").Value
GreenOff.Value = Sheet2.Range("H5").Value
BlueOff.Value = Sheet2.Range("K5").Value

End Sub
