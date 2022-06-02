VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Conditional Formatting Choices"
   ClientHeight    =   5346
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   5772
   OleObjectBlob   =   "BespokeCondForm_MultiChoice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()


If OptionButton1 Then
    Call BubblegumUnicorn
    
ElseIf OptionButton2 Then
    Call RainbowBrite
    
ElseIf OptionButton3 Then
    Call RainbowMedium
    
ElseIf OptionButton4 Then
    Call EyeOfSauron
    
ElseIf OptionButton5 Then
    Call CoralReef
    
ElseIf OptionButton6 Then
    Call RainbowSuperBrite
    
ElseIf OptionButton7 Then
    Call JungleSunrise
    
Else
    MsgBox ("Please choose one of the options.")

End If


End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub CommandButton3_Click()

Call ClearFormat

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub OptionButton4_Click()

End Sub

Private Sub OptionButton5_Click()

End Sub

Private Sub OptionButton6_Click()

End Sub
