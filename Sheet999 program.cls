VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub HTMLImage1_Click()

End Sub

Private Sub HTMLImage2_Click()

End Sub

Private Sub HTMLOption1_Click()

End Sub

Private Sub HTMLOption5_Click()

End Sub

Private Sub HTMLText1_Click()

End Sub

Private Sub HTMLText2_Click()

End Sub

Private Sub HTMLText3_Click()

End Sub

Private Sub HTMLText4_Click()

End Sub

Private Sub HTMLText5_Click()

End Sub

Private Sub HTMLText6_Click()

End Sub

Private Sub HTMLText6_GotFocus()

End Sub

       SelectedTool = "Multimeter"
       MsgBox "Tool selected: " & SelectedTool
   End Sub

   

   End Sub
       If HTMLText3.text = "" Then
           MsgBox "Please enter voltage value."
       Else
           Voltage = Val(HTMLText3.text)
           If Voltage > 240 Then
               MsgBox "Overvoltage detected. Check circuit."
           End If
       End If
   
      HTMLText6.BackColor = RGB(255, 255, 200)
       MsgBox "Enter AI configuration parameters here."
   End Sub

