VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEDERPUorINN 
   OleObjectBlob   =   "frmEDERPUorINN.frx":0000
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5430
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   20
End
Attribute VB_Name = "frmEDERPUorINN"
Attribute VB_Base = "0{168D9F52-FC8A-4F58-9AA4-B73CF0B9FBBF}{F9699688-3270-4592-ADD0-36F0F815A5C0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub ComboBox1_Change()
edrpu
End Sub

Private Sub ComboBox10_Change()
edrpu
End Sub

Private Sub ComboBox2_Change()
edrpu
End Sub

Private Sub ComboBox3_Change()
edrpu
End Sub

Private Sub ComboBox4_Change()
edrpu
End Sub

Private Sub ComboBox5_Change()
edrpu
End Sub

Private Sub ComboBox6_Change()
edrpu
End Sub

Private Sub ComboBox7_Change()
edrpu
End Sub

Private Sub ComboBox8_Change()
edrpu
End Sub

Private Sub ComboBox9_Change()
edrpu
End Sub



Private Sub CommandButton1_Click()
ActiveCell.value = Me.Frame2.Caption
Me.Hide
End Sub

Private Sub OptionButton2_Change()
If OptionButton2.value Then
Me.Caption = "ÅÄÐÏÓ"
Me.ComboBox9.Visible = False
Me.ComboBox10.Visible = False
Me.Frame2.Width = 204
Me.Width = 220
Else
Me.ComboBox9.Visible = True
Me.ComboBox10.Visible = True
Me.Frame2.Width = 265
Me.Width = 280
Me.Caption = "²ÍÍ"
End If

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
For I = 1 To VBA.Len(ActiveCell.value)
Me.Controls("ComboBox" & I).Clear
For j = 0 To 9
Me.Controls("ComboBox" & I).AddItem j
Next j

Next I




If VBA.Len(ActiveCell.value) <> 10 And VBA.Len(ActiveCell.value) <> 8 Then

Exit Sub
Else
If VBA.Len(ActiveCell.value) = 10 Then Me.OptionButton1.value = True Else Me.OptionButton2.value = True
For I = 1 To VBA.Len(ActiveCell.value)
Me.Controls("ComboBox" & I).Clear
For j = 0 To 9
Me.Controls("ComboBox" & I).AddItem j
Next j
Me.Controls("ComboBox" & I).value = Mid(ActiveCell.value, I, 1)
Next I
End If

Frame2.Caption = ActiveCell.value
End Sub
Function edrpu()
For I = 1 To VBA.Len(ActiveCell.value)
z = z & Me.Controls("ComboBox" & I).value
Next I
Frame2.Caption = z
End Function
