VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fData 
   OleObjectBlob   =   "fData.frx":0000
   Caption         =   "Акт від"
   ClientHeight    =   1080
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4320
   TypeInfoVer     =   11
End
Attribute VB_Name = "fData"
Attribute VB_Base = "0{15813BBB-1E54-4573-B3C5-768DC1B9E1B7}{F332F6FF-A3B8-4EF6-8067-9ED4B28B9EA5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub ComboBox_Mouth_Change()
If Val(Me.ComboBox_Day) > ДнейВМесяце(Val(ComboBox_Mouth), Val(Me.ComboBox_Year)) Then Me.ComboBox_Day = Val(Me.ComboBox_Day - 1)
End Sub

Private Sub ComboBox_Year_Change()
If Val(Me.ComboBox_Day) > ДнейВМесяце(Val(ComboBox_Mouth), Val(Me.ComboBox_Year)) Then Me.ComboBox_Day = Val(Me.ComboBox_Day - 1)
End Sub

Private Sub CommandButton_Ok_Click()
fAkt.TextBox_ДатаЗакДП.value = Format(VBA.DateSerial(Val(Me.ComboBox_Year), Val(Me.ComboBox_Mouth), Val(Me.ComboBox_Day)), "Short Date")
fAkt.Label_Log.Caption = "Номер и дата акта сформирована !  Нажмить кнопку Ок "
Me.Hide
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
For I = 1 To 31
Me.ComboBox_Day.AddItem I
Next I
m = VBA.Month(Now())
For I = 1 To 12
Me.ComboBox_Mouth.AddItem I
Next I
If m = 1 Then
Me.ComboBox_Mouth.value = 12
k = 12
Else
Me.ComboBox_Mouth.value = m - 1
k = m - 1
End If
D = Year(Now())
For I = D - 4 To D + 4
Me.ComboBox_Year.AddItem I
Next I
If m = 1 Then
Me.ComboBox_Year.value = D - 1
z = D - 1
Else
Me.ComboBox_Year.value = D
z = D
End If
dn = ДнейВМесяце(k, z)
Me.ComboBox_Day.value = dn
End Sub
