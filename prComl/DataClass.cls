VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataClass"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
    Public WithEvents TCBData As MSForms.TextBox
Attribute TCBData.VB_VarHelpID = -1

Private Sub TCBData_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
frmConfSPR.FrameFiltr.Calendar1.Visible = False
End Sub

Private Sub TCBData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim s$
s = TCBData.Text
If IsNumeric(ChrW(KeyAscii)) Then
ElseIf KeyAscii = 46 Or KeyAscii = 44 Then
 KeyAscii = 44
 If InStr(s, ChrW(KeyAscii)) > 1 Then KeyAscii = 0
 If Len(s) = 0 Then KeyAscii = 0
Else
   KeyAscii = 0
 End If
End Sub
