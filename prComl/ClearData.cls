VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClearData"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public WithEvents Cleare As MSForms.CommandButton
Attribute Cleare.VB_VarHelpID = -1
   
    Private Sub Cleare_Click()
    ClearCB = Cleare.name
    Dim k As Integer
        k = CInt(Right(ClearCB, (Len(ClearCB) - 13)))
        frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & k).value = ""
        frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & k).value = ""
        frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & k).value = ""
   End Sub
