VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBclass"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
    Public WithEvents CB As MSForms.CommandButton
Attribute CB.VB_VarHelpID = -1
   
    Private Sub CB_Click()
    ButtonActive = CB.name
    Dim k As Integer
    If Left(ButtonActive, 10) = "CalendFrom" Then
        CalendLeft = 345
        k = CInt(Right(ButtonActive, (Len(ButtonActive) - 10)))
        CalendTop = (k * 18) - 12 - 108
    Else
        CalendLeft = 455 - 100
        k = CInt(Right(ButtonActive, (Len(ButtonActive) - 8)))
        CalendTop = (k * 18) - 12 - 108
    End If
    frmConfSPR.FrameFiltr.Calendar1.Left = CalendLeft
    frmConfSPR.FrameFiltr.Calendar1.Top = CalendTop
    frmConfSPR.FrameFiltr.Calendar1.Visible = True
    End Sub
