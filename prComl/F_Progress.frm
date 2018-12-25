VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Progress 
   OleObjectBlob   =   "F_Progress.frx":0000
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6240
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   11
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "F_Progress"
Attribute VB_Base = "0{673D07FD-C40F-4C6E-BDFE-DC3526FC92EC}{08AD043F-8ECC-4D22-AE87-47468BCA5E14}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.PrBar.Caption = ""
End Sub
