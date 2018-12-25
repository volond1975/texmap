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
Attribute VB_Base = "0{339667A4-23EA-4CDE-86F8-7849433170F8}{EC9B5AF1-9D01-4280-A80E-B87D366DC129}"
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
