VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   OleObjectBlob   =   "frmSettings.frx":0000
   Caption         =   "Настройки Калькулятора"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   15
End
Attribute VB_Name = "frmSettings"
Attribute VB_Base = "0{125C636E-89FB-4C23-807D-87E19301C108}{6337445D-E3B7-42EF-9B8E-F79AA868A45C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    frmCalc.Show vbModeless
End Sub

Private Sub cmdSave_Click()
    Call SaveSettingsInRegestry("CloseByInsertKey", Me.ChkBoxCloseProg.value)
    Unload Me
    frmCalc.Show vbModeless
End Sub

Private Sub SaveSettingsInRegestry(iParam As String, iValue As String)
    Call SaveSetting(PROGID, "Settings", iParam, iValue)
End Sub

Private Function GetRegValue(iParam As String)
    GetRegValue = GetSetting(PROGID, "Settings", iParam, False)
End Function

Private Sub lblLink3_Click()
    ForwardLink lblLink3.Caption
End Sub

Private Sub ForwardLink(sUrl As String)
Dim oShell As Object
    Set oShell = CreateObject("Wscript.Shell")
    oShell.Run (sUrl)
    Set oShell = Nothing
End Sub

Private Sub UserForm_Initialize()
    Me.ChkBoxCloseProg.value = GetRegValue("CloseByInsertKey")
End Sub
