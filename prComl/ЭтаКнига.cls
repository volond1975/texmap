VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1

 Private Sub app_SheetBeforeDoubleClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
     On Error Resume Next: Err.Clear
     If Not (sh.Parent Is ThisWorkbook) Then Exit Sub

'     If Not SettingsBoolean("CheckBox_DisableDoubleClickOnHeader") Then
         If Target.Column = HEADER_Column And Target.Cells.Count = 1 Then
             If Len(Target) > 0 Then
                 Cancel = True
                 If Target <> Trim(Target) Then Target = Trim(Target)
                 With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
                     .SetText "{" & Trim(Target) & "}"
                     .PutInClipboard
                 End With
                 If Err = 0 Then Shell "Cmd.exe /c echo " & Chr(7), vbHide
             End If
         End If
'     End If

     If Target.Cells.Count = 1 Then
         If HasLinkToObject(Target.Cells(1)) Then
             Cancel = True
             CtrlShiftT
         End If
     End If
End Sub

Private Sub Workbook_Activate()
Call CalculationM
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Call DeleteMenu
Call CalculationA
End Sub

Private Sub Workbook_Open()
Call CreateMenu
ActiveWorkbook.RefreshAll
End Sub
