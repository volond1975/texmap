VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист21"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton1, 1, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 2, 1, MSForms, CommandButton"






Private Sub ControlVivoz()
1 If ActiveCell.Offset(1, 0).value <> Empty Then
2 ActiveCell.Offset(1, 0).Range("A1").Select
3 Calculate
4 Range("EA11:EV15").Copy
5 ActiveCell.Offset(0, 8).PasteSpecial xlPasteValues
6 ActiveCell.Offset(0, -8).Select
7 Calculate
'8 GoTo 2 'Control4
End If

9 If ActiveCell.Offset(1, 0).value = Empty And ActiveCell.Offset(5, 0).value <> Empty And ActiveCell.Offset(5, 0).value <> "Всього по ділянці" Then
10 ActiveCell.Offset(1, 0).Range("A5").Select
11 Calculate
12 Range("EA11:EV15").Copy
13 ActiveCell.Offset(0, 8).PasteSpecial xlPasteValues
14 ActiveCell.Offset(0, -8).Select
15 Calculate
16 GoTo 9 'Control4
End If
If ActiveCell.Offset(5, 0).value = "Всього по ділянці" Then 'And ActiveCell.Offset(3, 6).Value = "Всього по відводу" Then
Calculate
'ActiveCell.Offset(1, 0).Range("A3").Select

Range("DY1:EV5").Copy
ActiveCell.Offset(5, 6).PasteSpecial xlPasteValues
ActiveCell.Offset(0, -6).Select
ActiveCell.Offset(0, 0).value = "Всього по ділянці"
ElseIf ActiveCell.Offset(1, 0).value = Empty And ActiveCell.Offset(5, 0).value = Empty Then
MsgBox "Дані оновлені", , "Повідомлення"
End If
If ActiveCell.Offset(0, 0).value = "Всього по ділянці" Then
ActiveCell.Offset(1, 0).Range("A5").Select
If ActiveCell.Offset(0, 0).value <> Empty And ActiveCell.Offset(0, 0).value <> "Всього по ділянці" Then
Calculate
Range("EA11:EV15").Copy
ActiveCell.Offset(0, 8).PasteSpecial xlPasteValues
ActiveCell.Offset(0, -8).Select
Calculate
GoTo 9 'Control4
End If
End If

Range("A11").Select
End Sub

Private Sub CommandButton1_Click()
    Application.ScreenUpdating = False
Range("A1").Select
Skrivaet

End Sub

Private Sub OnovDaniVivozka_Click()
Application.ScreenUpdating = False
Range("A1").Select
ActiveSheet.Unprotect "sirius"
'    Range("A6:A30006").Select
 '   Selection.EntireRow.Hidden = False
    With Application
        .Calculation = xlManual
        .MaxChange = 0.001
    End With
 
    Range("A10").Select
ControlVivoz
    With Application
        .Calculation = xlAutomatic
        .MaxChange = 0.001
    End With
ActiveSheet.Protect "sirius"
End Sub
