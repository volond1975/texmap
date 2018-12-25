Attribute VB_Name = "nerv_DropDownList"

 '=========================================================
 ' Author: nerv            | E-mail: nerv-net@yandex.ru
 ' Last Update: 27/03/2012 | Yandex.Money: 41001156540584
 '=========================================================
 ' module options
 Option Private Module                   ' only current project

 Public DDLRange As Range                ' range of create list
 Public DDLCell As Range                 ' active cell
 Public DLLSheetSettings As Worksheet    ' sheet settings

 Private Sub DropDownListShow()
     Dim settings, i As Long
     On Error GoTo L1
     With ActiveWorkbook.Sheets("DDLSettings")
         settings = .Range(.Cells(.Rows.Count, 1).End(xlUp).Address, .Range("D3").Address).value
     End With
     For i = 1 To UBound(settings)
         If ActiveSheet.name = settings(i, 1) Then
             With ActiveSheet.Range(settings(i, 2))
                 If ActiveCell.Row >= .Row And ActiveCell.Row <= .Row + .Rows.Count - 1 Then
                     If ActiveCell.Column >= .Column And ActiveCell.Column <= .Column + .Columns.Count - 1 Then
                         Set DDLRange = ActiveWorkbook.Sheets(settings(i, 3)).Range(settings(i, 4))
                         Set DLLSheetSettings = ActiveWorkbook.Sheets("DDLSettings")
                         Set DDLCell = ActiveCell
                         nerv_FDropDownList.Show
                         Exit Sub
                     End If
                 End If
             End With
         End If
     Next
L1:
     If Err Then
         Select Case Err
             Case 9: MsgBox "Not found sheet settings or his name invalid", vbExclamation
             Case 1004: MsgBox "Incorrect set range", vbExclamation
             Case Else: MsgBox "Unknown error", vbExclamation
         End Select
     End If
 End Sub
