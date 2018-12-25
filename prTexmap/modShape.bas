Attribute VB_Name = "modShape"
Sub ActiveShape()
'PURPOSE: Determine the currently selected shape
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ActiveShape As Shape
Dim UserSelection As Variant

'Pull-in what is selected on screen
  Set UserSelection = ActiveWindow.Selection

'Determine if selection is a shape
  On Error GoTo NoShapeSelected
    Set ActiveShape = ActiveSheet.Shapes(UserSelection.Name)
  On Error Resume Next

'Do Something with your Shape variable
  MsgBox "You have selected the shape: " & ActiveShape.Name

Exit Sub

'Error Handler
NoShapeSelected:
  MsgBox "You do not have a shape selected!"

End Sub
Sub Shape_Clicked()
'PURPOSE: Determine what shape was clicked to initiate macro
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ActiveShape As Shape

'Get Name of Shape that initiated this macro
  ButtonName = Application.Caller

'Set variable to active shape
  Set ActiveShape = ActiveSheet.Shapes(ButtonName)

'Do Something with your Shape variable
  MsgBox "You just clicked the " & ActiveShape.Name & " shape!"

End Sub
Function fShape_Clicked()
'PURPOSE: Determine what shape was clicked to initiate macro
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim ActiveShape As Shape

'Get Name of Shape that initiated this macro
  ButtonName = Application.Caller

'Set variable to active shape
  Set ActiveShape = ActiveSheet.Shapes(ButtonName)

'Do Something with your Shape variable
txt = ActiveShape.TextFrame2.TextRange.Text
If txt Like "Кошторис" Then
txt = ActiveShape.parent.Name & "_кошторис"
Else
txt = ActiveShape.TextFrame2.TextRange.Text
 
  End If
   fShape_Clicked = "You just clicked the " & txt & " shape!"
ThisWorkbook.Worksheets(txt).Activate
End Function
Sub ClickShape()
MsgBox fShape_Clicked()
End Sub
Sub Макрос3()
'
' Макрос3 Макрос
'

 Set PicRange = ActiveCell
 Set sh = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 531, 0.75, 63, _
        17.25)
        With sh
       .Select
        Frmul = "=" & PicRange.Address
   Selection.Formula = Frmul
    .OnAction = "ClickShape"
    sh.Top = PicRange.Top: sh.Left = PicRange.Left
    End With
End Sub
