Attribute VB_Name = "������"

Sub ����������������()
'
' ���������������� ������
' ������ ������� 25.06.2012 (������������)
'

'
With Worksheets("������")
.Range("F8:G22").ClearContents
.Range("I8:P22").ClearContents
.Range("S8:U22").ClearContents
End With

'Call CalculationA
'Call CalculationM
End Sub
Sub �������������������()
'
' ������������������� ������
' ������ ������� 25.06.2012 (������������)
'

'

    Range("G29:G40").Select
    Selection.copy
    Sheets("���").Select
    Range("F40:F50").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
  Sheets("���").Range("gNel").value = Sheets("������").Range("G42")
        
End Sub
