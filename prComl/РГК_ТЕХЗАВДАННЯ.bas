Attribute VB_Name = "���_�����������"
Sub �����������������()
Attribute �����������������.VB_Description = "������ ������� 27.10.2012 (volond)"
Attribute �����������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������5 ������
' ������ ������� 27.10.2012 (volond)
Dim r1 As Range
Dim r2 As Range
Set r1 = Workbooks("�����������.xls").Worksheets(1).Range("B20:C34")
Set r2 = Worksheets("������").Range("F8")
Call ���������(r1, r2)


Set r1 = Workbooks("�����������.xls").Worksheets(1).Range("D20:F34")
Set r2 = Worksheets("������").Range("I8")
Call ���������(r1, r2)
    
Set r1 = Workbooks("�����������.xls").Worksheets(1).Range("G20:H34")
Set r2 = Worksheets("������").Range("L8")
Call ���������(r1, r2)

    Set r1 = Workbooks("�����������.xls").Worksheets(1).Range("I20:J34")
Set r2 = Worksheets("������").Range("O8")
Call ���������(r1, r2)
'
End Sub
Sub ���������(r1 As Range, r2 As Range)
Attribute ���������.VB_Description = "������ ������� 27.10.2012 (volond)"
Attribute ���������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������6 ������
' ������ ������� 27.10.2012 (volond)
'

'
    Windows("�����������.xls").Activate
    r1.Copy
'    Selection.Copy
    Windows("�������� �������������� �������� 1.1.xls").Activate
    r2.Select
    Windows("�����������.xls").Activate
    r1.Copy
    Windows("�������� �������������� �������� 1.1.xls").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub �������������()
'
' ������5 ������
' ������ ������� 27.10.2012 (volond)
Dim r1 As Range
Dim r2 As Range

'("�������� " & Me.TextBox_���_�������� & "." & Me.ComboBox_����������_��������, ThisWorkbook.Path & "\")

Set r1 = Workbooks("�������� " & Me.TextBox_���_�������� & "." & Me.ComboBox_����������_��������).Worksheets("������").Range("F8:G22")
Set r2 = Worksheets("���� ������� ���").Range("B3")
Call ���������(r1, r2)


Set r1 = Workbooks("�����������.xls").Worksheets(1).Range("D20:F34")
Set r2 = Worksheets("������").Range("I8")
Call ���������(r1, r2)
    
Set r1 = Workbooks("�����������.xls").Worksheets(1).Range("G20:H34")
Set r2 = Worksheets("������").Range("L8")
Call ���������(r1, r2)

    Set r1 = Workbooks("�����������.xls").Worksheets(1).Range("I20:J34")
Set r2 = Worksheets("������").Range("O8")
Call ���������(r1, r2)
'
End Sub
Sub �����������������(r1 As Range, r2 As Range)
'
' ������6 ������
' ������ ������� 27.10.2012 (volond)
'

'
    Windows("�������� " & Me.TextBox_���_�������� & "." & Me.ComboBox_����������_��������).Activate
    r1.Copy
'    Selection.Copy
    Windows("�������� �������������� �������� 1.1.xls").Activate
    r2.Select
    Windows("�����������.xls").Activate
    r1.Copy
    Windows("�������� �������������� �������� 1.1.xls").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub ������13()
'
' ������13 ������
'

'Workbooks( _
'        "�������� " & fAkt.TextBox_���_�������� & "." & fAkt.ComboBox_����������_��������).Activate
'
'    Workbooks( _
'        "�������� " & fAkt.TextBox_���_�������� & "." & fAkt.ComboBox_����������_��������).Sheets("������").Copy Before:=Workbooks( _
'        "��������� ���� 2 (���������������).xlsm").Sheets(1)
    Sheets("������").Select
    Range("F8:G22").Select
    Selection.Copy
    Sheets("���� ������� ���").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("������").Select
    Range("I8:K22").Select
    Selection.Copy
    Sheets("���� ������� ���").Select
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("������").Select
    Range("L8:M22").Select
    Selection.Copy
    Sheets("���� ������� ���").Select
    Range("G3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("������").Select
    Range("O8:P22").Select
    Selection.Copy
    Sheets("���� ������� ���").Select
    Range("I3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("������").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
End Sub
