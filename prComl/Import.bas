Attribute VB_Name = "Import"
Sub ��������������(���_�����)
Dim EOBookName As String
Dim EOBookSheetName As String
Dim EOBookSheetNewName As String
Dim twb As Workbook
Dim r As Range
Set twb = ThisWorkbook

EOBookName = ��������������������("������", "������", "����� ��� �����", "���� ��� �������", ���_�����)
EOBookSheetName = ��������������������("������", "������", "����� ��� �����", "����", ���_�����)
EOBookSheetNewName = ���_�����
Call ������(EOBookName, EOBookSheetName, EOBookSheetNewName)


End Sub

Sub ���������������()
Dim EOBookName As String
Dim EOBookSheetName As String
Dim EOBookSheetNewName As String
Dim twb As Workbook
Set twb = ThisWorkbook
EOBookName = "����� ��������� ������.xls"
EOBookSheetName = "TDSheet"
EOBookSheetNewName = "������_��"
Call ������(EOBookName, EOBookSheetName, EOBookSheetNewName)


End Sub
Sub ���������������(EOBookName As String)

Dim EOBookSheetName As String
Dim EOBookSheetNewName As String
Dim twb As Workbook
Call Intro
Set twb = ThisWorkbook

EOBookSheetName = "TDSheet"
EOBookSheetNewName = "�������� ��������"
����������������������� = 100
������������� = 50
Call ������(EOBookName, EOBookSheetName, EOBookSheetNewName)
Set ����_��������� = Cells.Find(What:="���� ���������", LookAt:=xlWhole)
Set ������� = Cells.Find(What:="�������", LookAt:=xlWhole)
Set ��_� = Cells.Find(What:="�� �", LookAt:=xlPart)
Set ������ = Cells.Find(What:="������", LookAt:=xlPart)
Set ��� = Range(����_���������, Cells(������.Offset(rowoffset:=-1).Row, �������.Column))
ActiveSheet.ListObjects.Add(xlSrcRange, Range(���.address), , xlYes).name = _
        "��������"
    
    ActiveSheet.ListObjects("��������").TableStyle = "TableStyleMedium2"
   Call ���������������
 EOBookSheetNewName = "������� ��������"
  Call ��������_����("���������������", EOBookSheetNewName, 3)


Set sh = twb.Worksheets(EOBookSheetNewName)
'�������� ������������� ���������
sh.Cells(1, 8).value = "�������� ����, ���./�3"
sh.Cells(1, 9).value = "����, ���."
Call ����������������(Range(sh.Cells(1, 1), sh.Cells(1, 9)))
With Range(sh.Cells(1, 1), sh.Cells(1, 9))
.HorizontalAlignment = xlCenter
.WrapText = True
With .Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With .Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With



End With





Call ���������������������������(�����������������������, �������������)
 Call Outro
End Sub

Sub ������������ADO()
    Dim ADO As New ADO
Dim oCN   As Object    'Connection
Dim oRS   As Object
Dim fl As Object
Dim wb As Workbook   'Recordset
Dim lo As ListObject
Dim sQuery As String  ' ������� ��������� ������
ListObjectName = "���������"
Set wb = Workbooks("��������.xlsm")
 ShetNameActivate = ActiveSheet.name

Set sha = wb.Worksheets(ShetNameActivate)
Set sh = SheetExistBookCreate(wb, "����3", True)

   ADO.DataSource = ThisWorkbook.Path & "\" & "�������� 2016.xls"
   sQuery = "SELECT * FROM [" & ListObjectName & "$]"
ADO.Query (sQuery)
wb.Activate
sh.Activate

 'Set fl = ADO.Fields
'Set oRS = ADO.Recordset

' lo.QueryTable.Recordset = oRS
''     Set .Recordset = oRS
'    sh.QueryTable.Refresh

  sh.Range("A1").CopyFromRecordset ADO.Recordset


'  Set lo = sh.ListObjects.Add
'sh.Name = "������"
'    ADO.Query ("SELECT F2 FROM [����1$];")
'    Range("F1").CopyFromRecordset ADO.Recordset
'
'    ' ��������� ����������, ����� �� ������ : )
'    ADO.Disconnect
'
'    ADO.Query ("SELECT F1 FROM [����1$] UNION SELECT F2 FROM [����1$];")
'    Range("G1").CopyFromRecordset ADO.Recordset
        
    ' ��� ������������� ��������� ����������
    ' � ������������ ������� Recordset � Connection
End Sub





























Sub ���������������������������(�����������������������, �������������)
Dim twb As Workbook
Dim EOBookSheetNewName As String
Dim ��������� As Range
Dim v As Variant
Dim sh As Worksheet
Set twb = ThisWorkbook
EOBookSheetNewName = "������� ��������"
Set sh = twb.Worksheets(EOBookSheetNewName)

����������� = LastRow(EOBookSheetNewName)
For I = 2 To �����������
 If I > ����������� Then Exit For
 
 If sh.Cells(I, 1) = "����� ���������� (���)" And sh.Cells(I, 3) <> "" Then
 fAkt.ListBox_��������.Clear
 Set ��������� = sh.Rows(I)
 For j = 1 To 3
 �������� = LastRow(EOBookSheetNewName) + 1
 sh.Cells(��������, 1).value = ���������.Cells(1).value
 sh.Cells(��������, 2).value = ���������.Cells(2).value
 sh.Cells(��������, 10).value = ���������.Cells(1).value & " " & ���������.Cells(2).value & " " & j & "�"
 sh.Cells(��������, 4).value = j
 sh.Cells(��������, 8).value = fAkt.TextBox_�����������������������.value
 fAkt.ListBox_��������.AddItem sh.Cells(��������, 10).value
 Call ���������������(sh.Cells(��������, 9))
 Next j
 ���������.Delete
 I = I - 1
 ����������� = ����������� - 1
 Else
 
 If sh.Cells(I, 1) = "�����" Then
 If I <> 1 Then sh.Cells(I, 8).value = fAkt.TextBox_�������������.value
 
 Else
  If I <> 1 Then sh.Cells(I, 8).value = fAkt.TextBox_�����������������������.value
 
 
 End If
 Call ���������������(sh.Cells(I, 9))
 End If

Next I
'������� ������ ��������
For I = 1 To 2
�������� = LastRow(EOBookSheetNewName) + 1
'������������,'������,�������,������ ��������,����(������),���������,�����,����
v = Array("������� �������", "", "", "", I, "", "", fAkt.TextBox_�����������������������.value, "", "������� ������� " & I & " ���")
�������� = LastRow(EOBookSheetNewName) + 1
Call ����������������������(sh, ��������, v)
fAkt.ListBox_��������.AddItem sh.Cells(��������, 10).value
Next I
'������� ����� �� ������
�������� = LastRow(EOBookSheetNewName) + 1
'������������
sh.Cells(��������, 1).value = "������ ������"
Call ����������������������������(sh.Cells(��������, 7), sh.Cells(2, 7))
Call ����������������������������(sh.Cells(��������, 9), sh.Cells(2, 9))
Call ����������������(Range(sh.Cells(��������, 1), sh.Cells(��������, 9)))
'�������� ������ ������ �������
Set ��������� = sh.Rows(��������)
'������� ������ ����������
�������� = LastRow(EOBookSheetNewName) + 1
'������������
sh.Cells(��������, 1).value = "���� ���������"
sh.Cells(��������, 10).value = "���� ���������"
fAkt.ListBox_��������.AddItem sh.Cells(��������, 10).value
'������� �����
�������� = LastRow(EOBookSheetNewName) + 1
'������ ��� ���
sh.Cells(��������, 1).value = "������ ��� ���"
sh.Cells(��������, 1).Font.Bold = True
sh.Cells(��������, 7).Formula = "=" & ���������.Cells(7).address & "+" & sh.Cells(�������� - 1, 7).address
sh.Cells(��������, 7).Font.Bold = True
sh.Cells(��������, 9).Formula = "=" & ���������.Cells(9).address & "+" & sh.Cells(�������� - 1, 9).address
sh.Cells(��������, 9).Font.Bold = True
'���
�������� = LastRow(EOBookSheetNewName) + 1
sh.Cells(��������, 1).value = "���"
sh.Cells(��������, 1).Font.Bold = True
sh.Cells(��������, 8).value = 0.2
sh.Cells(��������, 8).Font.Bold = True
sh.Cells(��������, 8).NumberFormat = "0.00%"

sh.Cells(��������, 9).Formula = "=" & sh.Cells(��������, 8).address & "*" & sh.Cells(�������� - 1, 9).address
sh.Cells(��������, 9).Font.Bold = True
'������ � ���
�������� = LastRow(EOBookSheetNewName) + 1
sh.Cells(��������, 1).value = "������ � ���"
sh.Cells(��������, 1).Font.Bold = True
sh.Cells(��������, 7).Formula = "=" & sh.Cells(�������� - 2, 7).address & "+" & sh.Cells(�������� - 1, 7).address
sh.Cells(��������, 9).Formula = "=" & sh.Cells(�������� - 2, 9).address & "+" & sh.Cells(�������� - 1, 9).address

Call ����������������(Range(sh.Cells(��������, 1), sh.Cells(��������, 9)))
sh.UsedRange.Select
Call ������������

�������� = LastRow(EOBookSheetNewName)
Range(sh.Cells(2, 1), sh.Cells(��������, 5)).Select
Call ������������

Range(sh.Cells(�������� - 2, 1), sh.Cells(��������, 5)).Select
Selection.Font.Size = 12
Call ������������
Range(sh.Cells(�������� - 2, 6), sh.Cells(��������, 9)).Select
Selection.Font.Size = 12
Call ������������

Range(sh.Cells(�������� - 1, 1), sh.Cells(��������, 5)).Select
Selection.Font.Size = 12
Call ������������
Range(sh.Cells(�������� - 1, 6), sh.Cells(��������, 9)).Select
Selection.Font.Size = 12
Call ������������


Range(sh.Cells(��������, 1), sh.Cells(��������, 5)).Select
Selection.Font.Size = 12
Call ������������
Range(sh.Cells(��������, 6), sh.Cells(��������, 9)).Select
Selection.Font.Size = 12
Call ������������

Range(sh.Cells(���������.Row, 1), sh.Cells(���������.Row, 5)).Select
Selection.Font.Size = 12
Call ������������
Range(sh.Cells(���������.Row, 6), sh.Cells(���������.Row, 9)).Select
Selection.Font.Size = 12
Call ������������

End Sub
Sub ����������������������(sh As Worksheet, ��������, v As Variant)
'������������,'������,�������,������ ��������,����(������),���������,�����


'������������
sh.Cells(��������, 1).value = v(0)
'������
sh.Cells(��������, 2).value = v(1)
'�������
sh.Cells(��������, 3).value = v(2)
'������ ��������
sh.Cells(��������, 4).value = v(3)
'����(������)
sh.Cells(��������, 5).value = v(4)
'���������
sh.Cells(��������, 6).value = v(5)
'�����
sh.Cells(��������, 7).value = v(6)
'��������
sh.Cells(��������, 8).value = v(7)
'�����
sh.Cells(��������, 9).value = v(8)
'��������
sh.Cells(��������, 10).value = v(9)

End Sub



Sub ������(EOBookName As String, EOBookSheetName As String, EOBookSheetNewName As String)

Dim twb As Workbook
Set twb = ThisWorkbook

If SheetExistBook(twb, EOBookSheetNewName) Then

Set ws = twb.Worksheets(EOBookSheetNewName)
ws.Cells.Clear
Else
Set ws = twb.Worksheets.Add
ws.name = EOBookSheetNewName
End If

Set b = mywbBook(EOBookName, twb.Path)
If b Is Nothing Then MsgBox ("���� " & twb.Path & "\" & EOBookName)
Workbooks(EOBookName).Worksheets(EOBookSheetName).Cells.Copy
twb.Activate
ws.Activate
Cells.Activate

ActiveSheet.Paste
'ActiveSheet.Name = EOBookSheetNewName
Workbooks(EOBookName).Close
'ThisWorkbook.Worksheets("�����").Activate


End Sub

Sub �����������������()
Dim twb As Workbook
Dim shPriymannya As Worksheet
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("���������")
nrow = 9
erow = LastRow("���������")
For I = nrow To erow
Next I
End Sub

Sub ����������������()
Dim twb As Workbook
Dim shPriymannya As Worksheet
Dim shPrivyaska As Worksheet
Dim v As Range
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("��������")
shPriymannya.Activate
Set ar = ActiveCell
q = ar.Row
Set shPrivyaska = twb.Worksheets("��������")
Worksheets("�����").Activate
nrow = 2
erow = LastRow("��������")
For I = nrow To erow
If shPrivyaska.Cells(I, 3).value <> "" Then

Set r = FindAll(shPriymannya.Rows(1), shPrivyaska.Cells(I, 1).value)
Set v = Range(shPrivyaska.Cells(I, 3))
If r.value = "����" Then
z = Split(shPriymannya.Cells(q, r.Column).value, "_")
If UBound(z) > 0 Then
v.value = z(0)
v.Offset(rowoffset:=1).value = z(1)
Else
v.value = shPriymannya.Cells(q, r.Column).value
v.Offset(rowoffset:=1).value = "__"
End If
Else
v.value = shPriymannya.Cells(q, r.Column).value
End If
'shPriymannya.Activate
End If
Next I
End Sub
