VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fAkt 
   OleObjectBlob   =   "fAkt.frx":0000
   Caption         =   "UserForm2"
   ClientHeight    =   10170
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9510
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   414
End
Attribute VB_Name = "fAkt"
Attribute VB_Base = "0{493AE1D8-4A19-4ADC-A666-F40273384684}{22E85829-7D04-4644-9CEF-7C8FB0F91DB6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub cbKorectCall3_Click()
'[f�������_������_���] = [f�������_����_�����]
'TextBox_����� = [fm����2] + [fm����3]
'TextBox_����� = [fm����5] + [fm����6]
'TextBox_����� = [fm����8] + [fm����9]
End Sub

Private Sub CheckBox_������_Click()

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox_�_���_Click()
Set r = FindAll(Worksheets("������� ��������").Cells, "���")
If Me.CheckBox_�_��� Then
r.Offset(columnoffset:=7).value = 0.2
Else
r.Offset(columnoffset:=7).value = 0

End If
Call ����������������

End Sub

Private Sub ComboBox_Full_Dilyanka_Change()
Dim FullDil As String
FullDil = ComboBox_Full_Dilyanka.Text
v = ������������(FullDil)
Me.ComboBox_������� = v(0)
Me.ComboBox_���� = v(1)
Me.ComboBox_ĳ����� = v(2)
Me.TextBox_C���_˳���.Text = Range("fSokrLis").value
Dim twb As Workbook
Dim shMastera As Worksheet
Set twb = ThisWorkbook
Set shMastera = twb.Worksheets("�������")
'Dim ��������������������(��������, ����������, ����������������, ������������������, ��������������)

Me.TextBox_�����.value = ""
Me.TextBox_�����.value = ""
Me.TextBox_�����.value = ""


End Sub

Private Sub ComboBox_Lisnuctvo_Change()

Dim twb As Workbook
Dim shPriymannya As Worksheet
Dim col As Collection
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("���������")
nrow = 9
erow = LastRow("���������")

Me.ComboBox_Full_Dilyanka.Clear
Set dil = shPriymannya.Cells.Find("ĳ�����")
Set lis = shPriymannya.Cells.Find("˳�������")
Set dp = shPriymannya.Cells.Find("������, ����� ��������")
Me.ComboBox_Full_Dilyanka.Clear
Me.ComboBox_���.Clear
For I = nrow To erow
If shPriymannya.Cells(I, lis.Column).value = Me.ComboBox_Lisnuctvo.value Then
Me.ComboBox_Full_Dilyanka.AddItem shPriymannya.Cells(I, dil.Column).value
z = Split(shPriymannya.Cells(I, dp.Column).value, ",")
Me.ComboBox_���.AddItem z(1)
End If
Next I
Range("f˳�������").value = Me.ComboBox_Lisnuctvo.value
Me.TextBox_C���_˳���.Text = Range("fSokrLis").value

Me.ListBox_�������.RowSource = "=�������[" & Me.TextBox_C���_˳���.value & "]"
Me.ComboBox_��������.RowSource = "=�������[" & Me.TextBox_C���_˳���.value & "]"

'D = FiltUT(ThisWorkbook, "������_��", "������_��", vf, vc)
'D = FiltUT(ThisWorkbook, "��������", "��������", vf, vc)

vf = Array("2")


vc = Array(fAkt.ComboBox_Lisnuctvo)
D = FiltUT(ThisWorkbook, "��������", "��������", vf, vc)

vf = Array("1")
vc = Array(fAkt.ComboBox_Lisnuctvo)
'D = FiltUT(ThisWorkbook, "������_��", "������_��", vf, vc)

End Sub

Private Sub ComboBox_���_Change()
Me.ComboBox_����.value = fAkt.TextBox_C���_˳��� & "\" & fAkt.ComboBox_Full_Dilyanka & "\" & VBA.Month(Me.TextBox_�������)

End Sub

Private Sub ComboBox_����_Change()

End Sub

Private Sub ComboBox_³��������_Change()

End Sub

Private Sub ComboBox_�������_Change()
'vf = Array("2", "8")
'If fAkt.ComboBox_ĳ����� = 0 Then
'z = fAkt.ComboBox_����
'Else
'z = fAkt.ComboBox_���� & "_" & fAkt.ComboBox_ĳ�����
'End If
'
'vc = Array(fAkt.ComboBox_Lisnuctvo, fAkt.ComboBox_�������)
'D = FiltUT(ThisWorkbook, "��������", "��������", vf, vc)

'vf = Array("1", "4")
'vc = Array(fAkt.ComboBox_Lisnuctvo, fAkt.ComboBox_Full_Dilyanka)
'D = FiltUT(ThisWorkbook, "������_��", "������_��", vf, vc)
End Sub

Private Sub ComboBox_��������_Change()
Worksheets("�����").Range("f������_����") = Me.ComboBox_��������
Call CalculationA
Call CalculationM
End Sub

Private Sub ComboBox1_���_Change()

End Sub

Private Sub CommandButton_�������_Click()
 Call fDataShow
End Sub

Private Sub CommandButton_������_Click()
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("���������")
Set shChoden = twb.Worksheets("�������� ��������")
nrow = 9
erow = LastRow("���������")
shPriymannya.Activate
Set dil = shPriymannya.Cells.Find("ĳ�����")
Set lis = shPriymannya.Cells.Find("˳�������")
Set mast = shPriymannya.Cells.Find("�������")
Set mast = shPriymannya.Cells.Find("�������", mast)
Set lk = shPriymannya.Cells.Find("˳�������� ������")
Set dp = shPriymannya.Cells.Find("������, ����� ��������")
Set vik = shPriymannya.Cells.Find("˳��������������.������������")
'Me.ComboBox_Full_Dilyanka.Clear
For I = nrow To erow
If shPriymannya.Cells(I, lis.Column).value = Me.ComboBox_Lisnuctvo.value Then
 If shPriymannya.Cells(I, dil.Column).value = Me.ComboBox_Full_Dilyanka.value Then
   If shPriymannya.Cells(I, mast.Column).value <> "" Then
master = shPriymannya.Cells(I, mast.Column).value
Me.ListBox_�������.value = master
End If
z = Split(shPriymannya.Cells(I, lk.Column).value, " ")
Me.ComboBox_���.value = z(2)
Me.TextBox_������.Text = z(4)
'Range("f�������").Value = z(4)
z = Split(shPriymannya.Cells(I, dp.Column).value, ",")
v = Split(Trim(z(0)), " ")
Me.ComboBox_���.value = z(1)
Me.TextBox_������.Text = v(3)

Me.ComboBox_����������.value = shPriymannya.Cells(I, vik.Column).value
End If
End If
Next I

Me.Label_Log.Caption = "�������� ������� �� ������ ! ����� ���� ������� ������ ���� "

End Sub

Private Sub CommandButton1_Click()
If Me.CheckBox_Import_All Then
Else
Call ��������������(Me.ComboBox_Report)
End If
Me.Label_Log.Caption = "��� ������ �����������"
End Sub

Private Sub CommandButton11_Click()
 Set ThisWorkbook.app = Application
End Sub

Private Sub CommandButton12_Click()
Call ����������������
End Sub

Private Sub CommandButton13_Click()
Dim shp As Worksheet
Dim lo As ListObject
Dim q As Range
Set wb = ThisWorkbook
Set sh = wb.Worksheets("�����")
���������� = sh.Range("fDot").value
��������������� = ���������� & ".dot"

��������������������������� = ".doc"
���������� = "�����"
��������������� = "���������"
������������������ = "���������"
������������������ = "��������"



Set ��������������������� = FindAll(sh.Rows(1), ���������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)

  
    ���������� = NewFolderName & Application.PathSeparator

 'sh.Activate
'  v = Split(Range("f����").Value, "\")
'����� = sh.Range("f������ĳ�").Value
'����� = Replace(�����, ".", "")
'            �������� = ����� & "-" & sh.Range("fDot").Value & "-" & sh.Range("f���").Value ' Trim$(MonthTK(Range("cMonth")) & "_" & ����TK(Range("cLicVa")) & "_" & Range("�N"))
'            FileNameDoc = ���������� & �������� & ���������������������������
            
        �������� = NewFaleName
'            FileNameDoc = ���������� & �������� & ���������������������������
FileNameDoc = NewFileFullName(����������, ��������, ���������������������������)
            
            
If Me.ListBox_������.ListIndex = -1 Then
MsgBox "�� ������ ��� ��������� ��� ������"
Exit Sub
End If
'If Me.ListBox_������ Like "�������*" Then doc = Me.ListBox_������.ListIndex + 2 Else doc = Me.ListBox_������.ListIndex + 1
doc = Me.ListBox_������.ListIndex + 1
' Me.ListBox_������.Selected(doc-1) = True
Set shp = wb.Worksheets("������")
Set lo = shp.ListObjects("������")
Set q = lo.ListColumns("��������").DataBodyRange.Cells
Set r = ����������������(q, "D" & doc, 0, 4)
Me.TextBox_����� = r.value
Call ������Word��Excel(FileNameDoc, doc)

End Sub

Private Sub CommandButton14_Click()
Call �����������������
End Sub

Private Sub CommandButton15_Click()
Dim twb As Workbook
Dim wb As Workbook
Dim NameTX As Range
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
Set twb = ThisWorkbook
Path = twb.Path & Application.PathSeparator
Set NameTX = ��������������������("������", "������", "����", "���� ��� �������", "��������")
If fso.FileExists(Path & NameTX.value) Then
Set wb = Workbooks.Open(Path & NameTX.value)
Else
MsgBox "����-" & Path & NameTX.value & "�� ���������"
End If
End Sub

Private Sub CommandButton_������_Click()

Call ����������������
If Me.ComboBox_³�������� = "���������" Then
[f�������_������_���] = [f�������_����_�����]
TextBox_����� = [fm����2] + [fm����3]
TextBox_����� = [fm����5] + [fm����6]
TextBox_����� = [fm����8] + [fm����9]
[f�������] = ""
End If

Me.ComboBox_�������� = Worksheets("�����").Range("f������_����")
Me.ComboBox_Full_Dilyanka = Worksheets("�����").Range("f������ĳ�")
If Me.ComboBox_³�������� = "���������" Then
'Me.TextBox_��������.Value = Worksheets("�����").Range("f�������_������_���")
Else
Sum = 0
'For i = 1 To 9 Step 4
'Sum = Sum + Worksheets("�����").Range("����" & i)
'Next i
'For i = 2 To 10 Step 4
'Sum = Sum + Worksheets("�����").Range("����" & i)
'Next i
For I = 1 To 6
Sum = Sum + Worksheets("�����").Range("����" & I)
Next I

Me.TextBox_�����.value = Sum
Sum = 0
'For i = 3 To 11 Step 4
'Sum = Sum + Worksheets("�����").Range("����" & i)
'Next i
'For i = 4 To 12 Step 4
'Sum = Sum + Worksheets("�����").Range("����" & i)
'Next i
For I = 7 To 12
Sum = Sum + Worksheets("�����").Range("����" & I)
Next I


Me.TextBox_�����.value = Sum
If Worksheets("�����").Range("f�������") = "" Then
Me.TextBox_�����.value = 0
Else
Me.TextBox_�����.value = Worksheets("�����").Range("f�������")
End If


'Me.TextBox_�����.Value = Worksheets("�����").Range("fĳ����")
'Me.TextBox_�����.Value = Worksheets("�����").Range("f�������")
'Me.TextBox_�����.Value = Worksheets("�����").Range("f�����")

'If Me.ComboBox_³�������� = "��������" Then
'Me.TextBox_��������.Value = Worksheets("�����").Range("f�������_������_���")
'Else
'Me.TextBox_��������.Value = Worksheets("�����").Range("f�������_����_�����")
'End If


'��������
Me.TextBox_����������������������� = Worksheets("�����").Range("f��������")
Me.TextBox_������������� = Worksheets("�����").Range("f��������_��")
Me.TextBox_������ = Worksheets("�����").Range("f������__������_��������")
If Val(Worksheets("�����").Range("f�������").value) = 0 Then
Worksheets("�����").Range("f�����_����").value = 0
Else
 Worksheets("�����").Range("f�����_����").value = Worksheets("�����").Range("f�������").value
 End If
 End If
Call UbdateSumm







End Sub

Private Sub CommandButton17_Click()
Call fff
End Sub

Private Sub CommandButton18_Click()
Worksheets("��������").Activate

End Sub

Private Sub CommandButton19_Click()
Worksheets("�����_��").Activate

End Sub

Private Sub CommandButton2_Click()
ListBox_��������.SetFocus
If ListBox_��������.ListIndex = ListBox_��������.ListCount - 1 Then
ListBox_��������.ListIndex = 0
Else
ListBox_��������.ListIndex = ListBox_��������.ListIndex + 1
End If
End Sub

Private Sub CommandButton20_Click()
Me.TextBox_�����.value = 0
Me.TextBox_�����.value = 0
Me.TextBox_�����.value = 0
End Sub

Private Sub CommandButton21_Click()
'ThisWorkbook.Path & "\�������� " & Me.TextBox_���_�������� & "." & Me.ComboBox_����������_��������

Set b = mywbBook("�������� " & Me.TextBox_���_�������� & "." & Me.ComboBox_����������_��������, ThisWorkbook.Path & "\")
If b Is Nothing Then MsgBox ("���� " & twb.Path & "\" & EOBookName)
ThisWorkbook.Save
ThisWorkbook.Close
End Sub

Private Sub CommandButton22_Click()
Call ������_��������
lr = LastRow("��������")
Worksheets("��������").Cells(lr, 10).Select
Me.MultiPage1.value = Me.MultiPage1.value + 1
Me.CommandButton_������.value = True
CommandButton3.value = True
End Sub

Private Sub CommandButton23_Click()
D = FiltUT(ThisWorkbook, "��������", "��������", vf, vc)
End Sub

Private Sub CommandButton3_Click()
Dim shp As Worksheet
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject


 
 NameSheet = "��������"
 
 
 
 Me.ListBox_������.value = "����²� �"
 
'Call CalculationA
'Call CalculationM
'If ActiveSheet.Name <> NameSheet Then Exit Sub
Set sh = Worksheets(NameSheet)

If Selection.Cells.Count = 1 Then
'Call ����������������
Me.Label57.Caption = ��������������������
If fso.FileExists(FileNameDoc) Then Call �������_����������(Worksheets("�����"), Worksheets("�����").Range("fDot"), FileNameDoc, Worksheets("�����").Range("fDot").value) Else Worksheets("�����").Range("fDot").Hyperlinks.Delete
Worksheets(NameSheet).Activate
If fso.FileExists(FileNameDoc) Then Call �������_����������(Worksheets(NameSheet), Worksheets(NameSheet).Cells(Selection.Row, 46), FileNameDoc, FileNameDoc) Else Worksheets(NameSheet).Cells(Selection.Row, 46).Hyperlinks.Delete
Worksheets("�����").Activate
' FilenName = ThisWorkbook.Path & "\" & "��������" & "\" & Range("m" & "NTK") & ".xls"
doc = Me.ListBox_������.ListIndex
If CheckBox_������ Then Call ������Word��Excel(FileNameDoc, doc)
MsgBox FileNameDoc
Else
Set r = Selection
Dim z As Range
For Each z In r.Cells
z.Activate
'Call ����������������


Call ��������������������

If fso.FileExists(FileNameDoc) Then Call �������_����������(Worksheets("�����"), Worksheets("�����").Range("fDot"), FileNameDoc, Worksheets("�����").Range("fDot").value) Else Worksheets("�����").Range("fDot").Hyperlinks.Delete
If fso.FileExists(FileNameDoc) Then Call �������_����������(Worksheets(NameSheet), Worksheets(NameSheet).Cells(Selection.Row, 46), FileNameDoc, FileNameDoc) Else Worksheets(NameSheet).Cells(Selection.Row, 46).Hyperlinks.Delete
' FilenName = ThisWorkbook.Path & "\" & "��������" & "\" & Range("m" & "NTK") & ".xls"
doc = Me.ListBox_������.ListIndex
If Me.ListBox_������.ListIndex = -1 Then
MsgBox "�� ������ ��� ��������� ��� ������"
Exit For
End If
doc = Me.ListBox_������.ListIndex + 1
Set wb = ThisWorkbook
Set shp = wb.Worksheets("������")
Set lo = shp.ListObjects("������")
Dim q As Range
Set q = lo.ListColumns("��������").DataBodyRange.Cells
Set r = ����������������(q, "D" & doc, 0, 4)
Me.TextBox_����� = r.value









If CheckBox_������ Then Call ������Word��Excel(FileNameDoc, doc)
MsgBox FileNameDoc
Set sh = wb.Worksheets(NameSheet)
sh.Activate
Next
End If
End Sub

Private Sub CommandButton4_Click()
Dim EOBookName As String
Dim pth
pth = ThisWorkbook.Path
EOBookName = Me.ComboBox_Lisnuctvo & " " & Me.ComboBox_Full_Dilyanka & ".xls"
'If WorkbookExist(pth, EOBookName) Then
Call ���������������(EOBookName)
Me.MultiPage1.value = 1
'Else
'Call MsgBox("����-" & vbLf & EOBookName & vbLf & "�� ���� " & vbLf & pth & vbLf & "�� ������" & vbLf & "����������� ��� ��� ����������� � ����� � ����������", vbCritical, "���� �� ������")
'
'End If
Me.Label_Log.Caption = "������ ��������� ��������!  ������� ������ ����� "
End Sub

Private Sub CommandButton5_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub CommandButton6_Click()
Select Case Me.MultiPage1.value
Case 0
Me.MultiPage1.value = Me.MultiPage1.value + 1
Me.Label_Log.Caption = "������� �������� ���� � ���� ������� ����� ������ ��� ������� �� ������ � ����� ������� �������� ������ �� ������� ������ �������� "
ThisWorkbook.Worksheets("�����").Activate
Case 1
Me.MultiPage1.value = Me.MultiPage1.value + 1
Me.Label_Log.Caption = "������� �������� ������ ��� ��������������� ������� ������ � ���� ������� ������� ������ ��� ������������ ����� ��������� ������ �� ������� ������ �������� "
ThisWorkbook.Worksheets("������� ��������").Activate
Case 2
Me.MultiPage1.value = Me.MultiPage1.value + 1
Case 3
Me.MultiPage1.value = 0
End Select
End Sub

Private Sub CommandButton7_Click()
Me.Hide
End Sub

Private Sub CommandButton8_Click()
Call ����������������
Range("f����").Activate
Call InsertOrEditTableLink
End Sub

Private Sub CommandButton9_Click()
Call ������_��������
End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label39_Click()

End Sub

Private Sub Label56_Click()

End Sub

Private Sub ListBox_�������_Change()
Range("f������").value = Me.ListBox_�������.value
Me.Label_Log.Caption = "������� ������ ���� "
End Sub

Private Sub ListBox_�������_Click()

End Sub

Private Sub ListBox_������_Click()

End Sub

Private Sub ListBox_��������_Change()
Dim r As Range
Set r = Worksheets("������� ��������").Columns(10).Find(ListBox_��������.value)
Me.TextBox_��������_�����.ControlSource = r.Offset(columnoffset:=-3).address
Me.TextBox_��������_�����.SetFocus

Me.TextBox_��������_�����.SelStart = 0
Me.TextBox_��������_�����.SelLength = Len(Me.TextBox_��������_�����)
End Sub

Private Sub ListBox_��������_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub



Private Sub TextBox_�����������������������_Change()
'��������
Worksheets("�����").Range("f�������_����_�����").value = Me.TextBox_�����������������������.value
Call CalculationA
Call CalculationM
Call UbdateSumm
End Sub

Private Sub TextBox_�������������_Change()
'��������

Worksheets("�����").Range("f�������_������_���").value = Me.TextBox_�������������.value
Call CalculationA
Call CalculationM
End Sub

Private Sub TextBox_���������_Change()
Range("f����������").value = TextBox_���������.value
End Sub

Private Sub TextBox_�����_Change()
'If TextBox_�����.Value <> "" And Me.TextBox_�����.Value <> "" Then Me.TextBox_���������������� = (CDbl(Me.TextBox_�����.Value) + CDbl(Me.TextBox_�����.Value)) * Val(Me.TextBox_��������.Value)

Call UbdateSumm

End Sub



Private Sub TextBox_�����_Change()

End Sub

Private Sub TextBox_�����_Change()
'If Me.TextBox_�����.Value <> "" And Me.TextBox_�����.Value <> "" Then Me.TextBox_���������������� = (CDbl(Me.TextBox_�����.Value) + CDbl(Me.TextBox_�����.Value)) * Val(Me.TextBox_��������.Value)
Call UbdateSumm
End Sub

Private Sub TextBox_�����_Change()

End Sub

Private Sub TextBox_���_Change()

End Sub

Private Sub TextBox_��������_�����_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

'Private Sub TextBox_��������_Change()
'
''If Me.TextBox_�����.Value <> "" And Me.TextBox_�����.Value <> "" Then Me.TextBox_���������������� = (CDbl(Me.TextBox_�����.Value) + CDbl(Me.TextBox_�����.Value)) * Val(Me.TextBox_��������.Value)
'End Sub

Private Sub TextBox_�����_Change()
Call UbdateSumm
'Call CalculationA
'Call CalculationM
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox_���������_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call fDataShow
End Sub

Private Sub TextBox_�����_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
ThisWorkbook.Activate
'For i = 1 To 31
'Me.ComboBox_Day.AddItem i
'Next i
'M = VBA.Month(Now())
'For i = 1 To 12
'Me.ComboBox_Mouth.AddItem i
'Next i
'If M = 1 Then
'Me.ComboBox_Mouth.Value = 12
'k = 12
'Else
'Me.ComboBox_Mouth.Value = M - 1
'k = M - 1
'End If
'd = Year(Now())
'For i = d - 4 To d + 4
'Me.ComboBox_Year.AddItem i
'Next i
'If M = 1 Then
'Me.ComboBox_Year.Value = d - 1
'z = d - 1
'Else
'Me.ComboBox_Year.Value = d
'z = d
'End If
'dn = �����������(k, z)
'Me.ComboBox_Day.Value = dn
Worksheets("�����").Range("fĳ����_����") = 0
Worksheets("�����").Range("f�������_����") = 0
Worksheets("�����").Range("f�����_����") = 0



������������ = ���������������������������������("������", "������", "����� ��� �����")
Me.ComboBox_Report.list = ������������
Me.ComboBox_Report.value = "���������"
z = Me.ComboBox_Lisnuctvo
For I = 1 To Me.ComboBox_Lisnuctvo.ListCount - 1
Me.ComboBox_Lisnuctvo.ListIndex = I
If Me.ComboBox_Lisnuctvo.value <> z Then Exit For
Next I
Me.ComboBox_Lisnuctvo.value = z
Me.MultiPage1.value = 0
'ThisWorkbook.RefreshAll
Me.ComboBox_����_��������.value = ThisWorkbook.Path & "\�������� " & Me.TextBox_���_�������� & "." & Me.ComboBox_����������_��������
End Sub
Sub ����������������()
With fAkt
Set r = FindAll(Worksheets("������� ��������").Cells, "������ ��� ���")
.TextBox_����.value = r.Offset(columnoffset:=8).value

Set r = FindAll(Worksheets("������� ��������").Cells, "���")
.TextBox_���.value = r.Offset(columnoffset:=8).value

Set r = FindAll(Worksheets("������� ��������").Cells, "������ � ���")
.TextBox_������.value = r.Offset(columnoffset:=8).value
End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call CalculationA
'Call CalculationM
End Sub

Sub UbdateSumm()
Call CalculationA
Worksheets("�����").Range("fĳ����_����") = Me.TextBox_�����.value
Worksheets("�����").Range("f�������_����") = Me.TextBox_�����.value
Worksheets("�����").Range("f�����_����") = Me.TextBox_�����.value

Me.TextBox_�����.value = Worksheets("�����").Range("fĳ����")
Me.TextBox_�����.value = Worksheets("�����").Range("f�������")
Me.TextBox_�����.value = Worksheets("�����").Range("f�����")
s = Replace(Worksheets("�����").Range("f��������"), ",", ".")
Me.TextBox_����������������������� = s

Me.TextBox_������������� = s

Me.TextBox_����������.Text = Worksheets("�����").Range("f����������").value
Me.TextBox_���������.Text = Worksheets("�����").Range("f���������").value
Me.TextBox_����������������.Text = Worksheets("�����").Range("f��������������").value

Me.TextBox_����������_��.Text = Worksheets("�����").Range("f��������������").value
Me.TextBox_���������_��.Text = Worksheets("�����").Range("f���_������_��").value
Me.TextBox_����������������_��.Text = Worksheets("�����").Range("f��������������_��").value


Call CalculationM
End Sub
