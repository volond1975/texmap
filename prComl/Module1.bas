Attribute VB_Name = "Module1"
Type �������������
�� As String
��� As String
ĳ� As String
End Type
Dim MX As �������������
Function ��������������(�����_������)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(�����_������)) Then
    With dict
  .Add "1", "ѳ����"
  .Add "2", "�����"
  .Add "3", "��������"
  .Add "4", "������"
  .Add "5", "�������"
  .Add "6", "�������"
  .Add "7", "������"
  .Add "8", "�������"
  .Add "9", "��������"
  .Add "10", "�������"
  .Add "11", "��������"
  .Add "12", "�������"
  
  
  
    End With
    
    
    
End If

�������������� = dict(Trim(str(�����_������)))
End Function
Function ��������������(�����_������)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(�����_������)) Then
    With dict
  .Add "1", "ѳ���"
  .Add "2", "������"
  .Add "3", "�������"
  .Add "4", "�����"
  .Add "5", "������"
  .Add "6", "������"
  .Add "7", "�����"
  .Add "8", "������"
  .Add "9", "�������"
  .Add "10", "������"
  .Add "11", "���������"
  .Add "12", "������"
  
  
  
    End With
    
    
    
End If

�������������� = dict(Trim(str(�����_������)))
End Function
Function ��������������(�����_������)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(�����_������)) Then
    With dict
  .Add "1", "ѳ��"
  .Add "2", "������"
  .Add "3", "������"
  .Add "4", "����"
  .Add "5", "�����"
  .Add "6", "�����"
  .Add "7", "����"
  .Add "8", "�����"
  .Add "9", "������"
  .Add "10", "�����"
  .Add "11", "��������"
  .Add "12", "�����"
  
  
  
    End With
    
    
    
End If

�������������� = dict(Trim(str(�����_������)))
End Function
Function �����������(�����_������, �����_����)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(�����_������)) Then
    With dict
  .Add "1", "31"
  If IsDate("02/29/" & �����_����) = True Then
  
  .Add "2", "29"
 Else
 .Add "2", "28"
 End If
  .Add "3", "31"
  .Add "4", "30"
  .Add "5", "31"
  .Add "6", "30"
  .Add "7", "31"
  .Add "8", "31"
  .Add "9", "30"
  .Add "10", "31"
  .Add "11", "30"
  .Add "12", "31"
  
  
  
    End With
    
    
    
End If

����������� = dict(Trim(str(�����_������)))
End Function

Function ������������(ĳ����� As String)


Call ������ĳ�������(ĳ�����)
v = Array(MX.��, MX.���, MX.ĳ�)
������������ = v
End Function
Sub ������ĳ�������(ĳ����� As String)
'66 �� (1, 2 ���)  ��.
'79 �� (12, 13, 15, 16 ���)  ��.
'27 �� (5 ���) 0 ��.



Dim v
If ĳ����� = "" Then Exit Sub


v = Split(ĳ�����, "(")
z = Split(v(1), ")")
MX.�� = Val(v(0))
If z(0) Like "*,*" Then
MX.��� = Left(z(0), Len(z(0)) - 4)
Else
MX.��� = Val(z(0))
End If

MX.ĳ� = Val(z(1))
'������ĳ������� = MX
End Sub

Function ����������������(������� As String)
'�� ������� �� �������� (�����-����)  � 09\��\1_9 �� 02.09.2013 ����
'������� As String


Dim v

v = Split(�������, "�")
z = Split(VBA.Trim(v(1)), " ")
'0-����� ��������
'1-��
'2-���� ��������
'3-����
���������������� = z
End Function


Function ����������������������(������_�_���������)
Dim ������_�_��������� As String
Dim fDogovor As Range
Dim fDilyanka As Range
Dim columnDogovor As Range
Dim columnDilyanka As Range
Dim shDogovor As Worksheet
'�������� �� ������� ����� ������ ��������
Set shDogovor = ThisWorkbook.Worksheets("������_��������")
mDogovor = ����������������(������_�_���������)
������� = mDogovor(0)
With shDogovor
Set fDogovor = .Cells.Find("����� ��������")
Set columnDogovor = .Columns(fDogovor.Column).Cells
Set fDilyanka = .Cells.Find("ĳ�����")
Set columnDilyanka = .Columns(fDilyanka.Column).Cells
'ĳ�����
Set fDogovor = .Cells.Find(�������)
Set fDilyanka = Application.Intersect(columnDilyanka, .Rows(fDogovor.Row))
���������������������� = fDilyanka.value
End With
End Function

Sub fff()
vf = Array("2") ', "8", "9"
If fAkt.ComboBox_ĳ����� = 0 Then
z = fAkt.ComboBox_����
Else
z = fAkt.ComboBox_���� & "_" & fAkt.ComboBox_ĳ�����
End If

vc = Array(fAkt.ComboBox_Lisnuctvo) ', fAkt.ComboBox_�������, z
D = FiltUT(ThisWorkbook, "��������", "��������", vf, vc)

'vf = Array("1", "4")
'vc = Array(fAkt.ComboBox_Lisnuctvo, fAkt.ComboBox_Full_Dilyanka)
'D = FiltUT(ThisWorkbook, "������_��", "������_��", vf, vc)

End Sub

Function FiltUT(twb As Workbook, ShName, UtNAme, vf, vc)
Dim sh As Worksheet
Dim lo As ListObject
On Error Resume Next
Set lo = twb.Worksheets(ShName).ListObjects(UtNAme)

With lo
If IsArray(vf) Then
lo.Range.AutoFilter
lo.Range.AutoFilter
For I = 0 To UBound(vf)
.Range.AutoFilter Field:=Val(vf(I)), Criteria1:=vc(I)
Next I
Else
lo.Range.AutoFilter
lo.Range.AutoFilter
End If

End With


End Function
