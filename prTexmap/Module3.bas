Attribute VB_Name = "Module3"
Sub ������1()
Attribute ������1.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.ListObjects("�����������Conect").Range.AutoFilter Field:=4, _
        Criteria1:="435278"
End Sub
Sub ������2()
Attribute ������2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������2 ������
'

'
    ActiveSheet.ShowAllData
End Sub
Sub ������3()
Attribute ������3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������3 ������
'

'
    ActiveSheet.ListObjects("�����������Conect").Range.AutoFilter Field:=4
End Sub
Function nearmax1(����� As Double, �������� As Range)
Dim k, rez As Double
For Each i In ��������
k = i.value
If rez < k Or rez = Empty Then rez = k
Next i
For Each i In ��������
If i.value > ����� And i.value < rez Then rez = i.value
Next i
If rez <= ����� Then rez = Error(1)
nearmax = rez
End Function
Function nearmax(����� As Double, �������� As Range)
Dim k, rez As Double, rez2 As Double, i
If ����� >= ��������.Cells(1) Then
rez = ��������.Cells(��������.Count)
For i = ��������.Count - 1 To 1 Step -1
If rez < ����� Then Exit For
rez2 = rez
rez = ��������.Cells(i)

Next i
Else
rez2 = ��������.Cells(1)
End If
nearmax = rez2
End Function
Sub ggggjh()
Dim locls As clsmListObjs
Dim lo_forma As ListObject
Dim loc As ListColumn
Dim lo As ListObject
Dim r As Range
Dim wb As Workbook
Set wb = ThisWorkbook
Set locls = New clsmListObjs
With locls
.Initialize wb
Set lo_forma = .items("�����")
Set loc = lo_forma.ListColumns("")

End With
End Sub

Function bRowHide(r As Range)
bRowHide = r.EntireRow.Hidden
End Function
