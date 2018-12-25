Attribute VB_Name = "Module3"
Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.ListObjects("РеестрЛКЕОДConect").Range.AutoFilter Field:=4, _
        Criteria1:="435278"
End Sub
Sub Макрос2()
Attribute Макрос2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос2 Макрос
'

'
    ActiveSheet.ShowAllData
End Sub
Sub Макрос3()
Attribute Макрос3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос3 Макрос
'

'
    ActiveSheet.ListObjects("РеестрЛКЕОДConect").Range.AutoFilter Field:=4
End Sub
Function nearmax1(Число As Double, диапазон As Range)
Dim k, rez As Double
For Each i In диапазон
k = i.value
If rez < k Or rez = Empty Then rez = k
Next i
For Each i In диапазон
If i.value > Число And i.value < rez Then rez = i.value
Next i
If rez <= Число Then rez = Error(1)
nearmax = rez
End Function
Function nearmax(Число As Double, диапазон As Range)
Dim k, rez As Double, rez2 As Double, i
If Число >= диапазон.Cells(1) Then
rez = диапазон.Cells(диапазон.Count)
For i = диапазон.Count - 1 To 1 Step -1
If rez < Число Then Exit For
rez2 = rez
rez = диапазон.Cells(i)

Next i
Else
rez2 = диапазон.Cells(1)
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
Set lo_forma = .items("Форма")
Set loc = lo_forma.ListColumns("")

End With
End Sub

Function bRowHide(r As Range)
bRowHide = r.EntireRow.Hidden
End Function
