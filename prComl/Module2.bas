Attribute VB_Name = "Module2"

Sub ���������������()
Attribute ���������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������2 ������
'

Dim twb As Workbook
Dim EOBookSheetNewName As String
Dim pt As PivotTable
Set twb = ThisWorkbook
EOBookSheetNewName = "���������������"
If SheetExistBook(twb, EOBookSheetNewName) Then

Set ws = twb.Worksheets(EOBookSheetNewName)
ws.Delete
End If
Set ws = twb.Worksheets.Add
ws.name = EOBookSheetNewName

 
    
    twb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "��������", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:=ws.Cells(3, 1), TableName:=EOBookSheetNewName, _
        DefaultVersion:=xlPivotTableVersion10
        Set pt = ws.PivotTables(EOBookSheetNewName)
   With pt
    With .PivotFields("���������")
        .Orientation = xlRowField
        .Position = 1
    End With
    With .PivotFields("������")
        .Orientation = xlRowField
        .Position = 2
    End With
    With .PivotFields("�������," & Chr(10) & "L �.")
        .Orientation = xlRowField
        .Position = 3
    End With
    With .PivotFields( _
        "����� �������, D, ��.")
        .Orientation = xlRowField
        .Position = 4
    End With
    With .PivotFields("����")
        .Orientation = xlRowField
        .Position = 5
    End With
  .AddDataField .PivotFields("ʳ������, ��."), _
        "����� �� ���� ʳ������, ��.", xlSum
    .AddDataField .PivotFields("�ᒺ�, " & Chr(10) & "V, �3"), "����� �� ���� �ᒺ�, " & Chr(10) & "V, �3" _
        , xlSum
    With .DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    Range("A4").Select
.PivotFields("���������").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )

 .PivotFields("������").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
 
    With .PivotFields("�������," & Chr(10) & "L �.")
        .Caption = "�������,"
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, _
        False, False, False)
    End With
  
  .PivotFields( _
        "����� �������, D, ��.").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
.DataPivotField.PivotItems( _
        "����� �� ���� ʳ������, ��.").Caption = "ʳ��-��, ��."
    
.DataPivotField.PivotItems( _
        "����� �� ���� �ᒺ�, " & Chr(10) & "V, �3").Caption = "�ᒺ�, " & Chr(10) & ", �3"
    

        .ColumnGrand = False
        .RowGrand = False
   
    
    End With
    
End Sub
Sub ������4()
Attribute ������4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������4 ������
'

'
    
    Range("F5").Select
End Sub
Sub ������5()
Attribute ������5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������5 ������
'

'
    Range("D4").Select
    ActiveSheet.PivotTables("���������������").PivotFields( _
        "����� �������, D, ��.").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
End Sub
Sub ������6()
Attribute ������6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������6 ������
'

'
    Range("A13").Select
    With ActiveSheet.PivotTables("���������������")
        .ColumnGrand = False
        .RowGrand = False
    End With
End Sub
