Attribute VB_Name = "Module2"

Sub СводнаяЩоденник()
Attribute СводнаяЩоденник.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос2 Макрос
'

Dim twb As Workbook
Dim EOBookSheetNewName As String
Dim pt As PivotTable
Set twb = ThisWorkbook
EOBookSheetNewName = "ЗведенаЩоденник"
If SheetExistBook(twb, EOBookSheetNewName) Then

Set ws = twb.Worksheets(EOBookSheetNewName)
ws.Delete
End If
Set ws = twb.Worksheets.Add
ws.name = EOBookSheetNewName

 
    
    twb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Щоденник", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:=ws.Cells(3, 1), TableName:=EOBookSheetNewName, _
        DefaultVersion:=xlPivotTableVersion10
        Set pt = ws.PivotTables(EOBookSheetNewName)
   With pt
    With .PivotFields("Сортимент")
        .Orientation = xlRowField
        .Position = 1
    End With
    With .PivotFields("Порода")
        .Orientation = xlRowField
        .Position = 2
    End With
    With .PivotFields("Довжина," & Chr(10) & "L м.")
        .Orientation = xlRowField
        .Position = 3
    End With
    With .PivotFields( _
        "Група діаметрів, D, см.")
        .Orientation = xlRowField
        .Position = 4
    End With
    With .PivotFields("Сорт")
        .Orientation = xlRowField
        .Position = 5
    End With
  .AddDataField .PivotFields("Кількість, шт."), _
        "Сумма по полю Кількість, шт.", xlSum
    .AddDataField .PivotFields("Об’єм, " & Chr(10) & "V, м3"), "Сумма по полю Об’єм, " & Chr(10) & "V, м3" _
        , xlSum
    With .DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    Range("A4").Select
.PivotFields("Сортимент").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )

 .PivotFields("Порода").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
 
    With .PivotFields("Довжина," & Chr(10) & "L м.")
        .Caption = "Довжина,"
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, _
        False, False, False)
    End With
  
  .PivotFields( _
        "Група діаметрів, D, см.").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
.DataPivotField.PivotItems( _
        "Сумма по полю Кількість, шт.").Caption = "Кіль-ть, шт."
    
.DataPivotField.PivotItems( _
        "Сумма по полю Об’єм, " & Chr(10) & "V, м3").Caption = "Об’єм, " & Chr(10) & ", м3"
    

        .ColumnGrand = False
        .RowGrand = False
   
    
    End With
    
End Sub
Sub Макрос4()
Attribute Макрос4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос4 Макрос
'

'
    
    Range("F5").Select
End Sub
Sub Макрос5()
Attribute Макрос5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос5 Макрос
'

'
    Range("D4").Select
    ActiveSheet.PivotTables("ЗведенаЩоденник").PivotFields( _
        "Група діаметрів, D, см.").Subtotals = Array(False, False, False, False, False, _
        False, False, False, False, False, False, False)
End Sub
Sub Макрос6()
Attribute Макрос6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос6 Макрос
'

'
    Range("A13").Select
    With ActiveSheet.PivotTables("ЗведенаЩоденник")
        .ColumnGrand = False
        .RowGrand = False
    End With
End Sub
