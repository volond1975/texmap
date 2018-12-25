Attribute VB_Name = "mСводная"

Sub СводнаяДляОстатков()
Attribute СводнаяДляОстатков.VB_Description = "Макрос записан 12.05.2011 (Владелец)"
Attribute СводнаяДляОстатков.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ABS_WB As Workbook
Dim Ost_lst As Worksheet
Dim SV_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Dim pi As PivotItem
Set ABS_WB = ThisWorkbook


Call New_Multi_Table_Pivot
Call CellAutoFilterVisible("3")
Set pc = ABS_WB.PivotCaches(1)
Set Ost_lst = ABS_WB.Worksheets("Остатки")
Set SV_lst = ABS_WB.Worksheets("Сводная")
Set pt = SV_lst.PivotTables("CV_Ob")

        Dim r As Range
        Set r = Worksheets("Оглавление").Range("B1")
    Set EndPeriodM = r.End(xlDown)
    Set EndPeriodY = EndPeriodM.Offset(columnoffset:=1)
    With pt.PivotFields("Год")
    For Each pi In .PivotItems
      
       pi.Visible = True

 Next
    For Each pi In .PivotItems
       If pi.value = EndPeriodY.value Then
       pi.Visible = True
Else
pi.Visible = False
End If

        
        Next
    End With
    
    With pt.PivotFields("Месяц")
     For Each pi In .PivotItems
     
       pi.Visible = True

Next
       For Each pi In .PivotItems
       If pi.value = Val(EndPeriodM.value) Then
       pi.Visible = True
Else
pi.Visible = False
End If
Next
    End With
' Сall lst_clear_cell(Ost_lst)
 Ost_lst.Cells.Clear
''Ost_lst.Cells.Font.ColorIndex = xlNone
'Ost_lst.Cells.Interior.ColorIndex = xlNone
 
 
 SV_lst.Cells.Copy
'    Sheets("Сводная").Select
'    Cells.Select
'    Selection.Copy
     
   Ost_lst.Select
   Cells.Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'    Rows("1:4").Select
'    Range("A4").Activate
    Application.CutCopyMode = False
'    Selection.Delete Shift:=xlUp
    Rows("1:4").Delete Shift:=xlUp
End Sub
Sub PivotFieldOrientation(PF As PivotField)
Attribute PivotFieldOrientation.VB_Description = "Макрос записан 12.05.2011 (Владелец)"
Attribute PivotFieldOrientation.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос3 Макрос
' Макрос записан 12.05.2011 (Владелец)
'
Select Case PF.Orientation
    Case xlHidden
        MsgBox "Hidden field"
    Case xlRowField
        MsgBox "Row field"
    Case xlColumnField
        MsgBox "Column field"
    Case xlPageField
        MsgBox "Page field"
    Case xlDataField
        MsgBox "Data field"
End Select

  
End Sub
Sub ColumnsAutoFit(r As Range)
Attribute ColumnsAutoFit.VB_Description = "Макрос записан 12.05.2011 (Владелец)"
Attribute ColumnsAutoFit.VB_ProcData.VB_Invoke_Func = " \n14"
'
' АвтоПодборШирины
'

    r.Columns.AutoFit
   
   
End Sub
Sub RowsAutoFit(r As Range)
'
' АвтоПодборВысоты
'

    r.Rows.AutoFit
 
   
End Sub
Sub АвтоПодборПоВыделению()
Dim r As Range
Set r = Selection
If r.Columns.Count < r.Rows.Count Then
Call ColumnsAutoFit(r)
Else
Call RowsAutoFit(Selection)
End If
End Sub
Sub ОставитьОдноЗначениеВСтолбцеСводной(pt As PivotTable, НаименованиеПоля As String, ТребуемоеЗначение As String)
On Error Resume Next
Dim pi As PivotItem

Call Intro
pt.ManualUpdate = True
k = 0

    With pt.PivotFields(НаименованиеПоля)
    For Each pi In .PivotItems
    Select Case ТребуемоеЗначение
    Case ""
    pi.Visible = True
    Case Else
       If pi.value = ТребуемоеЗначение Then
       pi.Visible = True
       k = k + 1
Else
pi.Visible = False
'If Err = 1004 Then MsgBox Err
End If
End Select
Next
Call Outro
pt.ManualUpdate = False
End With
End Sub

Sub АктивироватьОдноЗначениеВСтолбцеСводной(НаименованиеПоля As String, ТребуемоеЗначение As String)
Dim ABS_WB As Workbook
Dim O_lst As Worksheet
Dim SV_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Dim pi As PivotItem

Set ABS_WB = ThisWorkbook
Call Intro
НаименованиеПоля = "Наименование"
ТребуемоеЗначение = "Хлорофіліпт 2% 25мл олійний р-н+"
Set SV_lst = ABS_WB.Worksheets("Сводная")
Set pc = ABS_WB.PivotCaches(1)
Set pt = SV_lst.PivotTables("CV_Ob")
Cells(1, 1).Select
ActiveWindow.SmallScroll Up:=65000

     pt.PivotSelection = НаименованиеПоля & "[" & ТребуемоеЗначение & "]"
    
    ActiveWindow.SmallScroll Down:=ActiveCell.Row - 6
Call Outro

End Sub


Sub PVTableRefresh()
Dim ABS_WB As Workbook
Dim O_lst As Worksheet
Dim SV_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Set ABS_WB = ThisWorkbook

Set O_lst = ABS_WB.Worksheets("Общий Лист")
Set SV_lst = ABS_WB.Worksheets("Сводная")
lr = LastRow(O_lst.name)
col = 12
sd = Chr(39) & "" & O_lst.name & "'!R1C1:R" & lr & "C" & col
Set pc = ABS_WB.PivotCaches(1)
'pc.SourceData = "''" & O_lst.Name & "'!" & O_lst.Range("A1").CurrentRegion.Address
pc.SourceData = sd
Debug.Print pc.SourceData
Debug.Print sd
Set pt = SV_lst.PivotTables("CV_Ob")


End Sub
Sub ДобавитьИлиУдалитьИтогиПоСтолбцамВСводной(pt As PivotTable, zn As Boolean)
'
' УдалитьИтогиПоСтолбцамВСводной Макрос
' Макрос записан 16.05.2011 (Владелец)
'

'
    With pt
        .ColumnGrand = zn
        
    End With
End Sub
Sub ДобавитьИлиУдалитьИтогиПоСтрокамВСводной(pt As PivotTable, zn As Boolean)
'
' УдалитьИтогиПоСтолбцамВСводной Макрос
' Макрос записан 16.05.2011 (Владелец)
'

'
    With pt
                .RowGrand = zn
    End With
End Sub
Sub New_Multi_Table_Pivot()
ResultSheetName = fСводная.ComboBox1.value
ResultPivotTableName = fСводная.ComboBox1.value & "_ALL"
Шаблон = "ТАБЛИЦ1С"
Call New_Multi_Table_Pivot1(ResultSheetName, ResultPivotTableName, Шаблон)
End Sub


Sub New_Multi_Table_Pivot1(ResultSheetName, ResultPivotTableName, Шаблон)
    Dim I As Long
    Dim arSQL() As String
    Dim objPivotCache As PivotCache
    Dim objRS As Object
'    Dim ResultSheetName As String
    Dim SheetsNames As String
    Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim NEW_lst As Worksheet
Dim АK_lst As Worksheet
Dim АK_Cell As Range
Dim Find_Cell As Range

Set ABS_WB = ThisWorkbook
Set АK_Cell = ActiveCell
'Set O_lst = ABS_WB.Worksheets("Общий Лист")

  
    'имя листа, куда будет выводиться результирующая сводная
'    ResultSheetName = "СВ_Спецификация"
'    ResultPivotTableName = "СВ_Спецификация"
'    'массив имен листов с исходными таблицами
'For Each NEW_lst In ABS_WB.Worksheets
'If NEW_lst.Name Like "*_*" Then
'SheetsNames = SheetsNames & "SELECT * FROM [" & NEW_lst.Name & "$] UNION ALL "
'
'End If
'    Next
'  SheetsNames = Trim(SheetsNames)
'  SheetsNames = Left(SheetsNames, Len(SheetsNames) - 10)
'Шаблон = "Таблица"
  SheetsNames = ArraySheetName(ABS_WB, Шаблон)
    'формируем кэш по таблицам с листов из SheetsNames
    With ABS_WB
       Set objRS = CreateObject("ADODB.Recordset")
objRS.Open SheetsNames, _
Join$(Array("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=", _
.FullName, ";Extended Properties=""Excel 8.0;"""), vbNullString)
End With
 'создаем заново лист для вывода результирующей сводной таблицы
On Error Resume Next
Application.DisplayAlerts = False
Worksheets(ResultSheetName).Delete
Set wsPivot = SheetExistBookCreate(ThisWorkbook, ResultSheetName, False)

 
  
   
    
  
    'выводим на этот лист сводную по сформированному кэшу
    Set objPivotCache = ActiveWorkbook.PivotCaches.Add(xlExternal)
    Set objPivotCache.Recordset = objRS
    Set objRS = Nothing
    With wsPivot
        objPivotCache.CreatePivotTable TableDestination:=wsPivot.Range("A3")
       
Set pt = .PivotTables(1)
pt.name = ResultPivotTableName
        Set objPivotCache = Nothing
        Range("A3").Select
    End With
    k = 1
    For Each PF In pt.PivotFields
'PF.Name = O_lst.Cells(1, k)
'k = k + 1

Debug.Print PF.name
Next
'  Call CellAutoFilterVisible(1)
    
    
Application.DisplayAlerts = True
End Sub

Sub PVTableAddABC()
Dim ABS_WB As Workbook
Dim O_lst As Worksheet
Dim SV_lst As Worksheet
Dim ABC_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Dim PF As PivotField
Set ABS_WB = ThisWorkbook

Set O_lst = ABS_WB.Worksheets("Общий Лист")
Set SV_lst = ABS_WB.Worksheets("Сводная")
Set ABC_lst = ABS_WB.Worksheets("АВС")

'ABC_lst.Cells.Clear
Set pc = ABS_WB.PivotCaches(1)
Set pt = SV_lst.PivotTables("CV_Ob")
k = 1
For Each PF In pt.PivotFields
 PF.name = O_lst.Cells(1, k)
 k = k + 1
' pf.Position
Next
'lr = LastRow(O_lst.Name)
'col = 12
'sd = Chr(39) & "" & O_lst.Name & "'!R1C1:R" & lr & "C" & col
'Set pc = ABS_WB.PivotCaches(1)
''pc.SourceData = "''" & O_lst.Name & "'!" & O_lst.Range("A1").CurrentRegion.Address
'pc.SourceData = sd
'Debug.Print pc.SourceData
'Debug.Print sd
'Set pt = SV_lst.PivotTables("CV_Ob")
'pt.DataFields (1)

End Sub

Sub CellAutoFilterVisible(k)
Dim sh As Worksheet
Dim pt As PivotTable
Dim PF As PivotField
Dim PFС As PivotField
Set sh = Worksheets("ФорматСводных")
Dim sh_cv As Worksheet
'On Error Resume Next
z = 0
With sh
.Rows(1).AutoFilter Field:=1, Criteria1:=k
     If .AutoFilterMode = True And .FilterMode = True Then
        With .AutoFilter.Range.Columns(1)
             Set iFilterRange = _
             .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlVisible)
             cm = 2
             For Each iCell In iFilterRange
             
             Set sh_cv = SheetExistBookCreate(ThisWorkbook, iCell.Offset(columnoffset:=1).value, False)
             Set pt = sh_cv.PivotTables(iCell.Offset(columnoffset:=2).value)
              pt.DisplayErrorString = True
             pt.ErrorString = ""
             pt.ColumnGrand = iCell.Offset(columnoffset:=2).Font.Bold
             
           If iCell.Offset(columnoffset:=3).Font.Underline = 2 Then
pt.RowGrand = True
Else
pt.RowGrand = False
      End If
             If r = 0 Then
             For Each PF In pt.PivotFields
             If Not PF.Orientation = 0 Then PF.Orientation = xlHidden
             Next
             r = 1
             End If
             
             If iCell.Offset(columnoffset:=10).value <> "" Then
             'Помещаем в область страниц
             pt.PivotFields(iCell.Offset(columnoffset:=10).value).Orientation = xlPageField

 If CommentExist(iCell.Offset(columnoffset:=10)) Then
        'Ограничеваем значением содержащимся в коментарии если есть
        comentar = CommentTEXT(iCell.Offset(columnoffset:=10))
         pt.PivotFields(iCell.Offset(columnoffset:=10).value).CurrentPage = _
         "" & comentar & ""
         End If
        End If
    
             If iCell.Offset(columnoffset:=3).value <> "" Then
            'Помещаем в область строк
             pt.PivotFields(iCell.Offset(columnoffset:=3).value).Orientation = xlRowField
        If Not iCell.Offset(columnoffset:=3).Font.Bold Then
        pt.PivotFields(iCell.Offset(columnoffset:=3).value).Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
        End If
             End If
             If iCell.Offset(columnoffset:=4).value <> "" Then
              'Помещаем в область столбцов
             pt.PivotFields(iCell.Offset(columnoffset:=4).value).Orientation = xlColumnField
             If Not iCell.Offset(columnoffset:=4).Font.Bold Then
        pt.PivotFields(iCell.Offset(columnoffset:=4).value).Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
        End If
             
             End If
             
           If iCell.Offset(columnoffset:=16).value <> "" Then
         Formula = "=" & iCell.Offset(columnoffset:=15).value
'      Formula = "=Сума/'Об''єм'"
 Set pfc = pt.CalculatedFields.Add(iCell.Offset(columnoffset:=14).value _
        , Formula, True)
    pt.PivotFields(iCell.Offset(columnoffset:=14).value). _
        Orientation = ПоложениеСтолбца(iCell.Offset(columnoffset:=16).value)
         
         
         
         
         
         
           
             End If
             
             
             
             
             
             If iCell.Offset(columnoffset:=5).value <> "" Then
             If z = 0 Then
             For I = 1 To pt.DataFields.Count
             pt.DataFields(1).Orientation = xlHidden
            
             Next I
             z = 1
             End If
        
             
             
             
             
             
              
             If iCell.Offset(columnoffset:=5).value Like "*Сумма по*.*" Then
             l = InStr(1, iCell.Offset(columnoffset:=5).value, ".")
             Dim t As String
             t = iCell.Offset(columnoffset:=5).Text
            Mid(t, l, 1) = "#"
            Set PF = pt.PivotFields(t)
             Else
             Set PF = pt.PivotFields(iCell.Offset(columnoffset:=5).value)
             End If
            
              
             
             
             
             Call AddDataFildConsolidac(pt, PF, iCell.Offset(columnoffset:=6).value, iCell.Offset(columnoffset:=7).value)
             
           


             
             
             
             
             End If
              If iCell.Offset(columnoffset:=13).value <> "" Then
             pt.Format (iCell.Offset(columnoffset:=13).value)
             End If
                Debug.Print iCell.Row & "_" & iCell.Offset(columnoffset:=13).value
             Next
        End With
'        .ShowAllData 'Отобразить всё - необязательно
     End If
End With
End Sub
Sub AddDataFildConsolidac(pt As PivotTable, PF As PivotField, Zagolovok As String, CF As String)
'ФункцияСводной Значение
'Сумма xlSum
'Количество xlCount
'Среднее xlAverage
'Максимум xlMax
'Минимум xlMin
'Количество чисел    xlCountNums
'Произведение xlProduct
'Смещенное отклонение    xlStDev
'Несмещенное отклонение  xlStDevP
'Неизвестный xlUnknown
'Смещенная дисперсия xlVar
'Несмещенная дисперсия   xlVarP

             
             
             Select Case CF
                Case "Сумма"
              pt.AddDataField PF, Zagolovok, xlSum
                Case "Количество"
              pt.AddDataField PF, Zagolovok, xlCount
                Case "Среднее"
              pt.AddDataField PF, Zagolovok, xlAverage
                Case Else

            End Select
End Sub
  Function ArraySheetName(ABS_WB As Workbook, Таблица)
  Dim NEW_lst As Worksheet
  Dim SheetsNames As String
  'массив имен листов с исходными таблицами
For Each NEW_lst In ABS_WB.Worksheets
If NEW_lst.name Like Таблица Then
SheetsNames = SheetsNames & "SELECT * FROM [" & NEW_lst.name & "$] UNION ALL "
 
End If
    Next
  SheetsNames = Trim(SheetsNames)
  SheetsNames = Left(SheetsNames, Len(SheetsNames) - 10)
  ArraySheetName = SheetsNames
  End Function
  
Sub СписокПолей()
' ===========================================
' Скрыть или отобразить список полей сводной
' ===========================================
'

If ActiveWorkbook.ShowPivotTableFieldList Then
    ActiveWorkbook.ShowPivotTableFieldList = False
    Else
    ActiveWorkbook.ShowPivotTableFieldList = True
    End If
End Sub
Sub ФормулуСуммыВСтолбец(r As Range, col As Long)
'
' Макрос11 Макрос
' Макрос записан 25.05.2011 (Владелец)
Dim b As Range

Set b = LastColumn(ActiveSheet.name, 1)
Col_r = r.Column
NameCol = Cells(1, Col_r)
NamPol = "Сумма по " & NameCol



lr = LastRow(ActiveSheet.name)
    Cells(1, col).FormulaR1C1 = NamPol
    Set z = ЗаголовокСтолбца(ThisWorkbook, ActiveSheet.name, "Цена")
gABC = z.Column

formul = "=RC[-" & col - gABC & "]*RC[" & Col_r - col & "]"

    Range(Cells(2, col), Cells(lr, col)).FormulaR1C1 = formul
    
End Sub
Sub УбратьНулиВоВсех()
Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim NEW_lst As Worksheet
Dim АK_lst As Worksheet
Dim АK_Cell As Range
Dim Find_Cell As Range
Set ABS_WB = ThisWorkbook
Set АK_Cell = ActiveCell

For Each NEW_lst In ABS_WB.Worksheets

If NEW_lst.name Like "*_*" Then
NEW_lst.Activate

Call УбратьНули
End If
    Next
End Sub
Sub ДобавитьНулиВоВсех()
Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim NEW_lst As Worksheet
Dim АK_lst As Worksheet
Dim АK_Cell As Range
Dim Find_Cell As Range
Set ABS_WB = ThisWorkbook
Set АK_Cell = ActiveCell

For Each NEW_lst In ABS_WB.Worksheets

If NEW_lst.name Like "*_*" Then
NEW_lst.Activate

Call ДобавитьНули
End If
    Next
End Sub
Sub УбратьНули()
Dim b As Range
Dim Y As String
'v = Array("Нач.ост.к-во", "Прих.к -во", "Расх.к-во", "Ост.к-во")
v = Array(4, 5, 6, 7)
lr = LastRow(ActiveSheet.name)
For I = 0 To UBound(v)
'y = v(i)
'Set z = ЗаголовокСтолбца(ThisWorkbook, ActiveSheet.Name, y)
'gABC = z.Column
gABC = v(I)
Set r = Range(Cells(2, gABC), Cells(lr, gABC))
q = r
For W = 1 To UBound(q)
If q(W, 1) = 0 Then q(W, 1) = ""
Next W
r.value = q
Next I




    
End Sub

Sub ДобавитьНули()
Dim b As Range
Dim Y As String
'v = Array("Нач.ост.к-во", "Прих.к -во", "Расх.к-во", "Ост.к-во")
v = Array(4, 5, 6, 7)
lr = LastRow(ActiveSheet.name)
For I = 0 To UBound(v)
'y = v(i)
'Set z = ЗаголовокСтолбца(ThisWorkbook, ActiveSheet.Name, y)
'gABC = z.Column
gABC = v(I)
Set r = Range(Cells(2, gABC), Cells(lr, gABC))
q = r
For W = 1 To UBound(q)
If q(W, 1) = "" Then q(W, 1) = 0
Next W
r.value = q
Next I




    
End Sub



Sub СводнаяВ_АВС()
SheetName = "АВС"
RowsDelete = 4
Call СводнаяВ_Лист(SheetName, RowsDelete)
End Sub
Sub СводнаяRCВ_АВС()
SheetName = "АВС"
RowsDelete = 4
Call СводнаяВ_Лист(SheetName, RowsDelete)
End Sub

Sub СводнаяВ_Лист(CVSheetName, SheetName, RowsDelete)
'
' Макрос11 Макрос
' Макрос записан 24.05.2011 (Владелец)
'
Dim shcopy As Worksheet
Dim shpaste As Worksheet
Dim pt As PivotTable
Dim rf As Object
Dim v() As String
Dim rr As Range
Call Intro
Set shpaste = SheetExistBookCreate(ThisWorkbook, SheetName, True)
'Set shpaste = Worksheets(SheetName)
Set shcopy = Worksheets(CVSheetName)
    shpaste.Select
    Cells.Select
    Selection.ClearContents
    shcopy.Select
    Cells.Select
    Selection.Copy
    Sheets(SheetName).Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Application.CutCopyMode = False
    
    Rows("1:" & RowsDelete).Delete Shift:=xlUp
'    Columns("D:D").Select
'    With Selection.Font
'        .Name = "Georgia"
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'    End With
   Selection.Interior.ColorIndex = xlNone
   Selection.Font.ColorIndex = 0
   Rows(RowsDelete - 2 & ":" & RowsDelete - 2).AutoFilter
   
  Set pt = shcopy.PivotTables(1)
' RowsDelete = pt.PivotFields
 ActiveWindow.FreezePanes = False
 Cells(2, pt.RowFields.Count + 1).Select
' ActiveWindow.FreezePanes = True
 Set rr = Cells(1, (pt.RowFields.Count))
 
 Range(rr, rr.End(xlToLeft)).Interior.ColorIndex = 9
 Range(rr, rr.End(xlToLeft)).Font.ColorIndex = 2
 Range(rr, rr.End(xlToLeft)).Font.Bold = True
 Set rr = Cells(1, (pt.RowFields.Count + 1))
 
 Range(rr, rr.End(xlToRight)).Interior.ColorIndex = 37
 Range(rr, rr.End(xlToRight)).Font.Bold = True
 ReDim v(pt.RowFields.Count - 1)
 For g = 0 To pt.RowFields.Count - 1
 v(g) = Cells(2, g + 1).value
 Next g
 lr = LastRow(shpaste.name) ' - RowsDelete
' Set b = LastColumn(sh.Name, 1)
 Set z = shpaste.Range(shpaste.Cells(2, 1), shpaste.Cells(lr, pt.RowFields.Count))
 X = z
 For I = 1 To UBound(X, 1)
 For j = 1 To UBound(X, 2)
 If X(I, j) = "" Then X(I, j) = v(j - 1)
 If X(I, j) <> v(j - 1) Then v(j - 1) = X(I, j)
 Next j
 Next I
z.value = X
 Call Outro
End Sub
Sub КСПВСтолбец(r As Range, col As Long)
'
' Коеффициент Стабильности Продаж в Столбец АВС
' Макрос записан 25.05.2011 (Владелец)

Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim Nel_lst As Worksheet
Dim O_lst As Worksheet
Dim ABS_lst As Worksheet
Dim nr As Range
Dim f As Range
Dim b As Range
Set ABS_WB = ThisWorkbook
Set ABS_lst = ABS_WB.ActiveSheet
Set b = LastColumn(ABS_lst.name, 1)
Set zags = ABS_lst.Range(ABS_lst.Cells(1, 1), b)
Set zag = ЗаголовокСтолбца(ABS_WB, ABS_lst.name, "кСтбПродаж")
If zag Is Nothing Then
b.Offset(columnoffset:=1).value = "кСтбПродаж"
Set b = b.Offset(columnoffset:=1)
End If
lr = LastRow(ABS_lst.name)






    Cells(1, col).FormulaR1C1 = NamPol
    Set z = ЗаголовокСтолбца(ThisWorkbook, ActiveSheet.name, "кСтбПродаж")
gABC = z.Column

formul = "=RC[-" & col - gABC & "]*RC[" & Col_r - col & "]"

    Range(Cells(2, col), Cells(lr, col)).FormulaR1C1 = formul
    
End Sub
Sub ВыделениеВСводнойПоАктивнойЯчейке()
'
' МНожВыделение Макрос
' Макрос записан 11.07.2011 (Владелец)
'

Dim r As Range
Dim r_col As Range
Set r = ActiveCell
Set r_col = Cells(1, r.Column)
Dim psname As String
psname = "'" & r_col.value & "'"
psname = psname & "['" & r.value & "']"
    Worksheets("Сводная").PivotTables("CV_Ob").PivotSelection = psname
End Sub
Sub ВыделениеВСводнойПоАктивнойЯчейкеAll()
'
' МНожВыделение Макрос
' Макрос записан 11.07.2011 (Владелец)
'

Dim r As Range
Dim r_col As Range
Set r = ActiveCell
Set r_col = Cells(1, r.Column)
Dim psname As String
psname = "'" & r_col.value & "'"
psname = psname & "[" & "ALL" & "]"
    Worksheets("Сводная").PivotTables("CV_Ob").PivotSelection = psname
End Sub


Sub ПереносВСводнойПоПозицииЗаданойВНоменклатуре()
Dim wb As Workbook
Dim shNamenkl As Worksheet
Dim r As Range
Dim r_col As Range
Dim psname As String
Dim NameZag As String
Dim ColName As String
Set wb = ActiveWorkbook

Set shNamenkl = wb.Worksheets("Номерклатура")
Set r = ActiveCell
Set r_col = Cells(1, r.Column)
ColName = r_col.value
psname = PivotSelectionName(ColName, r.value)
NameZag = "Позиция"
NamePoz = r.value
poz_row = r.Row
Set poz_col = ЗаголовокСтолбца(wb, shNamenkl.name, NameZag)

Poz = shNamenkl.Cells(poz_row, poz_col.Column)


Call ПереносВСводнойПоПозицииЗаданойЯчейкой(psname, ColName, Poz, NamePoz)
End Sub

Function PivotSelectionName(ColName As String, it As String) As String
Dim psname As String
psname = "'" & ColName & "'"
'psname = ColName

psname = psname & "['" & it & "']"
PivotSelectionName = psname

End Function

Sub ПереносВСводнойПоПозицииЗаданойЯчейкой(psname As String, ColName As String, Poz As Variant, NamePoz)
Dim wb As Workbook
Dim shSvodnay As Worksheet
Dim pt As PivotTable
Dim PF As PivotField
Dim pi As PivotItem
On Error Resume Next
Set wb = ActiveWorkbook

Set shSvodnay = wb.Worksheets("Сводная")
Set pt = shSvodnay.PivotTables("CV_Ob")
Set PF = pt.PivotFields(ColName)
Set pi = PF.PivotItems(NamePoz)

Call СортировкаПоляСводнойВРучную(PF)
pi.Visible = True
 pt.PivotSelection = psname
 If Err = 1004 Then
 Else
 pi.Visible = True
 pt.PivotSelect "'" & NamePoz & "'", xlLabelOnly, True
 pi.Position = Poz
 
 End If

End Sub
Function ArrayFiltrVisible()



Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim Nel_lst As Worksheet

 
Dim O_lst As Worksheet
Dim ABS_lst As Worksheet
Dim nr As Range
Dim f As Range
Dim v() As Range
Set ABS_WB = ActiveWorkbook
Set ABS_lst = ABS_WB.ActiveSheet
'Set B = LastColumn(ABS_lst.Name, 1)
'Set zags = ABS_lst.Range(ABS_lst.Cells(1, 1), B)
'Set zag = ЗаголовокСтолбца(ABS_WB, ABS_lst.Name, "Позиция")
'If zag Is Nothing Then
'B.Offset(columnoffset:=1).Value = "Позиция"
'Set B = B.Offset(columnoffset:=1)
'Else
'Set B = zag
'End If
Set b = ЗаголовокСтолбцаСоздатьИлиВернуть(ABS_WB, ABS_lst.name, "Позиция")
Set m = ЗаголовокСтолбца(ABS_WB, ABS_lst.name, "Наименование")
NABC = m.Column




If ABS_lst.AutoFilterMode = True Then
   If ABS_lst.FilterMode = True Then
      iCountOfRows = ABS_lst.AutoFilter.Range.Columns(NABC).SpecialCells(xlVisible).Count
     
    Set f = ABS_lst.AutoFilter.Range.Columns(NABC).SpecialCells(xlVisible)
    ReDim Preserve v(f.Cells.Count - 2)
    k = 0
    For Each nr In f.Cells
     
If nr.Row <> 1 Then
Set v(k) = nr
k = k + 1
'если значение в столбце позиция пустое то устанавливается порядковый номер в выборке
If Cells(nr.Row, b.Column).value = "" Then Cells(nr.Row, b.Column).value = k
  
    End If
    Next
'      End If
   End If
End If
ArrayFiltrVisible = v
End Function



Sub ПеренестиОтфильтрованоеВНачалоСводной()
Dim wb As Workbook
Dim shNamenkl As Worksheet
Dim r As Range
Dim r_col As Range
Dim psname As String
Dim NameZag As String
Dim ColName As String
Dim Counter As Long             'считалка для прогрессбара
Dim TotalCells As Long          'прогрессбар
Dim pi As PivotItem
Dim pis As PivotItems
Dim PF As PivotField
Set wb = ActiveWorkbook
Set shNamenkl = wb.Worksheets("Номерклатура")
shNamenkl.Activate
v = ArrayFiltrVisible()
Call Intro
 TotalCells = UBound(v)
'             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             'Если количество ячеек большое, то скрываем форму и отображаем немодальный индикатор выполнения
             If TotalCells >= 10 Then

                 Application.ScreenUpdating = True
                 frmProgress.lPodpis.Caption = "Выполняется ОБНОВЛЕНИЕ СВОДНОЙ..."
                 frmProgress.LabelProgress.Width = 0
                 frmProgress.Show vbModeless
                 Application.ScreenUpdating = False
             End If
             Counter = 0
'             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'     pos = (1 / UBound(v)) * 100     'доля одной позиции в общем количестве позиций





For I = 0 To UBound(v)


Set r_col = Cells(1, v(I).Column)
ColName = r_col.value
psname = PivotSelectionName(ColName, v(I).value)
NameZag = "Позиция"
NamePoz = v(I).value
poz_row = v(I).Row
Set poz_col = ЗаголовокСтолбца(wb, shNamenkl.name, NameZag)

Poz = shNamenkl.Cells(poz_row, poz_col.Column)
Call ПереносВСводнойПоПозицииЗаданойЯчейкой(psname, ColName, Poz, NamePoz)


Counter = Counter + 1
                         If TotalCells >= 10 Then

                            If Counter Mod 100 <> 0 Then

                                 With frmProgress
                                     .FrameProgress.Caption = Format(Counter / TotalCells, "0%")
                                     .LabelProgress.Width = (Counter / TotalCells) * (.FrameProgress.Width - 10)
                                     .Repaint
                                 End With

                            End If

                         End If


Next I

frmProgress.Hide

Set shSvodnay = wb.Worksheets("Сводная")
Set PF = shSvodnay.PivotTables("CV_Ob").PivotFields("Наименование")
k = 1
Set pis = PF.PivotItems
 TotalCells = pis.Count - UBound(v)
  Counter = 0
'             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             'Если количество ячеек большое, то скрываем форму и отображаем немодальный индикатор выполнения
             If TotalCells >= 10 Then

                 Application.ScreenUpdating = True
                 frmProgress.lPodpis.Caption = "Выполняется ОБНОВЛЕНИЕ СВОДНОЙ..."
                 frmProgress.LabelProgress.Width = 0
                 frmProgress.Show vbModeless
                 Application.ScreenUpdating = False
             End If
             Counter = 0
For I = 1 + UBound(v) To pis.Count
Set pi = PF.PivotItems(I)
 If pi.Visible = True Then pi.Visible = False
 Counter = Counter + 1
                         If TotalCells >= 10 Then

                            If Counter Mod 100 <> 0 Then

                                 With frmProgress
                                     .FrameProgress.Caption = Format(Counter / TotalCells, "0%")
                                     .LabelProgress.Width = (Counter / TotalCells) * (.FrameProgress.Width - 10)
                                     .Repaint
                                 End With

                            End If

                         End If
       
  Next
Call Outro
End Sub
Sub jjjjj()
MsgBox ПоложениеСтолбца("ОбластьСтолбцов")
End Sub
Function ПоложениеСтолбца(ip)
en = Array(xlDataField, xlColumnField, xlPageField, xlRowField)

ru = Array("ОбластьДанных", "ОбластьСтолбцов", "ОбластьСтраниц", "ОбластьСтрок")
For I = 0 To 3
If ip = ru(I) Then ПоложениеСтолбца = en(I)
Next I
End Function
Sub СформироватьСводнуюПоЗначению(znach, Шаблон)
Dim wb As Workbook
Dim WSName As String
Dim ColumnName As String
Dim v As Variant
Dim r As Range
Set wb = ThisWorkbook
WSName = "ФорматСводных"
ColumnName = "Лист"
Set b = ЗаголовокСтолбца(wb, WSName, ColumnName)
Set KodColumn = ЗаголовокСтолбца(wb, WSName, "Код")
Set DelRowColumn = ЗаголовокСтолбца(wb, WSName, "Удалить строки")
Set PerenosSheetColumn = ЗаголовокСтолбца(wb, WSName, "Переносить в лист")
Worksheets("ФорматСводных").Activate

If Worksheets("ФорматСводных").AutoFilterMode = True And Worksheets("ФорматСводных").FilterMode = True Then
Worksheets("ФорматСводных").Rows(1).AutoFilter
Worksheets("ФорматСводных").Rows(1).AutoFilter
End If
Set r = Worksheets("ФорматСводных").Columns(b.Column)
Set r1 = r.Find(What:=znach, LookAt:=xlWhole, after:=Cells(1, b.Column))
Set sm = Worksheets("ФорматСводных").Cells(r1.Row, KodColumn.Column)
Set DelRow = Worksheets("ФорматСводных").Cells(r1.Row, DelRowColumn.Column)
Set PerenosSheet = Worksheets("ФорматСводных").Cells(r1.Row, PerenosSheetColumn.Column)
ResultSheetName = znach
ResultPivotTableName = znach & "_ALL"

Call New_Multi_Table_Pivot1(ResultSheetName, ResultPivotTableName, Шаблон)
Call CellAutoFilterVisible(sm.value)


If fСводная.cbPerenos Then
SheetName = PerenosSheet.value
RowsDelete = DelRow.value
Call СводнаяВ_Лист(Worksheets("ФорматСводных").Cells(r1.Row, 2).value, SheetName, RowsDelete)
'If fСводная.cbupr Then
'SheetName = PerenosSheet.Value
'RowsDelete = DelRow.Value
'Call СводнаяВ_Лист(Worksheets("ФорматСводных").Cells(r1.row, 2).Value, SheetName, RowsDelete)
'End If
End If



End Sub

Sub ПеренестиСводнуюПоЗначению(znach, Шаблон)
Dim wb As Workbook
Dim WSName As String
Dim ColumnName As String
Dim v As Variant
Dim r As Range
Set wb = ThisWorkbook
WSName = "ФорматСводных"
ColumnName = "Лист"
Set b = ЗаголовокСтолбца(wb, WSName, ColumnName)
Set KodColumn = ЗаголовокСтолбца(wb, WSName, "Код")
Set DelRowColumn = ЗаголовокСтолбца(wb, WSName, "Удалить строки")
Set PerenosSheetColumn = ЗаголовокСтолбца(wb, WSName, "Переносить в лист")
Worksheets("ФорматСводных").Activate

If Worksheets("ФорматСводных").AutoFilterMode = True And Worksheets("ФорматСводных").FilterMode = True Then
Worksheets("ФорматСводных").Rows(1).AutoFilter
Worksheets("ФорматСводных").Rows(1).AutoFilter
End If
Set r = Worksheets("ФорматСводных").Columns(b.Column)
Set r1 = r.Find(What:=znach, LookAt:=xlWhole, after:=Cells(1, b.Column))
Set sm = Worksheets("ФорматСводных").Cells(r1.Row, KodColumn.Column)
Set DelRow = Worksheets("ФорматСводных").Cells(r1.Row, DelRowColumn.Column)
Set PerenosSheet = Worksheets("ФорматСводных").Cells(r1.Row, PerenosSheetColumn.Column)
ResultSheetName = znach
ResultPivotTableName = znach & "_ALL"




SheetName = PerenosSheet.value
RowsDelete = DelRow.value
Call СводнаяВ_Лист(Worksheets("ФорматСводных").Cells(r1.Row, 2).value, SheetName, RowsDelete)




End Sub
