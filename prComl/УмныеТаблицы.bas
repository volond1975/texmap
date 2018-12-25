Attribute VB_Name = "УмныеТаблицы"

'УМНАЯ Таблица
    'myListObjectadd-добавляет и возвращает ссылку на умную таблицу
'СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
    'ListObjectColumnAdd-добавляет столбец с именем LCName в умную таблицу
    'ListObjectColumnRENAME-ПЕРЕИМЕНОВЫВАЕТ СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
'ФОРМУЛЫ В СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
'    ListObjectColumnFormulaLocal
'    ListObjectColumnFormulaR1C1
'    ListObjectColumnFormula-ПОЧЕМУТО РАБОТАЕТ НА СЛОЖНЫХ ФОРМУЛАХ ТОЛЬКО ЭТА
'    MAXListTableColumn()-возвращает максимальное значение столбца умной таблицы


'HeaderRowRangeStilNone-Удаляет заливку заголовка умной таблицы и цвет текста автомат




'----------------------------------------------------------------------------------------------------
'УМНАЯ Таблица
Function myListObjectadd(wb, NameSheet, NameListObject, r, Optional zag) As ListObject

Dim lo As ListObject
If IsMissing(zag) Then zag = 1
Set lo = wb.Worksheets(NameSheet).ListObjects.Add(xlSrcRange, r, , zag)
'ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$E$2"), , xlYes)
'Else
'Set lo = wb.Worksheets(NameSheet).ListObjects.Add(xlSrcRange, r, , xlNo)
'End If

lo.name = _
        NameListObject
', ListObjectTableStyle
'        LO.TableStyle = ListObjectTableStyle
        Set myListObjectadd = lo
End Function





Function myListObject(wb As Workbook, NameSheet, NameListObject)
Dim lo As ListObject
On Error Resume Next
Set lo = wb.Worksheets(NameSheet).ListObjects(NameListObject)
If VBA.Err = 0 Then
Set myListObject = lo
Else
Set myListObject = Nothing
End If

End Function
Function ListObjectExist(wb As Workbook, NameSheet, NameListObject)
Dim lo As ListObject
On Error Resume Next
Set lo = wb.Worksheets(NameSheet).ListObjects(NameListObject)
If VBA.Err = 0 Then ListObjectExist = True _
Else: ListObjectExist = False
End Function

'-----------------------------------------------------------------
'СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
'-------------------------------------------------------------------
Function ListObjectColumnAdd(lo As ListObject, LCName)
'добавляет столбец с именем LCName в умную таблицу
Dim lcs As ListColumn
Set lcs = lo.ListColumns.Add
     lo.HeaderRowRange(lcs.Index) = LCName
Set ListObjectColumnAdd = lcs
End Function
Function ListObjectColumnCount(ИмяЛиста, ИмяТаблицы)
'добавляет столбец с именем LCName в умную таблицу
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)


   ListObjectColumnCount = ls.HeaderRowRange.Columns.Count
End Function



Function ListObjectColumnAddFormulaOrValue(lo As ListObject, LCName As String, formul As String, form_at As String, v As Boolean) As ListColumn
'добавляет столбец с именем LCName в умную таблицу,вставляет формулу Formul,форматирует form_at,и при v As true превращает в значения
Dim lcs As ListColumn
Set lcs = ListObjectColumnAdd(lo, LCName)
    Call ListObjectColumnFormula(lcs, formul, form_at)
  If v Then
  For Each r In lcs.Range.Cells
  r.value = r.value
  Next
End If
Set ListObjectColumnAddFormulaOrValue = lcs
End Function



Function ListObjectColumnRENAME(lo As ListObject, LCName, NewName)
'ПЕРЕИМЕНОВЫВАЕТ СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
Dim lcs As ListColumn
Set lcs = lo.ListColumns(LCName)
     lo.HeaderRowRange(lcs.Index) = NewName
Set ListObjectColumnRENAME = lcs
End Function

'--------------------------------------------------------------------------------------

'ФОРМУЛЫ В СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
'---------------------------------------------------------------------------------------
Sub ListObjectColumnFormulaLocal(lcs As ListColumn, myFormulaLocal As String, Optional NumbFormat)
Dim lo As ListObject
Dim r As Range
Set lo = lcs.Parent
Set r = Intersect(lo.DataBodyRange, lcs.Range)
       r.FormulaLocal = myFormulaLocal
       r.NumberFormat = NumbFormat
       
End Sub
Sub ListObjectColumnFormulaR1C1(lcs As ListColumn, myFormulaR1C1, Optional NumbFormat)
Dim lo As ListObject
Dim r As Range
Set lo = lcs.Parent
Set r = Intersect(lo.DataBodyRange, lcs.Range)
       r.FormulaR1C1 = myFormulaR1C1
       r.NumberFormat = NumbFormat
       
End Sub

Sub ListObjectColumnFormula(lcs As ListColumn, myFormula, Optional NumbFormat)
Dim lo As ListObject
Dim r As Range
Set lo = lcs.Parent
Set r = Intersect(lo.DataBodyRange, lcs.Range)
       r.Formula = myFormula
       r.NumberFormat = NumbFormat
       
End Sub

Function ListObjectColumnExcelFormula(lcs As ListColumn)
ListObjectColumnFormula "=" & lsc.Parent.name & "[" & lsc.name & "]"
End Function

'---------------------------------------------------------------------------------------------------

Function MAXListTableColumn(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
'возвращает максимальное значение столбца умной таблицы

MAXListTableColumn = Application.WorksheetFunction.max(СписокЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца))
End Function

Function СписокЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбца)
Set СписокЗначенийСтолбцаУмнойТаблицы = Range(lsc.Range.Cells(2), lsc.Range.Cells(lsc.Range.Cells.Count))


End Function
Function СписокЗначенийСтолбцаУмнойТаблицыСЗаголовком(ИмяЛиста, ИмяТаблицы, ИмяСтолбца) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбца)
Set r = lsc.DataBodyRange
Set СписокЗначенийСтолбцаУмнойТаблицыСЗаголовком = r


End Function














Function ЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска) As Range
'ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
'"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
' Можно использовать индексы
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбцаПоиска)
Set lscf = ls.ListColumns(ИмяСтолбцаЗначения)
Set n = lsc.Range.Find(ЗначениеПоиска, LookIn:=xlValues, LookAt:=xlWhole)
If Not n Is Nothing Then
Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set ЗначениеУмнойТаблицы = r
Else
Set ЗначениеУмнойТаблицы = Nothing
End If
End Function
Function ЗначениеУмнойТаблицыV(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)
'ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
'"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
' Можно использовать индексы
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбцаПоиска)
v = ИмяСтолбцаЗначения
Set lscf = ls.ListColumns(v)
Set n = lsc.Range.Find(ЗначениеПоиска, LookIn:=xlValues, LookAt:=xlWhole)
If Not n Is Nothing Then
Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
ЗначениеУмнойТаблицыV = r.value
Else
ЗначениеУмнойТаблицыV = ""
End If
End Function














'Sub ЗначениеУмнойТаблицыИSubstring(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска, Delimiter, n)
'Dim txt
'Set txt = ЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)
'ЗначениеУмнойТаблицыИSubstring = Substring(txt, Delimiter, n)
''Номер договора подряда из запроса
'ф = ЗначениеУмнойТаблицыИSubstring("Приймання", Приймання, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска, Delimiter, n)
'
'
'm = Split(r.Offset(columnoffset:=1).value, ",")
'Glav.ComboBox_№ДП.Text = VBA.Trim(m(1))
''Виконавец договора подряда из запроса
'Glav.cbB.Text = m(2)
'm = Split(r.Offset(columnoffset:=1).value, " ")
''Дата договора подряда из запроса
'Glav.TextBox_ДатаДП.Text = m(3)
''Майстер  из запроса
'Glav.ComboBox_Мастер.Text = r.Offset(columnoffset:=2).value
'End Sub
















Function НовоеЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска, Добавить As Boolean) As Range
'ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
'"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
' Можно использовать индексы
'1 Добавить автоматично
'2 msgbox

Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбцаПоиска)
Set lscf = ls.ListColumns(ИмяСтолбцаЗначения)
Set n = lsc.Range.Find(ЗначениеПоиска, LookAt:=xlWhole)

If Not n Is Nothing Then
Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set НовоеЗначениеУмнойТаблицы = r
Else
    If Добавить Then
    Set n = ДобавитьСтрокуИКод(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ЗначениеПоиска)
    Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set НовоеЗначениеУмнойТаблицы = r
    Else
    Select Case MsgBox("Значение не найдено ! Внести в справочник", vbOKCancel Or vbCritical Or vbDefaultButton1, "Значение не найдено")
    
        Case vbOK
     Set n = ДобавитьСтрокуИКод(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ЗначениеПоиска)
    Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set НовоеЗначениеУмнойТаблицы = r
        Case vbCancel
    Set НовоеЗначениеУмнойТаблицы = Nothing
    End Select
    End If

End If
End Function
Function НовоеЗначениеСтрокиСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, lsr As ListRow, ИмяСтолбца, Значение) As Range
'ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
'"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
' Можно использовать индексы
'1 Добавить автоматично
'2 msgbox

Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
'Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбца)

'Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lsc.Range, lsr.Range)
r = Значение
Set НовоеЗначениеСтрокиСтолбцаУмнойТаблицы = r

    
 
End Function



Function ДобавитьСтрокуУмнойТаблицы(ИмяЛиста, ИмяТаблицы, Optional Положение, Optional AlwaysInserta As Boolean = True) As ListRow
'Добавляет новую строку в таблице, представленной указанным ListObject
'Положение-Определяет относительное положение новой строки
' Указывает, следует ли всегда переложить данные в ячейки ниже последней строки таблицы,
'когда вставляется новая строка, независимо от того, строка ниже таблице пуст.
'Если True , что ниже таблице клетки будут сдвинуты на одну строку вниз.
'Если False , если строка ниже таблице пуст, таблица будет расширяться, чтобы занять эту строку
'без смещения клеток под ним; но если строка ниже таблице содержатся данные,
'эти клетки будут сдвинуты вниз, когда вставляется новая строка.


Dim wb As Workbook
Dim ws As Worksheet
Dim lo As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set lo = ws.ListObjects(ИмяТаблицы)
If IsMissing(Положение) Then
Set lsr = lo.ListRows.Add
Else
Set lsr = lo.ListRows.Add(Положение, AlwaysInsert)
End If
Set ДобавитьСтрокуУмнойТаблицы = lsr
End Function











Function СтрокаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ЗначениеПоиска) As ListRow
'ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
'"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
' Можно использовать индексы
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбцаПоиска)
'Set lscf = ls.ListColumns(ИмяСтолбцаЗначения)
Set n = lsc.Range.Find(ЗначениеПоиска, LookAt:=xlWhole)
Set lsr = ls.ListRows(n.Row - 1)
'Set r = Application.Intersect(lscf.Range, lsr.Range)
Set СтрокаУмнойТаблицы = lsr
End Function

Sub СтрокуУмнойТаблицыУдалить(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ЗначениеПоиска)
Dim lsr As ListRow
Set lsr = СтрокаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ЗначениеПоиска)
'lsr.ClearContents
lsr.Delete
End Sub
Function СтрокаУмнойТаблицыПоИндексу(ИмяЛиста, ИмяТаблицы, Индекс) As ListRow
'ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
'"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
' Можно использовать индексы
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)


Set lsr = ls.ListRows(Индекс)

Set СтрокаУмнойТаблицыПоИндексу = lsr
End Function




Function ВидимыйДиапазонЗначенийУмнойТаблицыБезЗаголовка(ИмяЛиста, ИмяТаблицы) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
On Error Resume Next
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)

Set r = ls.DataBodyRange
Set ВидимыйДиапазонЗначенийУмнойТаблицыБезЗаголовка = r.SpecialCells(xlCellTypeVisible)


End Function
Function ВидимыйДиапазонЗначенийУмнойТаблицы(ИмяЛиста, ИмяТаблицы) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)

Set r = ls.Range
Set ВидимыйДиапазонЗначенийУмнойТаблицы = r.SpecialCells(xlCellTypeVisible)


End Function
Function ВидимыйДиапазонЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
On Error Resume Next
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбца)
Dim rr As Range
Set rr = ls.DataBodyRange
Set www = Intersect(lsc.Range, rr)
Set r = www.SpecialCells(xlCellTypeVisible)
'Set r = lsc.Range.SpecialCells(xlCellTypeVisible)
Set ВидимыйДиапазонЗначенийСтолбцаУмнойТаблицы = r


End Function
Sub СортировкаСтолбцаУмнойТаблици(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
'
' Макрос29 Макрос
'

'
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбца)
   ls.Sort.SortFields _
        .Clear
    ls.Sort.SortFields _
        .Add Key:=lsc.Range, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers
    With ls.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub







Sub ДобавитьБригадира(br)
Dim Findrange As Range
Dim f As Range
Set Findrange = СписокЗначенийСтолбцаУмнойТаблицы("Бригадир", "Бригадир", "Бригадир")
Set f = FindAll(SearchRange:=Findrange.Cells, FindWhat:=br, LookAt:=xlPart)

If f Is Nothing Then

Set r = ДобавитьСтрокуИКод("Бригадир", "Бригадир", "Бригадир", br)
End If
End Sub
Sub cvfd()
ИмяЛиста = "Список"
ИмяТаблицы = "Список"
ist = Array("Лісництво_Ділянка", "Бригадир", "Обхід", "Склад бріг")
br = Array("Хлипнівське 15 кв (4 вид) 0 діл.", "Тест4", "4", "6")
Set rr = ДобавитьMasiv(ИмяЛиста, ИмяТаблицы, ist, br)

End Sub

Function ДобавитьMasiv(ИмяЛиста, ИмяТаблицы, ist As Variant, br As Variant)
Dim Findrange As Range
Dim lsr As ListRow
Dim f As Range

Set Findrange = СписокЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ist(0))


Set f = FindAll(SearchRange:=Findrange.Cells, FindWhat:=br(0), LookAt:=xlPart)

If f Is Nothing Then

Set rr = ДобавитьСтрокуИКод(ИмяЛиста, ИмяТаблицы, ist(0), br(0))
For I = 0 To UBound(ist)
q = IndexColumn(ИмяЛиста, ИмяТаблицы, ist(I))
rr.Cells(q).value = br(I)

Next I

Else
Set lsr = СтрокаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ist(0), br(0))
Set rr = lsr
For I = 1 To UBound(ist)
q = IndexColumn(ИмяЛиста, ИмяТаблицы, ist(I))
rr.Range.Cells(q).value = br(I)

Next I
End If
'For i = 1 To UBound(ist)
'q = IndexColumn(ИмяЛиста, ИмяТаблицы, ist(i))
'rr.Range.Cells(q).Value = br(i)
'
'Next i
Set ДобавитьMasiv = rr
End Function

 Function IndexColumn(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
 Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set zagrow = ls.HeaderRowRange
IndexColumn = Application.WorksheetFunction.Match(ИмяСтолбца, zagrow, 0)


 End Function
 Sub DeleteColumn(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
 On Error Resume Next
 Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set zagrow = ls.HeaderRowRange
mIndexColumn = Application.WorksheetFunction.Match(ИмяСтолбца, zagrow, 0)
ls.ListColumns(mIndexColumn).Delete

 End Sub
'---------------------------------------------------------------------------------------
' Procedure : ФильтрСтолбцаУмнойТаблицы
' Author    : ВАЛЕРА
' Date      : 26.11.2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
 Sub ФильтрСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца, Optional myCriteria As String)
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Dim z As Integer
   On Error GoTo ФильтрСтолбцаУмнойТаблицы_Error

Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
z = IndexColumn(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
   If Not myCriteria = "" Then
ls.Range.AutoFilter Field:=z, Criteria1:= _
        myCriteria
        Else
    ls.Range.AutoFilter Field:=z
    End If

   On Error GoTo 0
   Exit Sub

ФильтрСтолбцаУмнойТаблицы_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ФильтрСтолбцаУмнойТаблицы of Module УмныеТаблицы"
End Sub
Function ДобавитьСтрокуИКод(ИмяЛиста, ИмяТаблицы, ИмяСтолбца, Код) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set zagrow = ls.HeaderRowRange
q = Application.WorksheetFunction.Match(ИмяСтолбца, zagrow.Cells, 0)
Set lsr = ls.ListRows.Add(AlwaysInsert:=True)
'Set lsc = ls.ListColumns(ИмяСтолбца)
Set rr = ls.Range.Rows(ls.Range.Rows.Count)
'Set cc = lsc.DataBodyRange
Set r = lsr.Range.Cells(q)
r.value = Код
Set ДобавитьСтрокуИКод = rr


End Function
Function ВыделитьСтрокуПоКоду(ИмяЛиста, ИмяТаблицы, ParamArray ЗначениеПоиска()) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set lsc = ls.ListColumns(ИмяСтолбца)
Set zagrow = ls.HeaderRowRange
q = Application.WorksheetFunction.Match(ИмяСтолбца, zagrow.Cells, 0)
Set n = lsc.Range.Find(ЗначениеПоиска, LookAt:=xlWhole)
Set lsr = ls.ListRows.Add(AlwaysInsert:=True)
'Set lsc = ls.ListColumns(ИмяСтолбца)
Set rr = ls.Range.Rows(ls.Range.Rows.Count)
'Set cc = lsc.DataBodyRange
Set r = lsr.Range.Cells(q)
r.value = Код
Set ВыделитьСтрокуПоКоду = rr


End Function


Sub CopyFilteredRowsAndHeadingsList1()
Dim wsL As Worksheet
Dim ws As Worksheet
Dim Rng As Range
Dim rng2 As Range
Dim Lst As ListObject

Application.ScreenUpdating = False
Set wsL = ActiveSheet
Set Lst = wsL.ListObjects(1)

With Lst.AutoFilter.Range
 On Error Resume Next
   Set rng2 = .Offset(1, 0).Resize(.Rows.Count - 1, 1) _
       .SpecialCells(xlCellTypeVisible)
 On Error GoTo 0
End With
If rng2 Is Nothing Then
   MsgBox "No data to copy"
Else
   Set ws = Sheets.Add
   Set Rng = Lst.AutoFilter.Range
   'copy rows with headings
   Rng.SpecialCells(xlCellTypeVisible).Copy _
     destination:=ws.Range("A1")
End If
   
Application.ScreenUpdating = True

End Sub
Function КоличествоСтрок(ИмяЛиста, ИмяТаблицы)
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)

КоличествоСтрок = ls.DataBodyRange.Rows.Count


End Function
Function СсылкаНаУмнуюТаблицу(имякниги As Workbook, ИмяЛиста As String, ИмяТаблицы) As ListObject
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range

If Not IsBookOpen(имякниги.name) Then Exit Function
If Not SheetExistBook(имякниги, ИмяЛиста) Then Exit Function
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
Set СсылкаНаУмнуюТаблицу = ls
End Function
Function Консолидация()
Dim ls1 As ListObject
Dim ls2 As ListObject


Set ls1 = СсылкаНаУмнуюТаблицу(ИмяКниги1, ИмяЛиста1, ИмяТаблицы1)
Set ls2 = СсылкаНаУмнуюТаблицу(ИмяКниги2, ИмяЛиста2, ИмяТаблицы2)






End Function



Sub ФильтрСолбца(ИмяЛиста, ИмяТаблицы, ИмяСтолбца, Criteria, Optional Убрать As Boolean)
'
' Макрос18 Макрос
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
Set ls = ws.ListObjects(ИмяТаблицы)
If Убрать Then
ws.ListObjects(ИмяТаблицы).Range.AutoFilter Field:=IndexColumn(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
Else
    ws.ListObjects(ИмяТаблицы).Range.AutoFilter Field:=IndexColumn(ИмяЛиста, ИмяТаблицы, ИмяСтолбца), Criteria1:= _
        Criteria
        End If
End Sub
 
Function MyTableRangeID(MyTable As ListObject, r As Range)
   Dim isect As Range
    
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
            MyTableRangeID = 1 ' "ячейка принадлежит таблице " & MyTable.Name
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
               MyTableRangeID = 2 'ячейка в заголовке"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.DataBodyRange)
            If Not (isect Is Nothing) Then
               MyTableRangeID = 3 'ячейка в области данных"
            End If
            Set isect = Nothing
            On Error Resume Next
            If MyTable.ShowTotals Then
                If Not Intersect(r, MyTable.TotalsRowRange) Is Nothing Then
                   MyTableRangeID = 5 '"Ячейка в строке итогов" & vbNewLine
                Else
                    Temp = Temp & "Строка итогов есть, но ячейка не в ней" & vbNewLine
                End If
 
            Else
                If Not Intersect(r, MyTable.Range.Rows(lo.Range.Rows.Count)) Is Nothing Then
                   MyTableRangeID = 4 ' "Строки итогов нет, ячейка в последней строке таблицы" & vbNewLine
                End If
            End If
            On Error GoTo 0
            Else
           MyTableRangeID = 0
        End If
   
End Function
Sub HeaderRowRangeStilNone(lo As ListObject)
'
' Удаляет заливку заголовка умной таблицы и цвет текста автомат
'

'
   
    With lo.HeaderRowRange
    .Interior.pattern = xlNone
    .Font.ColorIndex = xlAutomatic
      
    End With
End Sub
Function fgaAuthorList(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
count_tab = КоличествоСтрок(ИмяЛиста, ИмяТаблицы)
 Set v = ВидимыйДиапазонЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
 
lngLoopCount = 0
ReDim mgaAuthorList(count_tab - 1)
For I = 1 To count_tab
mgaAuthorList(lngLoopCount) = v(I)
lngLoopCount = lngLoopCount + 1

Next I
fgaAuthorList = mgaAuthorList
End Function
Function fgaAuthorListUnikum(ИмяЛиста, ИмяТаблицы, ИмяСтолбца, Optional Unikum)
count_tab = КоличествоСтрок(ИмяЛиста, ИмяТаблицы)
 Set v = ВидимыйДиапазонЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
 If Unikum Then Set v = mMacros.UnicumRange(v)
lngLoopCount = 0
ReDim mgaAuthorList(count_tab - 1)
For I = 1 To count_tab
mgaAuthorList(lngLoopCount) = v(I)
lngLoopCount = lngLoopCount + 1

Next I
fgaAuthorListUnikum = mgaAuthorList
End Function
Function fgaAuthorListVisible(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
 Set v = ВидимыйДиапазонЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
 count_tab = v.Cells.Count

lngLoopCount = 0
ReDim mgaAuthorList(count_tab - 1)
For I = 1 To count_tab
mgaAuthorList(lngLoopCount) = v(I)
lngLoopCount = lngLoopCount + 1

Next I
fgaAuthorListVisible = mgaAuthorList
End Function
Function ListVisible(ИмяЛиста, ИмяТаблицы)
 Set v = ВидимыйДиапазонЗначенийУмнойТаблицыБезЗаголовка(ИмяЛиста, ИмяТаблицы)
 If v Is Nothing Then Exit Function
 
 count_tab = ListObjectColumnCount(ИмяЛиста, ИмяТаблицы)
 count_row = v.Cells.Count / count_tab

lngLoopCount = 1
ReDim mgaAuthorList(1 To count_row, 1 To count_tab)
For I = 1 To count_row
For j = 1 To count_tab
mgaAuthorList(I, j) = v(lngLoopCount)
lngLoopCount = lngLoopCount + 1
Next j
Next I
ListVisible = mgaAuthorList

End Function



'Sub ListControlListobjectColumn(forma As UserForm, ИмяЛиста, ИмяТаблицы, ИмяСтолбца, ИмяКонтрола, Уникальность As Boolean)
'Dim contr
'
'Set contr = forma.Controls(ИмяКонтрола)
'
'If Уникальность Then
'v = fgaAuthorListUnikum(ИмяЛиста, ИмяТаблицы, ИмяСтолбца, Уникальность)
'Else
'v = fgaAuthorListVisible(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
'End If
'contr.List = v
'
'End Sub
Sub Test_Cell()
'Добавляет в активную умную таблицу(если активная ячейка находится в межах умной таблици)
'в конце новую пустую строку
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
            Set lsr = ДобавитьСтрокуУмнойТаблицы(ActiveSheet.name, MyTable.name)
            
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в заголовке"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в итогах"
            End If
            On Error GoTo 0
        End If
    Next
End Sub
Sub Test_Cell_TypE(Optional v)
'Добавляет в активную умную таблицу(если активная ячейка находится в межах умной таблици)
'в конце новую пустую строку и в первую ячейку вставляет
'значение строки в которой находилась активная ячейка
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        Set isect = Application.Intersect(ActiveSheet.Rows(ActiveCell.Row), MyTable.Range)
            Set lsr = ДобавитьСтрокуУмнойТаблицы(ActiveSheet.name, MyTable.name)
       If IsMissing(v) Then v = Array(1)
       For I = 0 To UBound(v)
      lsr.Range.Cells(v(I)) = isect.Cells(v(I))
      Next I
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в заголовке"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в итогах"
            End If
            On Error GoTo 0
        End If
    Next
End Sub
Sub DEl_ROw_TypE()
'удаляет из  активной умной таблицы(если активная ячейка находится в межах умной таблици)
'
'строку в которой находилась активная ячейка
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
      
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в заголовке"
                Exit Sub
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в итогах"
                Exit Sub
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveSheet.Rows(ActiveCell.Row), MyTable.Range)
            isect.Delete
            On Error GoTo 0
        End If
    Next
End Sub

Sub AddShowTotal()
'Добавляет( или убирает если уже есть)строку итогов в активную умную таблицу(если активная ячейка находится в межах умной таблици)
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
       If MyTable.ShowTotals = True Then MyTable.ShowTotals = False Else MyTable.ShowTotals = True
            
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в заголовке"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в итогах"
                
                
            End If
            On Error GoTo 0
        End If
    Next
End Sub


Sub AddShowAutoFilterDropDown()
'Добавляет( или убирает если уже есть)ярлыки фильтров в активную умную таблицу(если активная ячейка находится в межах умной таблици)
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
       If MyTable.ShowAutoFilterDropDown = True Then MyTable.ShowAutoFilterDropDown = False Else MyTable.ShowAutoFilterDropDown = True
            
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в заголовке"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell в итогах"
                
                
            End If
            On Error GoTo 0
        End If
    Next
End Sub
Sub RemoveTableBodyData()
'удаляет из  активной умной таблицы(если активная ячейка находится в межах умной таблици)
'все значения
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then


'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
    MyTable.DataBodyRange.Delete
    
    
  End If
End If
    Next
End Sub
Sub SortСolumnAscending()
' Сортировка по возростанию в стобце активной умной таблицы по положению активной ячейки
Call SortСolumn(1)
End Sub
Sub SortСolumnDescending()
' Сортировка по убыванию в стобце активной умной таблицы по положению активной ячейки
Call SortСolumn(2)
End Sub

Sub SortСolumn(myOrder)
'
' Сортировка  в стобце активной умной таблицы по положению активной ячейки
'1-xlAscending-по возростнию
'2-xlDescending-по убыванию
 Dim lo As ListObject
 Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim ActiveColumnListObject As Range
Dim ActiveColumnListObjectName As Range
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
'        Активный столбец активной умной таблицы
 Set ActiveColumnListObject = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
' Назва активного столбца активной умной таблицы
 Set ActiveColumnListObjectName = ActiveColumnListObject.Cells(1)
 KeyrangeStruktureAddress = MyTable.name & "[[#Headers],[#Data],[" & ActiveColumnListObjectName.value & "]]"
    MyTable.Sort. _
        SortFields.Clear
'  Range("Загальна[[#Headers],[#Data],[Тип]]")
    MyTable.Sort. _
        SortFields.Add Key:=Range(KeyrangeStruktureAddress), SortOn:= _
        xlSortOnValues, Order:=myOrder, DataOption:=xlSortTextAsNumbers
      v = MyTable.Sort.SortFields(1).Order
    With MyTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End If
    Next
End Sub



Sub SortFieldsClear()
'
' Сортировка по возростанию в стобце активной умной таблицы по положению активной ячейки
'

 Dim lo As ListObject
 Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim ActiveColumnListObject As Range
Dim ActiveColumnListObjectName As Range
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
'        Активный столбец активной умной таблицы
 Set ActiveColumnListObject = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
' Назва активного столбца активной умной таблицы
 Set ActiveColumnListObjectName = ActiveColumnListObject.Cells(1)
 KeyrangeStruktureAddress = MyTable.name & "[[#Headers],[#Data],[" & ActiveColumnListObjectName.value & "]]"
    MyTable.Sort. _
        SortFields.Clear
''  Range("Загальна[[#Headers],[#Data],[Тип]]")
'    MyTable.Sort. _
'        SortFields.Add Key:=Range(KeyrangeStruktureAddress), SortOn:= _
'        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'    With MyTable.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With


'ActiveWorkbook.Worksheets("Розрахунок2").ListObjects("Загальна").Sort. _
'        SortFields.Add Key:=Range("Загальна[[#Headers],[#Data],[Назва]]"), SortOn:= _
'        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortTextAsNumbers
'    With ActiveWorkbook.Worksheets("Розрахунок2").ListObjects("Загальна").Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With







    End If
    Next
End Sub







Sub ФильтрПоАктивнойЯчейке()
'удаляет из  активной умной таблицы(если активная ячейка находится в межах умной таблици)
'все значения
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then

Set isect = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
    Call ФильтрСтолбцаУмнойТаблицы(ActiveSheet.name, MyTable.name, isect.Cells(1), ActiveCell.value)
 
  End If
End If
    Next
End Sub

Sub УдаляетФильтрПоАктивнойЯчейке()
'удаляет из  активной умной таблицы(если активная ячейка находится в межах умной таблици)
'все значения
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then

Set isect = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
    Call ФильтрСтолбцаУмнойТаблицы(ActiveSheet.name, MyTable.name, isect.Cells(1))
 
  End If
End If
    Next
End Sub
Sub СправочникПоАктивнойЯчейке(Optional wb, Optional ShName, Optional loName, Optional r, Optional v)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(ShName) Then ShName = ActiveCell.value
If IsMissing(loName) Then loName = ActiveCell.value
If IsMissing(wb) Then Set wb = ActiveWorkbook
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
If ActiveCell.value = "" Then
Call MsgBox("Введите назву справочника в активную ячейку", vbCritical, "Создание справочника")
Exit Sub

End If
Set isect = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
  
 Set sh = SheetExistBookCreate(wb, ShName, True)
 If IsMissing(r) Then Set r = sh.Cells(1, 1)
 
   Set New_tbl = myListObjectadd(wb, ShName, loName, r)
    If IsMissing(v) Then v = Array("Назва", "Вага")
    For I = 0 To UBound(v)
    If I = 0 Then
 Call ListObjectColumnRENAME(New_tbl, 1, v(I))
 Else
 Call ListObjectColumnAdd(New_tbl, v(I))
 End If
 Next I
  End If
End If
    Next
End Sub
Sub СправочникПоДанным(Optional wb, Optional ShName, Optional loName, Optional r, Optional v)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(ShName) Then ShName = ActiveCell.value
If IsMissing(loName) Then loName = ActiveCell.value
If IsMissing(wb) Then Set wb = ActiveWorkbook

If loName = "" Then
Call MsgBox("Введите назву справочника в активную ячейку", vbCritical, "Создание справочника")
Exit Sub

End If

  
  
 Set sh = SheetExistBookCreate(wb, ShName, True)
 If IsMissing(r) Then Set r = sh.Cells(1, 1)
 
   Set New_tbl = myListObjectadd(wb, ShName, loName, r)
    If IsMissing(v) Then v = Array("Назва", "Вага")
    For I = 0 To UBound(v)
    If I = 0 Then
 Call ListObjectColumnRENAME(New_tbl, 1, v(I))
 Else
 Call ListObjectColumnAdd(New_tbl, v(I))
 End If
 Next I
 

End Sub


Sub AutoFillFormulasInListsRevers()
'Еще один вопрос по таблицам listobject.
'Если туда записать формулу, например "=890", она заполнится по всему столбцу. Если после нажать отмену, заполнение снимается.
'
'А как отменить автозаполнение в таблицах, если я вставляю формулу через vba? Можно ли отключить его временно до начала функции, затем снова включить?


'автозаполнение ListObject
If Application.AutoCorrect.AutoFillFormulasInLists Then

Application.AutoCorrect.AutoFillFormulasInLists = False ' Выключить автозаполнение ListObject
Else
    Application.AutoCorrect.AutoFillFormulasInLists = True  ' Включить автозаполнение ListObject
    End If
End Sub


Sub ShowAllRecordsList()
'Показати всі записи

'Наступний макрос показує всі записи у активной умной таблице  на активному аркуші, якщо фільтр був застосований.
Dim Lst As ListObject
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then


 If MyTable.AutoFilter.FilterMode Then
    MyTable.AutoFilter.ShowAllData
  End If
End If
    Next


End Sub

Sub HideArrowsList1()
'hides all arrows except list 1 column 2
'Приховати все Список автофильтр стрілки, крім одного

'Може бути, ви хочете, щоб користувачі фільтр тільки один із стовпців в список 1. Наступна процедура Excel Автофільтр VBA приховує стрілки для всіх стовпців, крім другій колонці в список 1
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
Dim r As Range
Application.ScreenUpdating = False
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then

Set ActiveColumnListObject = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
I = 1

For Each C In MyTable.HeaderRowRange
 If I <> ActiveColumnListObject.Column Then
    MyTable.Range.AutoFilter Field:=I, _
      VisibleDropDown:=False
 Else
     MyTable.Range.AutoFilter Field:=I, _
      VisibleDropDown:=True
 End If
 I = I + 1
Next
End If
    Next

Application.ScreenUpdating = True



End Sub
Sub AddParametr(ЗначениеПоиска, Значение)
ИмяЛиста = "Настройки"
ИмяТаблицы = "Parameters"
ИмяСтолбцаПоиска = "Parameter"
ИмяСтолбцаЗначения = "Value"
Set r = ЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)
r.value = Значение
End Sub
 
Function Parametr(ЗначениеПоиска)
ИмяЛиста = "Настройки"
ИмяТаблицы = "Parameters"
ИмяСтолбцаПоиска = "Parameter"
ИмяСтолбцаЗначения = "Value"
Set r = ЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)
Set Parametr = r
End Function
 
Sub NullValueListObjects(ИмяЛиста, ИмяТаблицы, ИмяСтолбца, v, ЗначениеПоиска)
Dim r As Range
For I = 0 To UBound(v)
Set r = ЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца, v(I), ЗначениеПоиска)
If Not (r Is Nothing) Then r.value = ""
Next I
