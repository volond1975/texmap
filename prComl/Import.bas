Attribute VB_Name = "Import"
Sub ИмпортПоСтроке(Имя_Листа)
Dim EOBookName As String
Dim EOBookSheetName As String
Dim EOBookSheetNewName As String
Dim twb As Workbook
Dim r As Range
Set twb = ThisWorkbook

EOBookName = ЗначениеУмнойТаблицы("Импорт", "Импорт", "Новое Имя Листа", "Файл для импорта", Имя_Листа)
EOBookSheetName = ЗначениеУмнойТаблицы("Импорт", "Импорт", "Новое Имя Листа", "Лист", Имя_Листа)
EOBookSheetNewName = Имя_Листа
Call Импорт(EOBookName, EOBookSheetName, EOBookSheetNewName)


End Sub

Sub ИмпортРеестраЛк()
Dim EOBookName As String
Dim EOBookSheetName As String
Dim EOBookSheetNewName As String
Dim twb As Workbook
Set twb = ThisWorkbook
EOBookName = "Реєстр лісорубних квитків.xls"
EOBookSheetName = "TDSheet"
EOBookSheetNewName = "Реестр_ЛК"
Call Импорт(EOBookName, EOBookSheetName, EOBookSheetNewName)


End Sub
Sub ИмпортЩоденника(EOBookName As String)

Dim EOBookSheetName As String
Dim EOBookSheetNewName As String
Dim twb As Workbook
Call Intro
Set twb = ThisWorkbook

EOBookSheetName = "TDSheet"
EOBookSheetNewName = "Зведений щоденник"
ВартистьЗнебособленного = 100
ВартистьХлист = 50
Call Импорт(EOBookName, EOBookSheetName, EOBookSheetNewName)
Set Дата_приймання = Cells.Find(What:="Дата приймання", LookAt:=xlWhole)
Set Примітка = Cells.Find(What:="Примітка", LookAt:=xlWhole)
Set ЛК_№ = Cells.Find(What:="ЛК №", LookAt:=xlPart)
Set Всього = Cells.Find(What:="Всього", LookAt:=xlPart)
Set Дані = Range(Дата_приймання, Cells(Всього.Offset(rowoffset:=-1).Row, Примітка.Column))
ActiveSheet.ListObjects.Add(xlSrcRange, Range(Дані.address), , xlYes).name = _
        "Щоденник"
    
    ActiveSheet.ListObjects("Щоденник").TableStyle = "TableStyleMedium2"
   Call СводнаяЩоденник
 EOBookSheetNewName = "Таблица Щоденник"
  Call СводнаяВ_Лист("ЗведенаЩоденник", EOBookSheetNewName, 3)


Set sh = twb.Worksheets(EOBookSheetNewName)
'Добавить отсутствующие заголовки
sh.Cells(1, 8).value = "Вартисть робіт, грн./м3"
sh.Cells(1, 9).value = "Сума, грн."
Call ФорматЗаголовков(Range(sh.Cells(1, 1), sh.Cells(1, 9)))
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





Call РазбивкаТехсировиниПоДлинам(ВартистьЗнебособленного, ВартистьХлист)
 Call Outro
End Sub

Sub ИмпортЗапросADO()
    Dim ADO As New ADO
Dim oCN   As Object    'Connection
Dim oRS   As Object
Dim fl As Object
Dim wb As Workbook   'Recordset
Dim lo As ListObject
Dim sQuery As String  ' Создаем экземпляр класса
ListObjectName = "Виконавец"
Set wb = Workbooks("Комплект.xlsm")
 ShetNameActivate = ActiveSheet.name

Set sha = wb.Worksheets(ShetNameActivate)
Set sh = SheetExistBookCreate(wb, "Лист3", True)

   ADO.DataSource = ThisWorkbook.Path & "\" & "Техкарта 2016.xls"
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
'sh.Name = "Запрос"
'    ADO.Query ("SELECT F2 FROM [Лист1$];")
'    Range("F1").CopyFromRecordset ADO.Recordset
'
'    ' Закрываем соединение, чтобы не висело : )
'    ADO.Disconnect
'
'    ADO.Query ("SELECT F1 FROM [Лист1$] UNION SELECT F2 FROM [Лист1$];")
'    Range("G1").CopyFromRecordset ADO.Recordset
        
    ' Тут автоматически закроется соединение
    ' и уничтожиться объекты Recordset и Connection
End Sub





























Sub РазбивкаТехсировиниПоДлинам(ВартистьЗнебособленного, ВартистьХлист)
Dim twb As Workbook
Dim EOBookSheetNewName As String
Dim ИсхСтрока As Range
Dim v As Variant
Dim sh As Worksheet
Set twb = ThisWorkbook
EOBookSheetNewName = "Таблица Щоденник"
Set sh = twb.Worksheets(EOBookSheetNewName)

ИсхКолСтрок = LastRow(EOBookSheetNewName)
For I = 2 To ИсхКолСтрок
 If I > ИсхКолСтрок Then Exit For
 
 If sh.Cells(I, 1) = "Дрова технологічні (штб)" And sh.Cells(I, 3) <> "" Then
 fAkt.ListBox_Разбивка.Clear
 Set ИсхСтрока = sh.Rows(I)
 For j = 1 To 3
 КолСтрок = LastRow(EOBookSheetNewName) + 1
 sh.Cells(КолСтрок, 1).value = ИсхСтрока.Cells(1).value
 sh.Cells(КолСтрок, 2).value = ИсхСтрока.Cells(2).value
 sh.Cells(КолСтрок, 10).value = ИсхСтрока.Cells(1).value & " " & ИсхСтрока.Cells(2).value & " " & j & "м"
 sh.Cells(КолСтрок, 4).value = j
 sh.Cells(КолСтрок, 8).value = fAkt.TextBox_ВартистьЗнебособленного.value
 fAkt.ListBox_Разбивка.AddItem sh.Cells(КолСтрок, 10).value
 Call ОбьемНаВартисть(sh.Cells(КолСтрок, 9))
 Next j
 ИсхСтрока.Delete
 I = I - 1
 ИсхКолСтрок = ИсхКолСтрок - 1
 Else
 
 If sh.Cells(I, 1) = "Хлист" Then
 If I <> 1 Then sh.Cells(I, 8).value = fAkt.TextBox_ВартистьХлист.value
 
 Else
  If I <> 1 Then sh.Cells(I, 8).value = fAkt.TextBox_ВартистьЗнебособленного.value
 
 
 End If
 Call ОбьемНаВартисть(sh.Cells(I, 9))
 End If

Next I
'Добавим Хворст ліквидний
For I = 1 To 2
КолСтрок = LastRow(EOBookSheetNewName) + 1
'Наименование,'Порода,Довжина,Группа диаметрів,Сорт(Группа),Килькисть,Обьем,Цена
v = Array("Хворост ліквідний", "", "", "", I, "", "", fAkt.TextBox_ВартистьЗнебособленного.value, "", "Хворост ліквідний " & I & " грп")
КолСтрок = LastRow(EOBookSheetNewName) + 1
Call ДобавимСтрокуЩоденника(sh, КолСтрок, v)
fAkt.ListBox_Разбивка.AddItem sh.Cells(КолСтрок, 10).value
Next I
'Добавим Итоги по ліквиду
КолСтрок = LastRow(EOBookSheetNewName) + 1
'Наименование
sh.Cells(КолСтрок, 1).value = "Всього ліквідної"
Call ФормулаСуммыРегионНадЯчейкой(sh.Cells(КолСтрок, 7), sh.Cells(2, 7))
Call ФормулаСуммыРегионНадЯчейкой(sh.Cells(КолСтрок, 9), sh.Cells(2, 9))
Call ФорматЗаголовков(Range(sh.Cells(КолСтрок, 1), sh.Cells(КолСтрок, 9)))
'Запомним строку итогов ликвида
Set ИсхСтрока = sh.Rows(КолСтрок)
'Добавим Хворст неліквидний
КолСтрок = LastRow(EOBookSheetNewName) + 1
'Наименование
sh.Cells(КолСтрок, 1).value = "Хмиз неліквідний"
sh.Cells(КолСтрок, 10).value = "Хмиз неліквідний"
fAkt.ListBox_Разбивка.AddItem sh.Cells(КолСтрок, 10).value
'Добавим Итоги
КолСтрок = LastRow(EOBookSheetNewName) + 1
'Всього без ПДВ
sh.Cells(КолСтрок, 1).value = "Всього без ПДВ"
sh.Cells(КолСтрок, 1).Font.Bold = True
sh.Cells(КолСтрок, 7).Formula = "=" & ИсхСтрока.Cells(7).address & "+" & sh.Cells(КолСтрок - 1, 7).address
sh.Cells(КолСтрок, 7).Font.Bold = True
sh.Cells(КолСтрок, 9).Formula = "=" & ИсхСтрока.Cells(9).address & "+" & sh.Cells(КолСтрок - 1, 9).address
sh.Cells(КолСтрок, 9).Font.Bold = True
'ПДВ
КолСтрок = LastRow(EOBookSheetNewName) + 1
sh.Cells(КолСтрок, 1).value = "ПДВ"
sh.Cells(КолСтрок, 1).Font.Bold = True
sh.Cells(КолСтрок, 8).value = 0.2
sh.Cells(КолСтрок, 8).Font.Bold = True
sh.Cells(КолСтрок, 8).NumberFormat = "0.00%"

sh.Cells(КолСтрок, 9).Formula = "=" & sh.Cells(КолСтрок, 8).address & "*" & sh.Cells(КолСтрок - 1, 9).address
sh.Cells(КолСтрок, 9).Font.Bold = True
'Всього з пдв
КолСтрок = LastRow(EOBookSheetNewName) + 1
sh.Cells(КолСтрок, 1).value = "Всього з ПДВ"
sh.Cells(КолСтрок, 1).Font.Bold = True
sh.Cells(КолСтрок, 7).Formula = "=" & sh.Cells(КолСтрок - 2, 7).address & "+" & sh.Cells(КолСтрок - 1, 7).address
sh.Cells(КолСтрок, 9).Formula = "=" & sh.Cells(КолСтрок - 2, 9).address & "+" & sh.Cells(КолСтрок - 1, 9).address

Call ФорматЗаголовков(Range(sh.Cells(КолСтрок, 1), sh.Cells(КолСтрок, 9)))
sh.UsedRange.Select
Call СеткаСРамкой

КолСтрок = LastRow(EOBookSheetNewName)
Range(sh.Cells(2, 1), sh.Cells(КолСтрок, 5)).Select
Call СеткаСРамкой

Range(sh.Cells(КолСтрок - 2, 1), sh.Cells(КолСтрок, 5)).Select
Selection.Font.Size = 12
Call СеткаСРамкой
Range(sh.Cells(КолСтрок - 2, 6), sh.Cells(КолСтрок, 9)).Select
Selection.Font.Size = 12
Call СеткаСРамкой

Range(sh.Cells(КолСтрок - 1, 1), sh.Cells(КолСтрок, 5)).Select
Selection.Font.Size = 12
Call СеткаСРамкой
Range(sh.Cells(КолСтрок - 1, 6), sh.Cells(КолСтрок, 9)).Select
Selection.Font.Size = 12
Call СеткаСРамкой


Range(sh.Cells(КолСтрок, 1), sh.Cells(КолСтрок, 5)).Select
Selection.Font.Size = 12
Call СеткаСРамкой
Range(sh.Cells(КолСтрок, 6), sh.Cells(КолСтрок, 9)).Select
Selection.Font.Size = 12
Call СеткаСРамкой

Range(sh.Cells(ИсхСтрока.Row, 1), sh.Cells(ИсхСтрока.Row, 5)).Select
Selection.Font.Size = 12
Call СеткаСРамкой
Range(sh.Cells(ИсхСтрока.Row, 6), sh.Cells(ИсхСтрока.Row, 9)).Select
Selection.Font.Size = 12
Call СеткаСРамкой

End Sub
Sub ДобавимСтрокуЩоденника(sh As Worksheet, КолСтрок, v As Variant)
'Наименование,'Порода,Довжина,Группа диаметрів,Сорт(Группа),Килькисть,Обьем


'Наименование
sh.Cells(КолСтрок, 1).value = v(0)
'Порода
sh.Cells(КолСтрок, 2).value = v(1)
'Довжина
sh.Cells(КолСтрок, 3).value = v(2)
'Группа диаметрів
sh.Cells(КолСтрок, 4).value = v(3)
'Сорт(Группа)
sh.Cells(КолСтрок, 5).value = v(4)
'Килькисть
sh.Cells(КолСтрок, 6).value = v(5)
'Обьем
sh.Cells(КолСтрок, 7).value = v(6)
'Вартисть
sh.Cells(КолСтрок, 8).value = v(7)
'Сумма
sh.Cells(КолСтрок, 9).value = v(8)
'Примитка
sh.Cells(КолСтрок, 10).value = v(9)

End Sub



Sub Импорт(EOBookName As String, EOBookSheetName As String, EOBookSheetNewName As String)

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
If b Is Nothing Then MsgBox ("Файл " & twb.Path & "\" & EOBookName)
Workbooks(EOBookName).Worksheets(EOBookSheetName).Cells.Copy
twb.Activate
ws.Activate
Cells.Activate

ActiveSheet.Paste
'ActiveSheet.Name = EOBookSheetNewName
Workbooks(EOBookName).Close
'ThisWorkbook.Worksheets("Форма").Activate


End Sub

Sub ИмпортПоПрийманню()
Dim twb As Workbook
Dim shPriymannya As Worksheet
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("Приймання")
nrow = 9
erow = LastRow("Приймання")
For I = nrow To erow
Next I
End Sub

Sub ИмпортИзТехкарти()
Dim twb As Workbook
Dim shPriymannya As Worksheet
Dim shPrivyaska As Worksheet
Dim v As Range
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("Техкарты")
shPriymannya.Activate
Set ar = ActiveCell
q = ar.Row
Set shPrivyaska = twb.Worksheets("Привязки")
Worksheets("Форма").Activate
nrow = 2
erow = LastRow("Привязки")
For I = nrow To erow
If shPrivyaska.Cells(I, 3).value <> "" Then

Set r = FindAll(shPriymannya.Rows(1), shPrivyaska.Cells(I, 1).value)
Set v = Range(shPrivyaska.Cells(I, 3))
If r.value = "Виділ" Then
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
