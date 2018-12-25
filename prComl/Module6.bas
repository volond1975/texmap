Attribute VB_Name = "Module6"
Sub Макрос11()
Attribute Макрос11.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос11 Макрос
'

'
    Range("AK4").Select
    ActiveWorkbook.RefreshAll
End Sub
Sub Макрос12()
Attribute Макрос12.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос12 Макрос
'

'
    Cells.Select
    Range("P1").Activate
    Selection.Delete Shift:=xlUp
    Range("T11").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("S2").Select
End Sub
Sub Запрос_Техкарта()
Attribute Запрос_Техкарта.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос13 Макрос
'
Call Удаление_Техкарта
mDBQ = fAkt.ComboBox_Путь_Техкарта
mDefaultDir = ThisWorkbook.Path
ThisWorkbook.Activate
msourse = "ODBC;DSN=Файлы Excel;DBQ=" & mDBQ & ";DefaultDir=" & mDefaultDir & ";DriverId=790;MaxBufferSize=2048;PageTimeout=5;"
'    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
'        "ODBC;DSN=Файлы Excel;DBQ=G:\Dropbox\Чат\Техкарта 2015.xls;DefaultDir=G:\Dropbox\Чат;DriverId=790;MaxBufferSize=2048;PageTimeout=5;" _
'        , Destination:=Range("$A$1")).QueryTable
        Worksheets("Техкарты").Activate
         With Worksheets("Техкарты").ListObjects.Add(SourceType:=0, Source:= _
        msourse _
        , destination:=Range("$A$1")).QueryTable
        .CommandText = Array( _
        "SELECT `Техкарты$`.Шаблон, `Техкарты$`.Лісництво, `Техкарты$`.`Вид рубки`, `Техкарты$`.`Чорнобильська зона`, `Техкарты$`.`Гол#пор#`, `Техкарты$`.`Розряд висот`, `Техкарты$`.`Умови роботи`, `Техкарты$`" _
        , _
        ".`Кв-л`, `Техкарты$`.Виділ, `Техкарты$`.`Площа,га`, `Техкарты$`.`Кільк-сть дерев`, `Техкарты$`.Знач1, `Техкарты$`.Знач2, `Техкарты$`.Знач3, `Техкарты$`.Знач4, `Техкарты$`.Знач5, `Техкарты$`.Знач6, `Те" _
        , _
        "хкарты$`.Знач7, `Техкарты$`.Знач8, `Техкарты$`.Знач9, `Техкарты$`.Знач10, `Техкарты$`.Знач11, `Техкарты$`.Знач12, `Техкарты$`.`Разм хлиста`, `Техкарты$`.Масса, `Техкарты$`.`Масса на1 га`, `Техкарты$`." _
        , _
        "`Витрати кбм`, `Техкарты$`.`Витрати хлисти`, `Техкарты$`.Нелеквид, `Техкарты$`.`Складання молодняк`, `Техкарты$`.`Номер Карты`, `Техкарты$`.`К#Схилу`, `Техкарты$`.`Дата тарифних ставок`, `Техкарты$`.`" _
        , _
        "ID Карты СокращенноЛісн\Від рубки\Кв\Вид`, `Техкарты$`.`Дата договору`, `Техкарты$`.`Дата начала робіт договору`, `Техкарты$`.`Дата закінчення робіт договору`, `Техкарты$`.Відповідальний, `Техкарты$`." _
        , _
        "`Цена договору`, `Техкарты$`.`Цена договору з ПДВ`, `Техкарты$`.`Сумма договору з ПДВ`, `Техкарты$`.ЛК, `Техкарты$`.`Дата ЛК`, `Техкарты$`.`Коэфициент складання`, `Техкарты$`.`Способ рубки молодняків`" _
        , _
        ", `Техкарты$`.`Путь к договору`, `Техкарты$`.`Путь к техкарте`, `Техкарты$`.`Путь к кошторису`, `Техкарты$`.`Залишок по ТК`, `Техкарты$`.Столбец1, `Техкарты$`.Виконавець, `Техкарты$`.Квиконавець, `Тех" _
        , "карты$`.`Підготовчі `, `Техкарты$`.Код FROM `Техкарты$`")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Таблица_Запрос_из_Файлы_Excel"
        .Refresh 'BackgroundQuery:=False
    End With
    ActiveSheet.ListObjects(1).name = "Техкарти"
End Sub
Sub Удаление_Техкарта()
Attribute Удаление_Техкарта.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос14 Макрос
'

On Error Resume Next
ThisWorkbook.Activate
 Worksheets("Техкарты").Activate
    ActiveSheet.ListObjects("Техкарти").Unlist
    Cells.Select
    Range("P1").Activate
    Selection.ClearContents
End Sub

Sub Макрос15()
'
' Макрос14 Макрос
'

'
    Range("Q4").Select
   Worksheets("Техкарты").ListObjects("Техкарти").Unlist
    Cells.Select
    Range("P1").Activate
    Selection.ClearContents
End Sub
