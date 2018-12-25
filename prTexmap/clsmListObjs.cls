VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsmListObjs"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim m_wb       As Excel.Workbook
Dim m_Tables   As Collection
Dim m_Tables_ListColumns As Collection
Dim m_Tables_ListRows As Collection
Dim m_ValueListObject As Range
'===============================================
'ВЕРСИЯ КЛАССА
Private Const clsmListObjser = "0.0 (18.06.2015)"

'тип для перечня ошибок
Private Type tError
    NoneErr    As String
    CreateADOErr As String
    NoData     As String
    DataNotFound As String
    QueryError As String
    ConnectError As String
    DBConnectError As String
    WorkBookErr As String
    MultyRange As String
End Type


'внутренние объекты класса
Private sDataSource As String   'имя книги
Private sSheetSource As String  'имя листа
Private sNameSource As String  'имя таблицы
Private sNameStyle As String  'имя стиля таблицы

Private sActiveListObjectName As String
'рабочие пременные
Private iExcelVer As Integer    'версия EXCEL

'переключатели
Private bHeader As Boolean      'требует переподключения
Private bSubTotal As Boolean      'требует переподключения
'процедура, доступна только в этом классе
'формируем описание возможных ошибок
' %1 - меняется на текст подстановки
Private Sub sbInitError()
    eError.NoneErr = ""
    eError.CreateADOErr = "КРИТИЧЕСКАЯ ОШИБКА! НЕВОЗМОЖНО СОЗДАТЬ ОБЪЕКТЫ ADO"
    eError.NoData = "Ошибка. Запрос не создан"
    eError.DataNotFound = "Нет записей удоволетворяющих запросу '%1'"
    eError.QueryError = "Ошибка запроса: '%1'"
    eError.ConnectError = "Нет подключения к DB (сперва надо подключиться к '%1' используя метод BookConnect)"
    eError.DBConnectError = "Ошибка подключения к DB (проверьте правильность пути и наличия файла '%1')"
    eError.WorkBookErr = "Файл данных '%1' отсутствует"
    eError.MultyRange = "Задан множественный диапазон '%1' запрос не возможен"
End Sub

Public Property Get NewEnum() As IUnknown
'the following line, added in a text editor,
'creates the ability to cycle through the items with For Each
    Set NewEnum = m_Tables.[_NewEnum]
End Property

Public Function Initialize(WbWithTables As Excel.Workbook)
'Call sbInitError
    Call sbCreateDefValue

    Set m_wb = WbWithTables
    Call Refresh

End Function

Public Sub Refresh()
    Dim ws     As Excel.Worksheet
    Dim lo     As Excel.ListObject

    Set m_Tables = New Collection
    For Each ws In m_wb.Worksheets
        For Each lo In ws.ListObjects
            m_Tables.Add lo, lo.name
        Next lo
    Next ws
End Sub

Public Property Get Item(Index As Variant) As Excel.ListObject
'the following line, added in a text editor,
'sets Item as the default property of the class
    Set Item = m_Tables(Index)
End Property

Public Property Get Count()
    Count = m_Tables.Count
End Property

Property Get Exists(Index As Variant) As Boolean
    Dim Test   As Variant
    On Error Resume Next
    Set Test = m_Tables(Index)
    Exists = Err.Number = 0
End Property
Property Get SheetExists(Index As Variant) As Boolean
    Dim Test   As Variant
    On Error Resume Next
    Set Test = m_wb.Worksheets(Index)
    SheetExists = Err.Number = 0
End Property



Public Property Get items() As Collection

    Set items = m_Tables
End Property

Function AddListObject(ByVal sSheetSource, Optional ByVal sNameSource, Optional ByVal sUSRRange As Range, Optional sColumnName As String, Optional ByVal sNameStyle As String = "TableStyleLight1", Optional bHead As Boolean = xlYes, Optional bAutoNum As Boolean = xlNo, Optional bTotals As Boolean = False, Optional bShowAutoFilter As Boolean = True, Optional bAddNumcolumn As Boolean = False) As ListObject


    Dim lo     As ListObject
Dim loc As ListColumn
    'SourceType
    'xlSrcExternal 0 External data source (Microsoft SharePoint Foundation site).
    'xlSrcModel 4 PowerPivot Model
    'xlSrcQuery 3 Query
    'xlSrcRange 1 Range
    'xlSrcXml 2 XML

If IsMissing(sNameSource) Then sNameSource = ActiveSheet.name
If sUSRRange Is Nothing Then Set sUSRRange = ActiveCell.CurrentRegion

    'Source , LinkSource, HasHeaders, Destination
    Set lo = m_wb.Worksheets(sUSRRange.parent.name).ListObjects.Add(xlSrcRange, sUSRRange, , bHead)
    With lo
        .name = _
        sNameSource

        .TableStyle = sNameStyle
        .ShowTotals = bTotals
   .ShowHeaders = bHead

        If bHead = True Then .ShowAutoFilter = bShowAutoFilter
If bAddNumcolumn Then
 Set loc = lo.ListColumns.Add
loc.name = "№"
End If
    If Not IsEmpty(sColumnName) And Not sColumnName = "" Then
            .ListColumns(1).name = sColumnName
        End If
        If bAutoNum Then

            .ListRows.Add
            lo.DataBodyRange.Cells(1, 1).FormulaR1C1 = "=ROW()-1"


        End If
    End With
    Set AddListObject = lo
Call Refresh
End Function
Sub ListTableStylesName()
' Use this procedure to list all the named table styles in the
' Immediate window. Put the cursor in this procedure and press F5
' to run it. Look in the Immediate window for the results.
    Dim ts     As TableStyle
    For Each ts In ActiveWorkbook.TableStyles
        Debug.Print ts.name
    Next ts
End Sub

'---------------------------------------------------------
'установить для класса рабочую книгу
'возможно автоматическое переподключение
Property Let Workbook(ByVal sFullNameBook As String)
    Dim bOk    As Boolean
    If sDataSource = sFullNameBook Then Exit Property
    If Dir(sFullNameBook) <> "" Then
        sDataSource = sFullNameBook
        '        bOk = fnReconnect()
    Else
        sbSetError sErrText:=eError.WorkBookErr, sReplaceParam:=sFullNameBook
    End If
End Property

'показать установленную для класса рабочую книгу
Property Get Workbook() As String
    Workbook = sDataSource
End Property

'установить для класса рабочий лист
Property Let Worksheet(ByVal sSheetName As String)
    If sSheetSource = sSheetName Then Exit Property
    sSheetSource = sSheetName
End Property

'показать текущий для класса рабочий лист
Property Get Worksheet() As String
    Worksheet = sSheetSource
End Property
'NameListObject

'установить для класса рабочий лист
Property Let name(ByVal sNameListObject As String)
    If sNameSource = sNameListObject Then Exit Property
    sNameSource = sNameListObject
End Property

'показать текущий для класса рабочий лист
Property Get name() As String
    name = sNameSource
End Property







'Property Get ActiveListObjectName() As String
'    ActiveListObjectName = sActiveListObjectName
'End Property




'установить для класса стиль таблицы
Property Let NameStyle(ByVal sNameStyleTable As String)
    If sNameStyle = sNameStyleTable Then Exit Property
    sNameStyle = sNameListObjectTable
End Property

'показать текущий для класса стиль таблицы
Property Get NameStyle() As String
    NameStyle = sNameStyle
End Property














'установить для класса нужную область на листе
'sRange - строка содержащая адрес
'например   "A" или "DF57" или "G3:EJ500" ...
'всегда нормализуется к диапазону
Property Let RangeStr(ByVal sRange As String)

    If sRange <> "" Then
        sRange = fnGetNormaliseAddr(sRange, True)
        eLastChange = eLastChange Or eRange    'пользователь меняет диапазон
    Else
        eLastChange = eLastChange Xor eRange    'пользователь отменил диапазон
    End If
    sUSRRange = sRange
End Property

'посмотреть (вернуть) заданную для класса область на листе
Property Get RangeStr() As String
    RangeStr = sUSRRange
End Property

'установить для класса нужную область на листе
'параметр - область
'напрмер Sheets("Info").Range("A23:D784") или Range("В3:B8") ...
Property Let RangeRng(ByVal rngRange As Range)
    Dim sRange As String, sBook As String, sBookPath As String
    'если заданы не связанные области - выходим
    If rngRange.Areas.Count > 1 Then
        sbSetError 0, eError.MultyRange, rngRange.Address
        Exit Property
    End If
    sRange = rngRange.Address

    'ПРОВЕРЯЕМ НА СООТВЕТСТВИЕ ЛИСТА И КНИГИ
    sSheetSource = rngRange.parent.name
    sBook = rngRange.parent.parent.name
    sBookPath = rngRange.parent.parent.Path

    sRange = Replace(sRange, "$", "")
    eLastChange = eLastChange Or eRange
    sUSRRange = sRange

    If fnGetFileName(sDataSource) <> sBook Then
    End If

End Property

'посмотреть (вернуть) заданную для класса область на листе
Property Get RangeRng() As Range
    Dim sBookName As String
    sBookName = fnGetFileName(sDataSource)
    Set RangeRng = Workbooks(sBookName).Sheets(sSheetSource).Range(sUSRRange)
End Property
'ФУНКЦИЯ!!!! - доступна тут и у наследников
'возвращает "правильно" сформированный адрес в формате А1
'если установлено bMulti - возвращает в виде диапазона
' fnGetNormaliseAddr("A") вернет "A1"
' fnGetNormaliseAddr("A", True) вернет "A:A"
' fnGetNormaliseAddr("A1", True) вернет "A1:A1"
' fnGetNormaliseAddr("Af233:bb13", False) вернет "AF13"
' fnGetNormaliseAddr("Af233:bb13", True) вернет "Af233:bb13"
' fnGetNormaliseAddr("Af233:6bb13", False) вернет "AF233"
' fnGetNormaliseAddr("Af233:6bb13", True) вернет "Af233:Af233"   <<<<< 6bb13 - не верно заданный адрес
Friend Function fnGetNormaliseAddr(ByVal sRng As String, Optional bMulti As Boolean = False) As String
    Dim iPos As Integer, lCnt As Long, tmpRng As Range
    fnGetNormaliseAddr = "A1"

    'убираем знаки доллара и пробелы
    sRng = Replace(sRng, "$", ""): sRng = Trim(sRng)

    'если в диапазоне что-то указано
    If sRng <> "" Then
        iPos = InStr(sRng, ":")
        ' ?????
        If iPos = 0 Then
            If fnCheckShablon(sRng) Then
                sRng = sRng & ":" & sRng
            Else
                sRng = "A1:A1"
            End If
            ' *:*
        Else
            ' ?????:
            If iPos = Len(sRng) Then
                'проверяем на соответствие шаблону 1 часть адреса
                If fnCheckShablon(Mid(sRng, 1, iPos - 1)) Then
                    sRng = sRng & Mid(sRng, 1, iPos - 1)
                Else
                    sRng = "A1:A1"
                End If
                ' ????:????
            Else
                'проверяем на соответствие шаблону 1 часть адреса
                If Not fnCheckShablon(Mid(sRng, 1, iPos - 1)) Then
                    sRng = "A1" & Mid(sRng, iPos)
                    iPos = 3
                End If
                'проверяем на соответствие шаблону 2 часть адреса
                If Not fnCheckShablon(Mid(sRng, iPos + 1)) Then
                    sRng = Mid(sRng, 1, iPos) & Mid(sRng, 1, iPos - 1)
                End If
            End If
        End If
        'sRng - в нем ни чего нет - пусто
    Else
        sRng = "A1:A1"
    End If
    'проверим на правильность формирования диапазона
    'путем попытки подсчета в нем кол-ва ячеек
    On Error Resume Next
    lCnt = Range(sRng).Count
    'ошибка! что-то не так :( - сформируем из 1 части
    If Err <> 0 Then
        sRng = Mid(sRng, 1, iPos) & Mid(sRng, 1, iPos - 1)
    End If
    On Error GoTo 0
    Err.Clear
    'если не задан режим области ("мульти"), сформируем
    ' из области sRng - адрес левой верхней ячейки
    If Not bMulti Then
        sRng = Range(sRng).Cells(1, 1).Address
    End If
    'вернем нормализованный адрес
    fnGetNormaliseAddr = Replace(sRng, "$", "")
End Function


'определить тип области
Private Function fnAreaType(RangeArea As Range) As String
'   Возврат типа диапазона в области
    Select Case True
    Case RangeArea.Cells.CountLarge = 1
        AreaType = "Ячейка"
    Case RangeArea.CountLarge = Cells.CountLarge
        AreaType = "Рабочий лист"
    Case RangeArea.Rows.Count = Cells.Rows.Count
        AreaType = "Столбец"
    Case RangeArea.Columns.Count = Cells.Columns.Count
        AreaType = "Строка"
    Case Else
        AreaType = "Блок"
    End Select
End Function
'---------------------------------------------------------
'---------------------------------------------------------
'ФУНКЦИЯ!!!! - доступна тут и у наследников
'проверить строчку на соответствие шаблону адреса стиля А1
'сначала буквы (латынь) максимум 3 штуки, затем цифры максимкм 7 цифр
'если соответствует шаблону - возвращает True
Friend Function fnCheckShablon(ByVal sRng As String) As Boolean
    fnCheckShablon = False
    'удаляем знаки доллара и пробелы
    sRng = Replace(sRng, "$", ""): sRng = Trim(sRng)

    'A - ZZZ (указаны только буквы)
    If sRng Like "[A-Za-z]" Or _
       sRng Like "[A-Za-z][A-Za-z]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z]" Then
        fnCheckShablon = True
        Exit Function
    End If
    'A1 - Z12345678  (указана буква и цифра/ы)
    If sRng Like "[A-Za-z][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
        fnCheckShablon = True
        Exit Function
    End If
    'AZ1 - AZ12345678 (2 буквы и цифры)
    If sRng Like "[A-Za-z][A-Za-z][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
        fnCheckShablon = True
        Exit Function
    End If
    'ABZ1 - ZZZ12345678 (3 буквы и цифры)
    If sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9]" Or _
       sRng Like "[A-Za-z][A-Za-z][A-Za-z][0-9][0-9][0-9][0-9][0-9][0-9][0-9]" Then
        fnCheckShablon = True
    End If
End Function

'---------------------------------------------------------
'---------------------------------------------------------
'ФУНКЦИЯ!!!! - доступна тут и у наследников
'fnGetFileName - вернуть имя файла вместе с расширением или ""
'pFulNameFile - имя файла (может содержать полный путь) (по значению)
Friend Function fnGetFileName(ByVal strFullNameFile As String) As String
    Dim iDirPos As Integer: iDirPos = 0
    Dim strPS  As String: strPS = Application.PathSeparator

    fnGetFileName = Mid(strFullNameFile, InStrRev(strFullNameFile, strPS) + 1)

End Function



'ФУНКЦИЯ!!!! - доступна тут и у наследников
'fnGetFileExt - вернуть расширение файла вместе с точкой или ""
'pNameFile - имя файла (может содержать несколько точек, обрабатывается последняя) (по значению)
Friend Function fnGetFileExt(ByVal strNameFile As String) As String
    Dim iDotPos As Integer: iDotPos = 0
    Dim strRetval As String: strRetval = ""

    iDotPos = InStrRev(strNameFile, ".")
    If iDotPos Then strRetval = LCase(Mid(strNameFile, iDotPos))

    fnGetFileExt = strRetval
End Function



'ФУНКЦИЯ!!!! - доступна тут и у наследников
'fnGetFilePath - вернуть путь к файлу
'pFulNameFile - имя файла включая путь (по значению)
Friend Function fnGetFilePath(ByVal strFullNameFile As String) As String
    Dim iDirPos As Integer: iDirPos = 0
    Dim strRetval As String: strRetval = ""
    Dim strPS  As String: strPS = Application.PathSeparator

    iDirPos = InStrRev(strFullNameFile, strPS)
    If iDirPos Then strRetval = Mid(strFullNameFile, 1, iDirPos)

    fnGetFilePath = strRetval

End Function


'ФУНКЦИЯ!!!! - доступна тут и у наследников
'возвращает путь до заданного вложения
Friend Function fnGetPathLevel(ByVal strFullPath As String, Optional ByVal iLvl As Integer = 9999) As String
    Dim tmpPos As Integer: tmpPos = 0
    Dim DirPos As Integer: DirPos = 0
    Dim strPS  As String: strPS = Application.PathSeparator
    Dim retval As String: retval = ""

    Do
        tmpPos = InStr(DirPos + 1, strFullPath, strPS)
        If tmpPos Then
            DirPos = tmpPos
            iLvl = iLvl - 1
            If iLvl = -1 Then Exit Do
        End If
    Loop While tmpPos
    If DirPos Then retval = Mid(strFullPath, 1, DirPos)

    fnGetPathLevel = retval

End Function
'---------------------------------------------------------
'процедура, доступна только в этом классе
'установить (дать,определить) описание ошибки
'sReplaceParam - заменяет в sErrText %1 на sReplaceParam
Private Sub sbSetError(Optional ByVal lParamErr As Long = 0, _
                       Optional ByVal sErrText As String = "", _
                       Optional ByVal sReplaceParam As String = "")
    sError = sErrText
    lErr = lParamErr
    If lParamErr <> 0 Then
        If sError <> "" Then sError = sError & vbCrLf
        sError = sError & Error(lErr) & " (" & lErr & ")"
    End If

    If InStr(sError, "%1") <> 0 Then
        sError = Replace(sError, "%1", sReplaceParam)
    End If
End Sub

'вернуть описание ошибки
Property Get GetError()
    GetError = sError
End Property

'Возвращает - True - если была ошибка
' Dim ado As New ClsADO
' ...
' ado.WorkBook = "БазаПроверок"
' If ado.IsError Then
'   тут что-то делаем в случае ошибки
' End If
Property Get IsError() As Boolean
    IsError = CBool(sError <> "" Or lErr <> 0)
End Property
'установить или отменить авто заголовки в рекордсете
'True - заголовки беруться из первой строки таблицы (диапазона)
'False - заголовки формируются и используются автоматически
'        в формате Fx - Например F1, F5, F158 ...
'возможно автоматическое переподключение
Property Let Header(ByVal bSetHeader As Boolean)
    Dim bOk    As Boolean
    If bHeader <> bSetHeader Then
        bHeader = bSetHeader
        bOk = fnReconnect()
    End If
End Property

'вернуть (показать) настройку по авто заголовкам
Property Get Header() As Boolean
    Header = bHeader
End Property
'---------------------------------------------------------
'процедура, доступна только в этом классе
'создание значений "по умолчанию"
'при создании объекта утанавливается активная рабочая книга, текущий рабочий лист
'и текщая занятая область
Sub sbCreateDefValue()
    Dim rngRaneSourse As Range

    iExcelVer = Val(Mid(Application.version, 1, InStr(1, Application.version, ".") - 1))
    sDataSource = ActiveWorkbook.FullName
    sSheetSource = ActiveSheet.name
    sNameStyle = "TableStyleLight10"
    bHeader = True
    bReadOnly = True
    bReplaseNull = False
    bAutoConnect = True
    lRecCount = 0
    Set rngRaneSourse = Sheets(sSheetSource).Cells(1, 1).CurrentRegion
    sUSRWhere = ""
    sUSRRange = ""
    sUSRSelect = ""
    sSelectSql = "SELECT * FROM [" & sSheetSource & "$" & Replace(rngRaneSourse.Address, "$", "") & "];"
    eLastChange = eNone
    sError = "": lErr = 0
    bRSIsReady = False
End Sub

'Отримання адреси діапазону

Function TABLEDATARANGE(ByVal loTest As ListObject) As String
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Returns the range address of the specified table's data body range.
'---------------------------------------------------------------------------------------------
    If Not loTest.DataBodyRange Is Nothing Then
        TABLEDATARANGE = loTest.DataBodyRange.Address
    End If
End Function

Function TABLEHEADERRANGE(ByVal loTest As ListObject) As String
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Returns the range address of the specified table's header row range.
'---------------------------------------------------------------------------------------------
    If Not loTest.HeaderRowRange Is Nothing Then
        TABLEHEADERRANGE = loTest.HeaderRowRange.Address
    End If
End Function

Function TABLETOTALSRANGE(ByVal loTest As ListObject) As String
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Returns the range address of the specified table's totals row range.
'---------------------------------------------------------------------------------------------
    If Not loTest.TotalsRowRange Is Nothing Then
        TABLETOTALSRANGE = loTest.TotalsRowRange.Address
    End If
End Function
Function TABLEHEADERVISIBLE(ByVal TableCell As Range) As Boolean
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Looks at a cell to see if it is part of a table and if the header is
'///                is visible.
'---------------------------------------------------------------------------------------------
    Dim loTest As ListObject

    '/// Set temporary table
    On Error Resume Next
    Set loTest = TableCell.ListObject
    On Error GoTo 0

    '/// Check if it is a table (ListObject)
    If loTest Is Nothing Then
        TABLEHEADERVISIBLE = CVErr(xlErrNA)
        Exit Function
    End If

    '/// Return visibility status of table header
    TABLEHEADERVISIBLE = loTest.ShowHeaders

End Function

Function TABLETOTALSROWVISIBLE(ByVal TableCell As Range) As Boolean
'---------------------------------------------------------------------------------------------
'/// Created on:    13-Jan-2013
'/// Created by:    Zack Barresse
'/// Purpose:       Looks at a cell to see if it is part of a table and if the totals row is
'///                is visible.
'---------------------------------------------------------------------------------------------
    Dim loTest As ListObject

    '/// Set temporary table
    On Error Resume Next
    Set loTest = TableCell.ListObject
    On Error GoTo 0

    '/// Check if it is a table (ListObject)
    If loTest Is Nothing Then
        TABLETOTALSROWVISIBLE = CVErr(xlErrNA)
        Exit Function
    End If

    '/// Return visibility status of table totalsrow
    TABLETOTALSROWVISIBLE = loTest.ShowTotals

End Function

Function DrawValueColumnListObject(ls As ListObject, NameColumn As String, Optional bHead As Boolean = False) As Range
    Dim wb     As Workbook
    Dim ws     As Worksheet

    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
    Dim HeadVisible As Boolean
    If ls.HeaderRowRange Is Nothing Then z = 1
    If Not ls.HeaderRowRange Is Nothing And bHead Then z = 1
    If Not ls.HeaderRowRange Is Nothing And bHead = False Then z = 2
    Set lsc = ls.ListColumns(NameColumn)

    Set DrawValueColumnListObject = Range(lsc.Range.Cells(z), lsc.Range.Cells(lsc.Range.Cells.Count))
End Function




Function СписокЗначенийСтолбцаУмнойТаблицыUnicum(ls As ListObject, NameColumn As String, Optional bHead As Boolean = True, Optional bTotals As Boolean = False)
    Dim v      As Range
    Set v = UnicumRange(СписокЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца))
    Set СписокЗначенийСтолбцаУмнойТаблицыUnicum = v
End Function

Function UnicumRange(r As Range) As Range

'Возвращает уникальные данные из региона

    Dim v()    As Variant
    s = 1
    For i = 1 To r.Cells.Count
        ReDim Preserve v(s - 1)

        For j = 0 To s - 1
            If r.Cells(i) = v(j) Then
                GoTo 111
            End If
        Next j
        v(s - 1) = r.Cells(i).value
        s = s + 1
111:
    Next i
    UnicumRange = v
End Function


Sub CreateSheetsFromList(wb As Workbook, tbl As ListObject)
'PURPOSE: Create new worksheets from a list of names within a table
'SOURCE: www.TheSpreadsheetGuru.com/the-tCode-vault

    Dim NewSheet As Worksheet
    Dim X      As Integer
    Dim tbl    As ListObject
    Dim cell   As Range

    'Opitimize tCode
    Application.ScreenUpdating = False

    'Store table in a variable
    '  Set tbl = Worksheets("Sheet1").ListObjects("Table1")

    'Create a new worksheet for every name inside the table
    For Each cell In tbl.DataBodyRange.Cells
        If SheetExists(cell.value) = False And cell.value <> "" Then
            Set NewSheet = wb.Sheets.Add(After:=Sheets(Sheets.Count))
            NewSheet.name = cell.value
        End If
    Next cell

    'Opitimize tCode
    Application.ScreenUpdating = True

End Sub

Function myFilterListObject(ls As ListObject, NameColumn As String, Optional myCriteria As String) As Integer

    Dim lsc    As ListColumn

    Dim cc     As Range

    Dim z      As Integer
    On Error GoTo FilterColumnListObject_Error


    z = IndexColumn(ls, NameColumn)
    If Not myCriteria = "" Then
        ls.Range.AutoFilter Field:=z, Criteria1:= _
                            myCriteria
    Else
        ls.Range.AutoFilter Field:=z
    End If
    myFilterListObject = z
    On Error GoTo 0
    Exit Function

FilterColumnListObject_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure FilterColumnListObject of Module УмныеТаблицы"
End Function



'-----------------------------------------------------------------
'СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
'-------------------------------------------------------------------
Public Function ListColumns_Refresh(ls As ListObject) As ListColumns
    Dim lc     As ListColumn

    Set m_Tables_ListColumns = New Collection
    For Each lc In ls.ListColumns
        If ls.HeaderRowRange Is Nothing Then
        Else
            m_Tables_ListColumns.Add ls, lc.Index
        End If
    Next lc
    Set ListColumns_Refresh = m_Tables_ListColumns
End Function






Function IndexColumn(ls As ListObject, NameColumn As String)
    Dim wb     As Workbook
    Dim ws     As Worksheet
    'Dim ls As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim cc     As Range
    Dim rr     As Range
    Dim r      As Range

    Set zagrow = ls.HeaderRowRange
    IndexColumn = Application.WorksheetFunction.match(NameColumn, zagrow, 0)


End Function
Function ExistColumn(Index, NameColumn As String)
    Dim wb     As Workbook
    Dim ws     As Worksheet
    'Dim ls As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim cc     As Range
    Dim rr     As Range
    Dim r      As Range

    Set zagrow = m_Tables(Index).HeaderRowRange
    Set ExistColumn = zagrow.Find(NameColumn, LookIn:=xlValues, LookAt:=xlWhole)


End Function




Function ListObjectColumnExist(ls As ListObject, NameColumn As String)
    On Error Resume Next
    'добавляет столбец с именем LCName в умную таблицу
    Dim n
    Set n = ls.ListColumns(NameColumn)

    ListObjectColumnExist = Err.Number = 0

End Function

Function ListObjectColumnAdd(ls As ListObject, NameColumn As String, Optional poz As Long, Optional xlFormulTotalConstant As Integer = -1, Optional formul As String, Optional form_at As String, Optional bReplaseValue As Boolean)
'добавляет столбец с именем NameColumn в умную таблицу в позицию poz или в конец таблицы
'xlFormulTotalConstant>=0 добавляем строку итогов и устанавливаем формулу из списка
'formul <> "" добавляем формулу в столбец форматируем в form_at и если bReplaseValue=Истина заменяем ее значениями
    Dim lcs    As ListColumn
    If poz > 0 Then
        Set lcs = ls.ListColumns.Add(poz)
    Else
        Set lcs = ls.ListColumns.Add
    End If
    ls.HeaderRowRange(lcs.Index) = NameColumn
    If xlFormulTotalConstant >= 0 Then
  If Not ls.ShowTotals Then ls.ShowTotals = True
  lcs.TotalsCalculation = formulTotal
    End If
    If formul <> "" Then
   Call ListObjectColumnFormula(lcs, formul, form_at)
    If bReplaseValue Then
        For Each r In lcs.Range.Cells
            r.value = r.value
        Next
    End If
    End If
    Set ListObjectColumnAdd = lcs
End Function



Function ВидимыйДиапазонЗначенийУмнойТаблицыБезЗаголовка(ИмяЛиста, ИмяТаблицы) As Range
    Dim wb     As Workbook
    Dim ws     As Worksheet
    Dim ls     As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
    Set ls = ws.ListObjects(ИмяТаблицы)

    Set r = ls.DataBodyRange
    Set ВидимыйДиапазонЗначенийУмнойТаблицыБезЗаголовка = r.SpecialCells(xlCellTypeVisible)


End Function

Function VisibleRangeColumnListObject(ls As ListObject, NameColumn As String) As Range
    Dim wb     As Workbook
    Dim ws     As Worksheet
    '    Dim ls     As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
    Dim rr     As Range
    On Error Resume Next
    '    Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
    '    Set ls = ws.ListObjects(ИмяТаблицы)
    Set lsc = ls.ListColumns(ИмяСтолбца)

    Set rr = ls.DataBodyRange
    Set www = Intersect(lsc.Range, rr)
    Set r = www.SpecialCells(xlCellTypeVisible)
    'Set r = lsc.Range.SpecialCells(xlCellTypeVisible)
    Set VisibleRangeColumnListObject = r


End Function


Sub AddComment(Index As Variant, strComment)
'===========================================================
'18.06.201520:34
'===========================================================
'добавить комментарий к таблице (показывает в качестве комментария к
'rangename в менеджере имен)
'Обратите внимание, что такие имена диапазонов не могут быть удалены !!
'Название диапазона удаляется , как только таблица преобразуется в диапазон
'===========================================================

    m_Tables(Index).Comment = strComment
End Sub
Function Comment(Index As Variant)
'===========================================================
'18.06.201520:34
'===========================================================
'добавить комментарий к таблице (показывает в качестве комментария к
'rangename в менеджере имен)
'Обратите внимание, что такие имена диапазонов не могут быть удалены !!
'Название диапазона удаляется , как только таблица преобразуется в диапазон
'===========================================================

   Comment = m_Tables(Index).Comment
End Function



Sub ShowTableParts(CreateListObjectTable As ListObject, Optional ByVal bState As Boolean = True)

'===========================================================
'22.06.201508:28
'===========================================================
'Ця процедура має логічне (True / False) змінна передається до нього,
'щоб включити або виключити всі частини таблиці (Не форматування)
'===========================================================
'/// Set misc variables

    With CreateListObjectTable
        .ShowTotals = bState
        .ShowHeaders = bState
        If bState = True Then .ShowAutoFilter = bState
    End With
End Sub





'======================================================================

'Эмуляция Active для умной таблицы
'======================================================================
'Обьекты
'======================================================================
'ActiveListObject Возвращает активную умную таблицу тип ListObject
'ActiveListObjectActiveRow Возвращает активную  строку умной таблицы тип ListRow
'ActiveListObjectActiveColumn Возвращает активный  столбец умной таблицы тип ListColumn
'======================================================================
'Регион
'======================================================================
'ActiveListObjectHederRowRange Возвращает регион Заголовков активной умной таблицы тип Range
'ActiveListObjectTotalRowRange Возвращает регион Итогов активной умной таблицы тип Range
'ActiveListObjectDataBodyRange Возвращает регион Данных активной умной таблицы  тип Range
'======================================================================

'======================================================================
'ActiveListObjectIndexColumn Возвращает индекс активного столбца  умной таблицы тип Long

'======================================================================
'Действия
'======================================================================
'  С активной строкой
'======================================================================
'ActiveListObjectRowDelete Удаляет активную строку умной таблицы
'ActiveListObjectRowAdd
'======================================================================
Friend Function ActiveListObject()
Dim ws     As Excel.Worksheet
Dim lo     As Excel.ListObject
If ActiveSheet.ListObjects.Count = 0 Then Exit Function
    Set ws = ActiveSheet
        For Each lo In ws.ListObjects
        If Not Intersect(ActiveCell, lo.Range) Is Nothing Then
          Set ActiveListObject = lo
          Exit Function
          Else
          Set ActiveListObject = Nothing
        End If
        Next lo
 End Function
Friend Function ActiveListObjectHederRowRange()
Dim ws     As Excel.Worksheet
Dim lo     As Excel.ListObject
If ActiveSheet.ListObjects.Count = 0 Then Exit Function


    Set ws = ActiveSheet
    
        For Each lo In ws.ListObjects
        If Not Intersect(ActiveCell, lo.Range) Is Nothing Then
        If lo.ShowHeaders Then
          Set ActiveListObjectHederRowRange = lo.HeaderRowRange
          Else
          Set ActiveListObjectHederRowRange = Nothing
        End If
        End If
        Next lo
   

End Function
Function ListObjectActivate(Index As Variant)
m_Tables(Index).parent.Activate
m_Tables(Index).Range.Cells(1).Activate
Set ListObjectActivate = m_Tables(Index)
End Function

Function ListObjectAddRecordset(sh As Worksheet, lo_Name, Optional TargetRange As String)
Dim lo As ListObject
If TargetRange = "" Then TargetRange = "A1"
sh.Range(TargetRange).Activate

sh.ListObjects.Add(xlSrcQuery, oRS, _
  Destination:=Selection).QueryTable.Refresh
Set lo = sh.ListObjects(1)
lo.name = lo_Name
Set ListObjectAddRecordset = lo_Name
End Function





Friend Function ActiveListObjectTotalRowRange()
Dim ws     As Excel.Worksheet
Dim lo     As Excel.ListObject
If ActiveSheet.ListObjects.Count = 0 Then Exit Function


    Set ws = ActiveSheet
    
        For Each lo In ws.ListObjects
        If Not Intersect(ActiveCell, lo.Range) Is Nothing Then
        If lo.ShowTotals Then
          Set ActiveListObjectTotalRowRange = lo.TotalsRowRange
          Else
          Set ActiveListObjectTotalRowRange = Nothing
        End If
        End If
        Next lo
   

End Function

Friend Function ActiveListObjectDataBodyRange()
Dim ws     As Excel.Worksheet
Dim lo     As Excel.ListObject
Set ws = ActiveSheet
If ws.ListObjects.Count = 0 Then Exit Function


    

        For Each lo In ws.ListObjects
        If Not Intersect(ActiveCell, lo.Range) Is Nothing Then
        
          Set ActiveListObjectDataBodyRange = lo.DataBodyRange
          Else
          Set ActiveListObjectDataBodyRange = Nothing
        End If
       
        Next lo
   

End Function




Friend Function ActiveListObjectActiveRow()
Dim ws     As Excel.Worksheet
Dim lo     As Excel.ListObject
Dim lr     As Excel.ListRow
If ActiveSheet.ListObjects.Count = 0 Then Exit Function


    Set ws = ActiveSheet
    
        For Each lo In ws.ListObjects
        If Not Intersect(ActiveCell, lo.Range) Is Nothing Then
        For Each lr In lo.ListRows
        If Not Intersect(ActiveCell, lr.Range) Is Nothing Then
          Set ActiveListObjectActiveRow = lr
        Exit Function
          End If
          Next lr
        Else
        Set ActiveListObjectActiveRow = Nothing
        
        End If
       
        Next lo
   

End Function
Friend Function ActiveListObjectActiveColumn()
Dim ws     As Excel.Worksheet
Dim lo     As Excel.ListObject
Dim lc     As Excel.ListColumn
If ActiveSheet.ListObjects.Count = 0 Then Exit Function


    Set ws = ActiveSheet
    
        For Each lo In ws.ListObjects
        If Not Intersect(ActiveCell, lo.Range) Is Nothing Then
        For Each lc In lo.ListColumns
        If Not Intersect(ActiveCell, lc.Range) Is Nothing Then
          Set ActiveListObjectActiveColumn = lc
        Exit Function
          End If
          Next lc
        Else
        Set ActiveListObjectActiveColumn = Nothing
        
        End If
       
        Next lo
   

End Function

Function ActiveListObjectColumnAdd(Index As Variant, Optional Положение, Optional name As String) As ListColumn



'Добавляет новый столбец, представленной указанным ListObject
'Положение-Определяет относительное положение нового столбца его имени



    Dim wb     As Workbook
    Dim ws     As Worksheet
  Dim lo As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
    
   On Error GoTo ActiveListObjectColumnAdd_Error

 Set lo = m_Tables(Index)
    If IsMissing(Положение) Then
        Set lsc = lo.ListColumns.Add
    Else
  
        Set lsc = lo.ListColumns.Add(Положение)
    End If
    
    If Not IsMissing(name) Then
    
    
     lsc.name = name
     End If
  
    
    Set ActiveListObjectColumnAdd = lsc

   On Error GoTo 0
   Exit Function

ActiveListObjectColumnAdd_Error:
'If Err = 1004 Then
'Err.Description
'Else
MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure ActiveListObjectColumnAdd of Class Module clsmListObjs"
'End If
    
End Function

'ФОРМУЛЫ В СТОЛБЕЦ УМНОЙ ТАБЛИЦЫ
'---------------------------------------------------------------------------------------
Sub ListObjectColumnFormulaLocal(lcs As ListColumn, myFormulaLocal As String, Optional NumbFormat)
    Dim lo     As ListObject
    Dim r      As Range
    Set lo = lcs.parent
    Set r = Intersect(lo.DataBodyRange, lcs.Range)
    r.FormulaLocal = myFormulaLocal
    r.NumberFormat = NumbFormat

End Sub
Sub ListObjectColumnFormulaR1C1(lcs As ListColumn, myFormulaR1C1, Optional NumbFormat)
    Dim lo     As ListObject
    Dim r      As Range
    Set lo = lcs.parent
    Set r = Intersect(lo.DataBodyRange, lcs.Range)
    r.FormulaR1C1 = myFormulaR1C1
    r.NumberFormat = NumbFormat

End Sub

Sub ListObjectColumnFormula(lcs As ListColumn, myFormula, Optional NumbFormat)
    Dim lo     As ListObject
    Dim r      As Range
    Set lo = lcs.parent
    Set r = Intersect(lo.DataBodyRange, lcs.Range)
    r.Formula = myFormula
    If Not IsMissing(NumbFormat) Then
    r.NumberFormat = NumbFormat
    Else
    r.NumberFormat = "General"
    End If

End Sub

Function ListObjectColumnAddFormulaOrValue(lcs As ListColumn, formul As String, Optional form_at As String, Optional v As Boolean = False) As ListColumn
'добавляет столбец с именем LCName в умную таблицу,вставляет формулу Formul,форматирует form_at,и при v As true превращает в значения
'    Dim lcs    As ListColumn
'    Set lcs = ListObjectColumnAdd(lcs.Parent, lcs.Name)
    Call ListObjectColumnFormula(lcs, formul, form_at)
    If v Then
        For Each r In lcs.Range.Cells
            r.value = r.value
        Next
    End If
    Set ListObjectColumnAddFormulaOrValue = lcs
End Function




Sub ActiveListObjectColumnDelete(lc As ListColumn)



'Добавляет новую строку в таблице, представленной указанным ListObject
'Положение-Определяет относительное положение новой строки
' Указывает, следует ли всегда переложить данные в ячейки ниже последней строки таблицы,
'когда вставляется новая строка, независимо от того, строка ниже таблице пуст.
'Если True , что ниже таблице клетки будут сдвинуты на одну строку вниз.
'Если False , если строка ниже таблице пуст, таблица будет расширяться, чтобы занять эту строку
'без смещения клеток под ним; но если строка ниже таблице содержатся данные,
'эти клетки будут сдвинуты вниз, когда вставляется новая строка.


    Dim wb     As Workbook
    Dim ws     As Worksheet
  
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
    
'   Set lsr = ws.ListObjects(ИмяТаблицы)
    If IsMissing(Положение) Then
        Set lsr = lo.ListColumns.Add
    Else
   z = ActiveListObjectIndexColumn(Положение)
        Set lsr = lo.ListRows.Add(Положение)
    End If
    Set ActiveListObjectColumnAdd = lsr
End Sub

Sub ListObjectDataBodyRangeDelete(lo As ListObject)
Dim ws     As Excel.Worksheet

Set ws = ActiveSheet
If Not Exists(lo.name) Then Exit Sub
lo.DataBodyRange.Delete

    

    
        
'          Set ActiveListObjectDataBodyRange = lo.DataBodyRange
        

End Sub



Sub ListObjectDataBodyClearContents(lo As ListObject)
Dim ws     As Excel.Worksheet

Set ws = ActiveSheet
If Not Exists(lo.name) Then Exit Sub
lo.DataBodyRange.ClearContents

    

    
        
'          Set ActiveListObjectDataBodyRange = lo.DataBodyRange
        

End Sub








Sub ActiveListObjectRowDelete()
'Удаляет активную строку умной таблицы


    Dim wb     As Workbook
    Dim ws     As Worksheet
  
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
    
   Set lsr = ActiveListObjectActiveRow
    If Not lsr Is Nothing Then
        lsr.Delete
   
    End If
   
End Sub
Function ActiveListObjectRowAdd(lo As ListObject, Optional Положение, Optional AlwaysInserta As Boolean = True) As ListRow



'Добавляет новую строку в таблице, представленной указанным ListObject
'Положение-Определяет относительное положение новой строки
' Указывает, следует ли всегда переложить данные в ячейки ниже последней строки таблицы,
'когда вставляется новая строка, независимо от того, строка ниже таблице пуст.
'Если True , что ниже таблице клетки будут сдвинуты на одну строку вниз.
'Если False , если строка ниже таблице пуст, таблица будет расширяться, чтобы занять эту строку
'без смещения клеток под ним; но если строка ниже таблице содержатся данные,
'эти клетки будут сдвинуты вниз, когда вставляется новая строка.


    Dim wb     As Workbook
    Dim ws     As Worksheet
  
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
    
'   Set lsr = ws.ListObjects(ИмяТаблицы)
    If IsMissing(Положение) Then
        Set lsr = lo.ListRows.Add
    Else
    
        Set lsr = lo.ListRows.Add(Положение, AlwaysInsert)
    End If
    Set ActiveListObjectRowAdd = lsr
End Function
Function ActiveListObjectIndexColumn(ИмяСтолбца)
    Dim wb     As Workbook
    Dim ws     As Worksheet
    Dim ls     As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim cc     As Range
    Dim rr     As Range
    Dim r      As Range
    
'    Set ls = ActiveListObject
    Set zagrow = ActiveListObjectHederRowRange
    ActiveListObjectIndexColumn = Application.WorksheetFunction.match(ИмяСтолбца, zagrow, 0)


End Function






Sub ActiveListObjectFiltr(ИмяСтолбца, Optional myCriteria As String)
    Dim wb     As Workbook
    Dim ws     As Worksheet
    Dim ls     As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim cc     As Range
    Dim rr     As Range
    Dim r      As Range
    Dim z      As Integer
    On Error GoTo ActiveListObjectFiltr_Error

'    Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
    Set ls = ActiveListObject
    z = ActiveListObjectIndexColumn(ИмяСтолбца)
    If Not myCriteria = "" Then
        ls.Range.AutoFilter Field:=z, Criteria1:= _
                            myCriteria
    Else
        ls.Range.AutoFilter Field:=z
    End If

    On Error GoTo 0
    Exit Sub

ActiveListObjectFiltr_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure ActiveListObjectFiltr of Module УмныеТаблицы"
End Sub

Sub ActiveListObjectFiltrActiveCell()
    Dim wb     As Workbook
    Dim ws     As Worksheet
    Dim ls     As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim cc     As Range
    Dim rr     As Range
    Dim r      As Range
    Dim z      As Integer
    Dim ИмяСтолбца
    Dim ИмяСтолбцаR As Range
    On Error GoTo ActiveListObjectFiltr_Error

'    Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
    Set ls = ActiveListObject
    Set lsc = ActiveListObjectActiveColumn
    Set ИмяСтолбцаR = lsc.Range.Cells(1)
  v = AutoFilter_Criteria(ИмяСтолбцаR)
  z = ActiveListObjectIndexColumn(ИмяСтолбцаR.value)
  
    If v = "" Then
    Call Sort(12232)
                            
    Else
        ls.Range.AutoFilter Field:=z
    End If

    On Error GoTo 0
    Exit Sub

ActiveListObjectFiltr_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure ActiveListObjectFiltr of Module УмныеТаблицы"
End Sub

Friend Function AutoFilter_Criteria(Header As Range) As String

Dim strCri1 As String, strCri2 As String



    Application.Volatile

    

    With Header.parent.AutoFilter

        With .Filters(Header.Column - .Range.Column + 1)

        

            If Not .On Then Exit Function

            

                strCri1 = .Criteria1

            If .Operator = xlAnd Then

                strCri2 = " AND " & .Criteria2

            ElseIf .Operator = xlOr Then

                strCri2 = " OR " & .Criteria2

            End If

            

        End With

    End With

        

    AutoFilter_Criteria = UCase(Header) & ": " & strCri1 & strCri2

End Function
Sub Макрос1()
    Dim rng As Range
 
    With Worksheets("Лист1") 'подставьте своё название листа или можно написать ActiveSheet (но это не универсально)
        '.Range("$A$1:$C$14").AutoFilter Field:=3, Criteria1:="=*персик*"
        Set rng = .Range("A1:C" & .Cells(Rows.Count, "C").End(xlUp).Row)    'присваиваем диапазон переменной Rng (от А1 до С... последняя строка)
        If Not .AutoFilterMode Then    'проверяем установлен Автофильтр на листе или нет
            rng.AutoFilter 'устанавливаем автофильтр на столбцы таблицы
        Else
            If .FilterMode = True Then .ShowAllData    'если Автофильтр применён к какому-либо столбцу, то снимаем все фильтры
        End If
 
        rng.AutoFilter Field:=3, Criteria1:="*персик*"    'через Автофильтр отфильтровываем таблицу по столбцу С
        If .AutoFilter.Range.Columns(1).SpecialCells(xlVisible).Count = 1 Then 'если нет отфильтрованных строк, кроме шапки таблице, то
            MsgBox "В столбце С нет слова 'персик'!", vbExclamation, "Ошибка"
            .ShowAllData 'снимаем установленные фильтр
            Exit Sub
        End If
 
        'Пример
        'переменной Rng присваиваем отфильтрованный диапазон
        With .AutoFilter.Range
            Set rng = .SpecialCells(xlCellTypeVisible) 'с шапкой таблицы
            'Set Rng = .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeVisible) 'без шапки таблицы
        End With
        With Worksheets("Лист2")
            .Cells.Clear
            rng.copy Destination:=.Range("A1") 'копируем отфильтрованные данные на Лист2
        End With
        'конец примера
         
        .AutoFilterMode = False 'отключаем Автофильтр
    End With
 
End Sub
'Friend Function ActiveListObjectActiveColumnNameRange()
'
'End Function





'Friend Function ActiveListObjectTotal()
'Dim ws     As Excel.Worksheet
'Dim lo     As Excel.ListObject
'If ActiveSheet.ListObjects.Count = 0 Then Exit Function
'
'
'    Set ws = ActiveSheet
'
'        For Each lo In ws.ListObjects
'        If Not Intersect(ActiveCell, lo.range) Is Nothing Then
'          Set ActiveListObjectHeder = lo.HeaderRowRange
'        End If
'        Next lo
'
'
'End Function

Sub Sort(Optional ID)
'Если ID:=12232 поменять на ID представленные ниже то будет сортировать по цвету.
'12948 - по цвету ячейки
'12234 - по цвету шрифта
If IsMissing(ID) Then ID = 12232

' Call the built-in filter option to filter on ACell
Application.CommandBars("Cell").FindControl(ID:=ID, Recursive:=True).Execute
End Sub
Function SortArray(mySortArray, Optional Index As Variant)
'
' Макрос1 Макрос
'

 Dim wb     As Workbook
    Dim ws     As Worksheet
    Dim ls     As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
   Dim rr      As Range
   If IsMissing(Index) Then
   Else
    Set ls = m_Tables(Index)
    End If
    
    If Not ls Is Nothing Then
    With ls
    .Sort.SortFields. _
        Clear
       
        If IsArray(mySortArray) Then
        For i = 0 To UBound(mySortArray)
    .Sort.SortFields.Add _
        key:=Range(.name & "[" & mySortArray(i) & "]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    Next i
    
    Else
     If mySortArray <> "" Then
    .Sort.SortFields.Add _
        key:=Range(.name & "[" & mySortArray & "]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
        End If
    End If
    
       With ls.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
    
    End With
    End If
End Function

Function formula_ПСТР_ListObject(locVal As ListColumn, start)
'
' Макрос3 Макрос
'

'
'    Range("P5").Select
'    ActiveCell.FormulaR1C1 = "=MID([@Столбец1],1,1)"
'    Range("J6").Select

formula_ПСТР_ListObject = "=MID(" & locVal.parent.name & "[@" & locVal.name & "]" & "," & start & ",1)"


'  Range("Таблица3[ID]").FormulaR1C1 = str
End Function








Function FormulaInderect(Optional lsc)

If IsMissing(lsc) Then Set lsc = ActiveListObjectActiveColumn()
FormulaInderect = "=INDIRECT(" & """" & lsc.parent.name & "[" & lsc.name & "]" & """" & ")"

End Function
Sub ValidationListAdd(r As Range, myFormula)
myIgnoreBlank = False '
myInCellDropdown = True '
myShowInput = True '
myShowError = False '
    
    With r.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=myFormula
        .IgnoreBlank = myIgnoreBlank
        .InCellDropdown = myInCellDropdown
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = myShowInput
        .ShowError = myShowError
    End With
End Sub

Function ValueListObject(Index As Variant, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска) As Range
'ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
'"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
' Можно использовать индексы
    Dim wb     As Workbook
    Dim ws     As Worksheet
    Dim ls     As ListObject
    Dim lsc    As ListColumn
    Dim lscf   As ListColumn
    Dim lsr    As ListRow
    Dim r      As Range
   Dim rr      As Range
    Set ls = m_Tables(Index)
    Set lsc = ls.ListColumns(ИмяСтолбцаПоиска)
    Set lscf = ls.ListColumns(ИмяСтолбцаЗначения)
    Set n = lsc.Range.Find(ЗначениеПоиска, LookIn:=xlValues, LookAt:=xlWhole)
    If Not n Is Nothing Then
    Set rr = Application.Intersect(ls.parent.Rows(n.Row), ls.DataBodyRange)
    
'        Set lsr = ls.ListRows(n.Row - 1)
        Set r = Application.Intersect(lscf.Range, rr)
        Set ValueListObject = r
    Else
        Set ValueListObject = Nothing
    End If
End Function
'Public Property Get ValueListObject() As Range
'    Set ValueListObject = m_ValueListObject
'End Property
Function ActiveCellValidationSourseRange()
s = ValidationType(ActiveCell)
If Not IsEmpty(s) Then
vf = ValidationFormula(ActiveCell)
If vf Like "*ДВССЫЛ*" Then
Set ActiveCellValidationSourseRange = Range(FormulaДВССЫЛSource())
'MsgBox r.Parent.Name
'r.Parent.Activate
'r.Select
End If
End If
End Function
Function FormulaДВССЫЛSource(Optional r)
Dim cl As clsVBScriptRegExp
Dim str As String
Set cl = New clsVBScriptRegExp
If IsMissing(r) Then Set lsc = ActiveListObjectColumnName()
If Not lsc Is Nothing Then
lapki = Chr(34)
cl.myPattern = "[^" & lapki & "]+"


cl.myGlobal = True
If IsMissing(r) Then Set r = ActiveCell
str = ValidationFormula(r)
Set objMatches = cl.objRegExpReplaceExecute(str)
Set objMatch = objMatches.Item(1)
    FormulaДВССЫЛSource = objMatch.value '& ", " & "FirstIndex=" & objMatch.FirstIndex & ", " & "Length=" & objMatch.Length

Else
FormulaДВССЫЛSource = ""
End If
End Function



Function Logik(value) As Boolean
Select Case value
Case "Да"
Logik = True
Case "Нет"
Logik = False
Case "True"
Logik = True
Case "False"
Logik = False
Case Else
End Select
End Function
Function ActiveListObjectComment(Optional Index As Variant, Optional strComment)




Dim mloj As ListObject

If IsMissing(Index) Then
Set mloj = ActiveListObject
Else
Set mloj = m_Tables(Index)
End If
If Not IsMissing(strComment) Then
mloj.Comment = strComment
End If
ActiveListObjectComment = mloj.Comment
End Function
Function MAXListTableColumn(ИмяЛиста, ИмяТаблицы, ИмяСтолбца)
'возвращает максимальное значение столбца умной таблицы

    MAXListTableColumn = Application.WorksheetFunction.max(СписокЗначенийСтолбцаУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбца))
End Function
Function LinkNameString(link As String, Вернуть As String) As Object
'link в виде Призвище[Ім'я]

'Вернуть-в виде:
'Англиская раскладка
'"t"-Возвращает  таблицу как обьект из примера:Призвище
'"c"-Возвращает столбец таблицы  как обьект из примера:Ім'я
'Проверка на наличие имени столбца и имени таблицы

Dim colName As String
Dim tbName
Dim exc As Range
spl = Split(link, "[")
tbName = spl(0)
spl = Split(spl(1), "]")
colName = spl(0)
'Проверка на наличие таблицы с именем tbName в книге wb
ext = Exists(tbName)
If Not ext Then
Set LinkNameString = Nothing
Exit Function
Else
Set exc = ExistColumn(tbName, colName)
If exc Is Nothing Then
Set LinkNameString = Nothing
Exit Function
End If

End If
Select Case Вернуть
Case "t"
Set LinkNameString = m_Tables(tbName)
'umt = tbName

Case "c"
'Проверка на наличие столбца с именем в таблице tbName в книге wb
Set LinkNameString = m_Tables(tbName).ListColumns(colName)
'umt = colName

End Select



End Function

Function SheetExistBookCreate(SheetName, Optional cl As Boolean = False) As Worksheet
'Определение есть ли в  книге "wb" лист с именем SheetName,если нет то создает его
Dim sh As Object
On Error Resume Next
Set sh = m_wb.Worksheets(SheetName)
If Err <> 0 Then
Set sh = m_wb.Worksheets.Add
sh.name = SheetName
Else
If cl Then sh.Cells.Clear
End If
Set SheetExistBookCreate = sh
End Function
Function CreateIDColumn(Optional Index As Variant, Optional НазваID, Optional mVisible = False) As ListColumn
Dim loj As clsmListObjs
Dim wb As Workbook
Dim lop As ListObject
Dim r As Range
Dim lcID As ListColumn
Dim rVisible As Range
Dim mName As String
'If IsMissing(zb) Then
'Set wb = ActiveWorkbook
'Else
'Set wb = zb
'End If
'
'Set loj = New clsmListObjs
'With loj
'.Initialize wb
   On Error GoTo CreateIDColumn_Error

If IsMissing(Index) Then
Set lop = ActiveListObject
Else
Set lop = m_Tables(Index)
End If
If IsMissing(НазваID) Then
mName = "ID"
Else
mName = НазваID
End If
Set z = ExistColumn(lop.name, mName)
 If z Is Nothing Then
Set lcID = ActiveListObjectColumnAdd(Index:=lop.name, name:=mName)
Else
Set lcID = lop.ListColumns(mName)
End If
k = 1

If mVisible Then
Set rVisible = lcID.DataBodyRange.SpecialCells(xlCellTypeVisible)
Else
Set rVisible = lcID.DataBodyRange
End If
 For Each r In rVisible
 
 r.value = k
 k = k + 1

 Next
 Set CreateIDColumn = lcID
' End With

   On Error GoTo 0
   Exit Function

CreateIDColumn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CreateIDColumn of Class Module clsmListObjs"
 
End Function
Function CreateAlphabetColumn(LinkAlphabet, StartAlphabet, Optional Index As Variant, Optional НазваAlphabet) As ListColumn
Dim loj As clsmListObjs
Dim wb As Workbook
Dim lop As ListObject
Dim r As Range
Dim lcAlphabet As ListColumn
Dim lcLinkAlphabet As ListColumn
Dim rVisible As Range
Dim mName As String
Dim formul As String
   On Error GoTo CreateAlphabetColumn_Error

If IsMissing(Index) Then
Set lop = ActiveListObject
Else
Set lop = m_Tables(Index)
End If
If IsMissing(НазваAlphabet) Then
mName = "Alphabet"
Else
mName = НазваAlphabet
End If
'Ссылка на столбец в котором будут содержать буквы
Set z = ExistColumn(lop.name, mName)
 If z Is Nothing Then
Set lcAlphabet = ActiveListObjectColumnAdd(Index:=lop.name, name:=mName)
Else

Set lcAlphabet = lop.ListColumns(mName)
End If
'Ссылка на столбец из которого будут браться буквы
'LinkAlphabet
mName = LinkAlphabet
Set z = ExistColumn(lop.name, mName)
 If z Is Nothing Then
Exit Function
Else
Set lcLinkAlphabet = lop.ListColumns(mName)
End If

'С какой буквы нач
'ПозицияАлфавит = loj.ValueListObject("Parametrs", "Parametr", "Value", "Алфавит Позиция")
formul = formula_ПСТР_ListObject(lcLinkAlphabet, StartAlphabet)

  Set lcAlphabet = ListObjectColumnAddFormulaOrValue(lcAlphabet, formul)


 Set CreateAlphabetColumn = lcAlphabet
' End With

   On Error GoTo 0
   Exit Function

CreateAlphabetColumn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CreateAlphabetColumn of Class Module clsmListObjs"
 
End Function


Function CreateSelectionColumn(Optional Index As Variant, Optional НазваSelection) As ListColumn
Dim loj As clsmListObjs
Dim wb As Workbook
Dim lop As ListObject
Dim r As Range
Dim lcSelection As ListColumn
Dim lcLinkAlphabet As ListColumn
Dim rVisible As Range
Dim mName As String
Dim formul As String
   On Error GoTo CreateSelectionColumn_Error

If IsMissing(Index) Then
Set lop = ActiveListObject
Else
Set lop = m_Tables(Index)
End If
If IsMissing(НазваSelection) Then
mName = "SL"
Else
mName = НазваSelection
End If
'Ссылка на столбец в котором будут происходить отбор
Set z = ExistColumn(lop.name, mName)
 If z Is Nothing Then
Set lcSelection = ActiveListObjectColumnAdd(Index:=lop.name, name:=mName)
Else

Set lcSelection = lop.ListColumns(mName)
End If


    With lcSelection.DataBodyRange.Font
        .name = "Marlett"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    lcSelection.DataBodyRange.Font.Bold = True
'    lcSelection.Range.AutoFit
    With lcSelection.DataBodyRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
   Set lohr = ActiveListObjectHederRowRange

Set r = Intersect(lohr, lcSelection.Range)
r.HorizontalAlignment = xlCenter



 Set CreateSelectionColumn = lcSelection
' End With

   On Error GoTo 0
   Exit Function

CreateSelectionColumn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure CreateSelectionColumn of Class Module clsmListObjs"
 
End Function

Sub ReversMarlettA(Optional r As Variant)
If IsMissing(r) Then Set r = ActiveCell
If r.Font.name = "Marlett" Then
If r.value = "a" Then
r.value = ""
Exit Sub
End If
If r.value = "" Then r.value = "a"
End If
End Sub
 Sub ListObjectList()
Dim v As ListObject
Dim hra
Dim shList As Worksheet
Dim lr As ListRow
Dim lop As ListObject
Dim k
Dim cell As Range
Set List = ValueListObject("Parametrs", "Parametr", "Value", "Оглавление для LO")

Set shList = SheetExistBookCreate(List.value)
shList.Cells.Clear
shList.Activate
shList.Cells(1).Activate
Set lop = AddListObject(List.value)
hra = Array("№", "Наименование", "Лист", "Заголовок", "Итоги", "Фильтр", "Строк", "Комментарий")

shList.Range(shList.Cells(1, 1), shList.Cells(1, UBound(hra) + 1)).value = hra
k = 1
Set lr = lop.ListRows(k)
For Each v In m_Tables
lr.Range.Cells(2).value = v.name
Set cell = lr.Range.Cells(2)
        shList.Hyperlinks.Add anchor:=cell, Address:="", _
     SubAddress:=v.name & "!A1"

lr.Range.Cells(3).value = v.parent.name



lr.Range.Cells(4).value = v.ShowHeaders
lr.Range.Cells(5).value = v.ShowTotals
lr.Range.Cells(6).value = v.ShowAutoFilter
lr.Range.Cells(7).value = v.ListRows.Count
lr.Range.Cells(8).value = v.Comment
Set lr = lop.ListRows.Add
'shList.Cells(k, 1)
k = k + 1
Next
lop.ListRows(k).Delete
 End Sub


Sub UnpivotListObject(Optional Index As Variant, Optional IndexAdd As Variant)
If IsMissing(Index) Then
Set lop = ActiveListObject
Else
Set lop = m_Tables(Index)
End If
If IsMissing(IndexAdd) Then
'возвращаем масив

Else
'Добавляем в таблицу
'возвращаем масив


End If
End Sub
Sub DeleteNListRowEnd(lop, DeleteEndRow)
'----------------------------------------------------------------------------
'Удалить с конца строк DeleteEndRow
 
For i = 1 To DeleteEndRow
lop.ListRows(lop.ListRows.Count).Delete
Next i
'------------------------------------------------------
End Sub
  Sub TransposeListObject(ListObjectVertikalas, ListObjectGorizontal, NameParamColumn As String)
  Dim lopV As ListObject
  Dim lopG As ListObject
  Dim locV As ListColumn
  Dim locG As ListColumn
  Dim r As Range
  Dim rFind As Range
  Dim rHederG As Range
  Set lopV = m_Tables(ListObjectVertikalas)
  Set lopG = m_Tables(ListObjectGorizontal)
  Set locV = lopV.ListColumns(NameParamColumn)
  Set rHederG = lopG.HeaderRowRange
  For Each r In locV.DataBodyRange.Cells
 Set rFind = rHederG.Find(r.value, , LookIn:=xlValues, LookAt:=xlWhole)
 If rFind Is Nothing Then
 Set locG = lopG.ListColumns.Add(r.Row)
 locG.name = r.value
 End If
  Next
  
  End Sub
