Attribute VB_Name = "UDFs"
'работает аналогично ВПР, но возвращает массив данных
Function VLOOKUP3(Table As Range, SearchColumnNum As Integer, _
SearchValue As Variant, ResultColumnNum As Integer)
    
    Dim I, j As Integer
    Dim out(1000) As Variant
    Dim rCol As Range
    j = 0
        For I = 1 To Table.Rows.Count
            If Table.Cells(I, SearchColumnNum) = SearchValue Then
                out(j) = Table.Cells(I, ResultColumnNum)
                j = j + 1
            End If
        Next I
    VLOOKUP3 = Application.Transpose(out)
End Function

'Транслитерация русского текста в английский
Function Translit(txt As String) As String
    Dim Rus As Variant
    Rus = Array("а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я", "А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я")
    Dim Eng As Variant
    Eng = Array("a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "sch", "''", "y", "'", "e", "ju", "ja", "A", "B", "V", "G", "D", "E", "JO", "ZH", "Z", "I", "J", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "KH", "TS", "CH", "SH", "SCH", "''", "Y", "'", "E", "JU", "JA")
    
    For I = 1 To Len(txt)
        с = Mid(txt, I, 1)
    
        flag = 0
        For j = 0 To 64
            If Rus(j) = с Then
                outchr = Eng(j)
                flag = 1
                Exit For
            End If
        Next j
        If flag Then outstr = outstr & outchr Else outstr = outstr & с
    Next I
    
    Translit = outstr
    
End Function

'слияние текста всех ячеек диапазона с разделителем
Function MultiCat(ByRef Rng As Excel.Range, Optional ByVal delim As String = "") As String
     Dim rCell As Range
     For Each rCell In Rng
         MultiCat = MultiCat & delim & rCell.Text
     Next rCell
     MultiCat = Mid(MultiCat, Len(delim) + 1)
  End Function

'вывод заднного количества неповторяющихся случайных чисел из диапазона
Function Lotto(Bottom As Integer, Top As Integer, Amount As Integer)
    Dim iArr As Variant
    Dim I As Integer
    Dim r As Integer
    Dim Temp As Integer
    Dim out(1000) As Variant
    
    Application.Volatile
    
    ReDim iArr(Bottom To Top)
    For I = Bottom To Top
        iArr(I) = I
    Next I
    
    For I = Top To Bottom + 1 Step -1
        r = Int(Rnd() * (I - Bottom + 1)) + Bottom
        Temp = iArr(r)
        iArr(r) = iArr(I)
        iArr(I) = Temp
    Next I
    j = 0
    For I = Bottom To Bottom + Amount - 1
        out(j) = iArr(I)
        j = j + 1
    Next I
    
    Lotto = Application.Transpose(out)
    
End Function
'выбор случайного элемента из диапазона
Function RandomSelect(TargetCells)
    RandomSelect = TargetCells.Cells(Int(Rnd * TargetCells.Count) + 1)
End Function

'вывод дня недели по дате словом
Function WeekdayWord(MyDate As Date) As String
    Dim days As Variant
    days = Array("понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье")
    WeekdayWord = days(Weekday(MyDate, vbMonday) - 1)
End Function

'выводит любой заданный разряд числа
Function Class(m, I)
       Class = Int(Int(m - (10 ^ I) * Int(m / (10 ^ I))) / 10 ^ (I - 1))
End Function

'сумма прописью на русском языке
Function PropisRus(n As Double, rub As Boolean) As String
    Dim Nums1, Nums2, Nums3, Nums4 As Variant
    Nums1 = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
    Nums2 = Array("", "десять ", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
    Nums3 = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
    Nums4 = Array("", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
    Nums5 = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
    
    If n <= 0 Then
        Propis = "ноль"
        Exit Function
    End If
    ed = Class(n, 1)
    dec = Class(n, 2)
    sot = Class(n, 3)
    tys = Class(n, 4)
    dectys = Class(n, 5)
    sottys = Class(n, 6)
    mil = Class(n, 7)
    decmil = Class(n, 8)
    
    Select Case decmil
        Case 1
            mil_txt = Nums5(mil) & "миллионов "
            GoTo www
        Case 2 To 9
            decmil_txt = Nums2(decmil)
    End Select
    
    Select Case mil
        Case 1
            mil_txt = Nums1(mil) & "миллион "
        Case 2, 3, 4
            mil_txt = Nums1(mil) & "миллиона "
        Case 5 To 20
            mil_txt = Nums1(mil) & "миллионов "
    End Select
www:
    sottys_txt = Nums3(sottys)
    Select Case dectys
        Case 1
            tys_txt = Nums5(tys) & "тысяч "
            GoTo eee
        Case 2 To 9
            dectys_txt = Nums2(dectys)
    End Select
    
    Select Case tys
        Case 0
            If dectys > 0 Then tys_txt = Nums4(tys) & "тысяч "
        Case 1
            tys_txt = Nums4(tys) & "тысячa "
        Case 2, 3, 4
            tys_txt = Nums4(tys) & "тысячи "
        Case 5 To 9
            tys_txt = Nums4(tys) & "тысяч "
    End Select
    If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & " тысяч "
eee:
    sot_txt = Nums3(sot)
    
    Select Case dec
    Case 1
        ed_txt = Nums5(ed)
        GoTo rrr
    Case 2 To 9
        dec_txt = Nums2(dec)
    End Select
    
    ed_txt = Nums1(ed)
rrr:
    If rub Then
        Select Case ed_txt
            Case "один "
                rub_txt = "рубль"
            Case "два ", "три ", "четыре "
                rub_txt = "рубля"
            Case Else
                rub_txt = "рублей"
        End Select
        kops = Round((n * 100 - Int(n) * 100), 0)
        PropisRus = decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt & rub_txt & " " & kops & " коп."
    Else
        PropisRus = decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
    End If
End Function

'сумма прописью на английском языке
Function PropisEng(ByVal strAmount As String, strCur As String, strDec As String, iPrec As Integer)
    Dim BigDenom As String, SmallDenom As String, Temp As String
    Dim iDecimalPlace As Integer
    Dim Count As Integer
    
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    
    ' String representation of amount.
    strAmount = Trim(str(strAmount))
    
    ' Position of decimal place 0 if none.
    iDecimalPlace = InStr(strAmount, ".")
    
    ' Separate the Integer part from the decimals.
    If iDecimalPlace > 0 Then
        SmallDenom = Left(Right(strAmount, Len(strAmount) - iDecimalPlace) & "0000000000", iPrec)
        SmallDenom = PropisEng(SmallDenom, strDec, "", 0)
        BigDenom = Left(strAmount, iDecimalPlace - 1)
        BigDenom = PropisEng(BigDenom, strCur, "", 0)
        PropisEng = BigDenom & " And " & SmallDenom
        Exit Function
    End If
    If iDecimalPlace = 0 Then
        Count = 1
        Do While strAmount <> ""
            Temp = GetHundreds(Right(strAmount, 3))
            If Temp <> "" Then BigDenom = Temp & Place(Count) & BigDenom
            If Len(strAmount) > 3 Then
                strAmount = Left(strAmount, Len(strAmount) - 3)
            Else
                strAmount = ""
            End If
            Count = Count + 1
        Loop
        Select Case BigDenom
            Case ""
                BigDenom = "No " & strCur
            Case "One"
                BigDenom = "One " & strCur
             Case Else
                BigDenom = BigDenom & " " & strCur
        End Select
        PropisEng = BigDenom
    End If
End Function

' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        result = result & GetTens(Mid(MyNumber, 2))
    Else
        result = result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = result
End Function

' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim result As String
    result = ""           ' Null out the temporary function value."
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19…
        Select Case Val(TensText)
            Case 10: result = "Ten"
            Case 11: result = "Eleven"
            Case 12: result = "Twelve"
            Case 13: result = "Thirteen"
            Case 14: result = "Fourteen"
            Case 15: result = "Fifteen"
            Case 16: result = "Sixteen"
            Case 17: result = "Seventeen"
            Case 18: result = "Eighteen"
            Case 19: result = "Nineteen"
            Case Else
        End Select
    Else                                 ' If value between 20-99…
        Select Case Val(Left(TensText, 1))
            Case 2: result = "Twenty "
            Case 3: result = "Thirty "
            Case 4: result = "Forty "
            Case 5: result = "Fifty "
            Case 6: result = "Sixty "
            Case 7: result = "Seventy "
            Case 8: result = "Eighty "
            Case 9: result = "Ninety "
            Case Else
        End Select
        result = result & GetDigit _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = result
End Function

' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

'суммирует ячейки в заданном интервале
Function SumBetween(TargetCells As Range, min As Long, max As Long, IncludeMin As Boolean, IncludeMax As Boolean) As Long
    Dim s As Long
    For Each C In TargetCells
        If IncludeMin And IncludeMax = True Then If C >= min And C <= max Then s = s + C
        If IncludeMin And Not IncludeMax Then If C >= min And C < max Then s = s + C
        If Not IncludeMin And IncludeMax Then If C > min And C <= max Then s = s + C
        If Not IncludeMin And Not IncludeMax Then If C > min And C < max Then s = s + C
    Next C
    SumBetween = s
End Function

'возвращает дату N-го дня недели (W) для заданного месяца М и года Y
Function NeedDate(n As Integer, W As Integer, m As Integer, Y As Integer) As Date
    Dim I, md As Integer
    Dim D As Date
    'определяем сколько дней в месяце
    Select Case m
    Case 1, 3, 5, 7, 8, 10, 12
        md = 31
    Case 4, 6, 9, 11
        md = 30
    Case 2
        'если год високосный, то в феврале 29 иначе 28 дней
        If (Y - 2000) Mod 4 = 0 Then md = 29 Else md = 28
    End Select
    
    For D = DateSerial(Y, m, 1) To DateSerial(Y, m, md)
        If Weekday(D, vbMonday) = W Then
            I = I + 1
            If I = n Then
                NeedDate = D
                Exit Function
            End If
        End If
    Next D
    NeedDate = " "
End Function

'выделяет числа из ячейки
Function GetNumbers(TargetCell As Range) As String
    Dim LenStr As Long
    For LenStr = 1 To Len(TargetCell)
        Select Case Asc(Mid(TargetCell, LenStr, 1))
        Case 48 To 57
            GetNumbers = GetNumbers & Mid(TargetCell, LenStr, 1)
        End Select
    Next
End Function

'выделяет текст из ячейки
Function GetText(TargetCell As Range) As String
    Dim LenStr As Long
    For LenStr = 1 To Len(TargetCell)
        Select Case Asc(Mid(TargetCell, LenStr, 1))
        Case 65 To 90
            GetText = GetText & Mid(TargetCell, LenStr, 1)
        Case 97 To 122
            GetText = GetText & Mid(TargetCell, LenStr, 1)
        Case 192 To 255
            GetText = GetText & Mid(TargetCell, LenStr, 1)
        End Select
    Next
End Function

'возвращает первое значение в указанной строке
Function FirstInRow(myRow As Range)
    If Cells(myRow.Row, 1) <> "" Then FirstInRow = Cells(myRow.Row, 1).value
    If Cells(myRow.Row, 1) = "" Then FirstInRow = Cells(myRow.Row, 1).End(xlToRight).value
End Function

'возвращает первое значение в указанном столбце
Function FirstInColumn(myColumn As Range)
Attribute FirstInColumn.VB_Description = "Возвращает значение первой непустой ячейки в заданном столбце."
Attribute FirstInColumn.VB_ProcData.VB_Invoke_Func = " \n14"
    If Cells(1, myColumn.Column) <> "" Then FirstInColumn = Cells(1, myColumn.Column).value
    If Cells(1, myColumn.Column) = "" Then FirstInColumn = Cells(1, myColumn.Column).End(xlDown).value
End Function

'возвращает последнее значение в указанной строке
Function LastInRow(myRow As Range)
    If Cells(myRow.Row, Sheets(1).Columns.Count) <> "" Then LastInRow = Cells(myRow.Row, Sheets(1).Columns.Count).value
    If Cells(myRow.Row, Sheets(1).Columns.Count) = "" Then LastInRow = Cells(myRow.Row, Sheets(1).Columns.Count).End(xlToLeft).value
End Function

'возвращает последнее значение в указанном столбце
Function LastInColumn(myColumn As Range)
    If Cells(Sheets(1).Rows.Count, myColumn.Column) <> "" Then LastInColumn = Cells(Sheets(1).Rows.Count, myColumn.Column).value
    If Cells(Sheets(1).Rows.Count, myColumn.Column) = "" Then LastInColumn = Cells(Sheets(1).Rows.Count, myColumn.Column).End(xlUp).value
End Function

'возвращает имя листа
Function SheetName1() As String
    SheetName = ActiveSheet.name
End Function

'возвращает имя книги
Function WorkbookName() As String
    WorkbookName = ActiveWorkbook.name
End Function

'возвращает полное имя файла (полный путь)
Function FullFileName() As String
    FullFileName = ActiveWorkbook.FullName
End Function

'возвращает имя текущего пользователя
Function UserName() As String
    UserName = Application.UserName
End Function

'код цвета заливки ячейки
Function CellColor(cell As Range)
Attribute CellColor.VB_Description = "Возвращает код цвета заливки ячейки"
Attribute CellColor.VB_ProcData.VB_Invoke_Func = " \n14"
    CellColor = cell.Interior.ColorIndex
End Function

'код цвета шрифта ячейки
Function CellFontColor(cell As Range)
    CellFontColor = cell.Font.ColorIndex
End Function


'выводит текущие условия автофильтра
Function AutoFilter_Criteria(Header As Range) As String
Dim strCri1 As String, strCri2 As String
    Application.Volatile
    With Header.Parent.AutoFilter
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

'выделяет подстроку из строки
Public Function Substring(txt, Delimiter, n) As String
Dim X As Variant
    X = Split(txt, Delimiter)
    If n > 0 And n - 1 <= UBound(X) Then
        Substring = X(n - 1)
    Else
        Substring = ""
    End If
End Function

'усовершенствованная версия ВПР
Function VLOOKUP2(Table As Range, SearchColumnNum As Integer, SearchValue As Variant, n As Integer, ResultColumnNum As Integer)

    Dim I As Integer
    Dim iCount As Integer
    Dim rCol As Range

        For I = 1 To Table.Rows.Count
            If Table.Cells(I, SearchColumnNum) = SearchValue Then
                iCount = iCount + 1
            End If

            If iCount = n Then
                VLOOKUP2 = Table.Cells(I, ResultColumnNum)
                Exit For
            End If
        Next I
End Function

'Проверка текста по шаблону
Function MaskCompare(txt As String, mask As String, CaseSensitive As Boolean)
    If Not CaseSensitive Then
        txt = UCase(txt)
        mask = UCase(mask)
    End If
        
    If txt Like mask Then
            MaskCompare = True
        Else
            MaskCompare = False
    End If
End Function

'подсчитывает количество ячеек в диапазоне, удовлетворяющих маске
Function CountByMask(Rng As Range, mask As String, CaseSensitive As Boolean)

    For Each C In Rng
        If Not CaseSensitive Then
            txt = UCase(C)
            mask = UCase(mask)
        Else
            txt = с
        End If
        If txt Like mask Then n = n + 1
    Next C
    CountByMask = n
End Function


'Проверка наличия в тексте символов латиницы
Function IsLatin(txt As String)
    txt = UCase(txt)
    mask = "*[ABCDEFGHIJKLMNOPQRSTUVWXYZ]*"
        
    If txt Like mask Then
            IsLatin = True
        Else
            IsLatin = False
    End If
End Function

'Сумма ячеек с определенным цветом заливки
Function SumByCellColor(SearchRange As Range, TargetCell As Range)
Application.Volatile True

Sum = 0

For Each cell In SearchRange
    If cell.Interior.ColorIndex = TargetCell.Interior.ColorIndex Then
        Sum = Sum + cell.value
    End If
Next
SumByCellColor = Sum
End Function

'Сумма ячеек с определенным цветом шрифта
Function SumByFontColor(SearchRange As Range, TargetCell As Range)
Application.Volatile True

Sum = 0

For Each cell In SearchRange
    If cell.Font.ColorIndex = TargetCell.Font.ColorIndex Then
        Sum = Sum + cell.value
    End If
Next
SumByFontColor = Sum
End Function

'Построение микрографиков
Function MicroCharts(Rng As Range)
    Dim ChrCodes() As Integer
    Dim outstr As String

    ReDim ChrCodes(Rng.Count)
    minval = Application.min(Rng)
    minpos = Application.Match(minval, Rng, 0)
    maxval = Application.max(Rng)
    maxpos = Application.Match(maxval, Rng, 0)

    If minval = 0 And maxval = 0 Then   'все нулевые значения
        For Each C In Rng
            ChrCodes(I) = 33
            I = I + 1
        Next C
        GoTo theend
    End If
    If minval >= 0 Then  'только положительные числа
        I = 0
        For Each C In Rng
            ChrCodes(I) = 68 + Round(C.value / maxval * 21)
            I = I + 1
        Next C
        GoTo theend
    End If

    If maxval <= 0 Then    ' только отрицательные числа
        I = 0
        For Each C In Rng
            ChrCodes(I) = 90 + Round(C.value / minval * 20)
            I = I + 1
        Next C
        GoTo theend
    End If

    If maxval > 0 And minval < 0 Then    'положительные и отрицательные вместе
        I = 0
        For Each C In Rng
            If C.value > 0 Then
                ChrCodes(I) = 33 + Round(C.value / maxval * 15)
            End If
            If C.value < 0 Then
                ChrCodes(I) = 50 + Round(C.value / minval * 16)
            End If
            If C.value = 0 Then ChrCodes(I) = 33
            I = I + 1
        Next C
    End If

theend:
    'формируем и выводим готовый массив символов
    For j = 0 To UBound(ChrCodes)
        outstr = outstr & Chr(ChrCodes(j))
    Next j

    MicroCharts = outstr
End Function

Function LastRow(SheetName As String) As Long

'Определение последней используемой строки на листе с именем SheetName
Dim sh As Worksheet
Set sh = Worksheets(SheetName)
LastRow = sh.UsedRange.Rows.Count
LastRow = LastRow + sh.UsedRange.Row - 1
End Function
Function LastColumn(SheetName As String, r As Long) As Range

'Определение последней используемой ячейки в строке r на листе с именем SheetName
Dim sh As Worksheet
Dim EndCell As Range
Set sh = Worksheets(SheetName)
Set EndCell = sh.Cells(r, 256)
Set LastColumn = EndCell.End(xlToLeft)
End Function
Function SheetExist(SheetName As String) As Boolean
'Определение есть ли в активной книге лист с именем SheetName
Dim sh As Object
On Error Resume Next
Set sh = ActiveWorkbook.Worksheets(SheetName)
If Err = 0 Then SheetExist = True _
Else SheetExist = False
End Function
Function SheetExistBook(wb As Workbook, SheetName As String) As Boolean
'Определение есть ли в  книге "wb" лист с именем SheetName
Dim sh As Object
On Error Resume Next
Set sh = wb.Worksheets(SheetName)
If Err = 0 Then SheetExistBook = True _
Else SheetExistBook = False
End Function
Function SheetExistBookCreate(wb As Workbook, SheetName As String, cl As Boolean) As Worksheet
'Определение есть ли в  книге "wb" лист с именем SheetName,если нет то создает его
Dim sh As Object
On Error Resume Next
Set sh = wb.Worksheets(SheetName)
If Err <> 0 Then
Set sh = wb.Worksheets.Add
sh.name = SheetName
Else
If cl Then sh.Cells.Clear
End If
Set SheetExistBookCreate = sh
End Function



Function InversiaValue(v As Range)
InversiaValue = Val(Trim(v.value)) * (-1)
End Function
Function Delimeter_Count(r As Range, Delimeter As String)
'Количество Delimeter разделителей в тексте
k = 0
For I = 1 To Len(r.value)
s = Mid(r.value, I, 1)
If s = Delimeter Then k = k + 1
Next I
Delimeter_Count = k
End Function
Public Function SelectFiles(MultiSelect As Boolean, fname As String, f As String)
Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
.InitialView = msoFileDialogViewList
.AllowMultiSelect = MultiSelect
.Filters.Clear
.Filters.Add fname, f
If .Show = -1 Then
Set SelectFiles = .SelectedItems
Else
Set SelectFiles = Nothing
End If
End With
End Function
Function ИмяЛистаВАпостров(Имя As String)
If Имя Like "* *" Then ИмяЛистаВАпостров = "'" & Имя & "'"
End Function
 Function PathExists(pname) As Boolean
' ??g??????? ??????, ???? ???? ??????????
    Dim X As String
    On Error Resume Next
    X = GetAttr(pname) And 0
    If Err = 0 Then PathExists = True _
      Else PathExists = False
End Function
'Функция подсчета активных книг
Function mywbBook(name As String, pathbook As String)

     Dim lCount As Long, wbBook As Workbook
     
   If PathExists(pathbook) Then
     For Each wbBook In Application.Workbooks
         If wbBook.name = name Then
         Set mywbBook = wbBook
         Exit Function
         
         End If
       
              Next wbBook
            Set mywbBook = Workbooks.Open(pathbook & "\" & name)
            
     Else
     Set mywbBook = Nothing
     End If
     
 End Function
Function ПоискСоСмещением(m As Range, strFind, smr, smc)
On Error Resume Next

Set r = FindAll(m, strFind)
If r.Cells.Count > 1 Then Exit Function

Set ПоискСоСмещением = r.Offset(rowoffset:=smr, columnoffset:=smc)
End Function
