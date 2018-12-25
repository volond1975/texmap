Attribute VB_Name = "macroEXCEL4"
'GET.FORMULA () XLM
Function getFormula(r As Range)
getFormula = r.Formula
End Function
Function getFormulaArray(r As Range)
getFormulaArray = r.FormulaArray
End Function
Function getFormulaR1C1(r As Range)
getFormulaR1C1 = r.FormulaR1C1
End Function
Function getFormulaLocal(r As Range)
getFormulaLocal = r.FormulaLocal
End Function
Function getFormulaR1C1Local(r As Range)
getFormulaR1C1Local = r.FormulaR1C1Local
End Function
Function getNamedRangeByName(name As String) As name
Dim twb As Workbook
Set twb = ThisWorkbook
Set getNamedRangeByName = twb.Names(name)
End Function

Function getReferToNamedRangeByName(name As String) As String

Set rname = getNamedRangeByName(name)
getReferToNamedRangeByName = rname.RefersTo
End Function
Function getNamedRangeByAddress(addr As String)
Dim twb As Workbook
Dim r As Range
Set r = Range(addr)
getNamedRangeByAddress = r.address(External:=True)
End Function
Function Intersect_Name(Optional addr)
    Dim nn As name, rRange As Range
    Set r = Range(addr)
    For Each nn In ActiveWorkbook.Names
        On Error Resume Next: Set rRange = nn.RefersToRange
        On Error GoTo 0
        If Not rRange Is Nothing Then
        
            If rRange.Parent.name = r.Parent.name Then
            
                If Not Intersect(rRange, r) Is Nothing Then
                Intersect_Name = nn.name
                Exit Function
                End If
            End If
            Set rRange = Nothing
        End If
    Next
'   // MsgBox "?????????? ?????? ?? ?????? ?? ? ???? ??????????? ????????", vbInformation, "Hay_from_The_Prist"
End Function
Sub addNamedFormula()

Application.Names.Add ActiveCell.Offset(0, -7).value, ActiveCell.value
'MsgBox Evaluate("x-2")
'Application.Names.Add name = ActiveCell.Offset(0, -7).value, _
                      RefersTo:="""" & ActiveCell.value & """"
End Sub

Function CellAddress()
CellAddress = Application.ThisCell.address
End Function
Function CellParentName()
CellParentName = Application.ThisCell.Parent.name

End Function
Function REGPLACE(myRangeValue As String, matchPattern As String, outputPattern As String) As Variant
    Dim regex As New VBScript_RegExp_55.RegExp
    Dim strInput As String

    strInput = myRangeValue

    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = matchPattern
    End With

    REGPLACE = regex.Replace(strInput, outputPattern)

End Function
Function REGPLACEFORMULA(myRangeFormula As String, matchPattern As String, outputPattern As String) As Variant
    Dim regex As New VBScript_RegExp_55.RegExp
    Dim strInput As String

 strInput = myRangeFormula

    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = matchPattern
    End With

    REGPLACEFORMULA = regex.Replace(strInput, outputPattern)

End Function
'http://qaru.site/questions/19369/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops


Function regex(strInput As String, matchPattern As String, Optional ByVal outputPattern As String = "$0") As Variant
    Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
    Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
    Dim replaceNumber As Integer

    With inputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = matchPattern
    End With
    With outputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "\$(\d+)"
    End With
    With outReplaceRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With

    Set inputMatches = inputRegexObj.Execute(strInput)
    If inputMatches.Count = 0 Then
        regex = False
    Else
        Set replaceMatches = outputRegexObj.Execute(outputPattern)
        For Each replaceMatch In replaceMatches
            replaceNumber = replaceMatch.SubMatches(0)
            outReplaceRegexObj.pattern = "\$" & replaceNumber

            If replaceNumber = 0 Then
                outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).value)
            Else
                If replaceNumber > inputMatches(0).SubMatches.Count Then
                    'regex = "A to high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
                    regex = CVErr(xlErrValue)
                    Exit Function
                Else
                    outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
                End If
            End If
        Next
        regex = outputPattern
    End If
End Function
Public Function CellReflist(Optional r As Range)  ' single cell
Dim result As Object: Dim testExpression As String: Dim objRegEx As Object
If r Is Nothing Then Set r = ActiveCell ' Cells(1, 2)  ' INPUT THE CELL HERE , e.g.    RANGE("A1")
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.IgnoreCase = True: objRegEx.Global = True: objRegEx.pattern = """.*?"""    ' remove expressions
testExpression = CStr(r.Formula)
testExpression = objRegEx.Replace(testExpression, "")
'objRegEx.Pattern = "(([A-Z])+(\d)+)"  'grab the address

objRegEx.pattern = "(['\[].*?['!])?([[A-Z0-9_]+[!])?(\$?[A-Z]+\$?(\d)+(:\$?[A-Z]+\$?(\d)+)?|\$?[A-Z]+:\$?[A-Z]+|(\$?[A-Z]+\$?(\d)+))"
If objRegEx.Test(testExpression) Then
    Set result = objRegEx.Execute(testExpression)
    If result.Count > 0 Then CellReflist = result(0).value
    If result.Count > 1 Then
        For I = 1 To result.Count - 1 'Each Match In result
            dbl = False ' poistetaan tuplaesiintymiset
            For j = 0 To I - 1
                If result(I).value = result(j).value Then dbl = True
            Next j
            If Not dbl Then CellReflist = CellReflist & "," & result(I).value 'Match.Value
        Next I 'Match
    End If
End If
End Function

Sub ReturnFormulaReferences()

    Dim objRegExp As New VBScript_RegExp_55.RegExp
    Dim objCell As Range
    Dim objStringMatches As Object
    Dim objReferenceMatches As Object
    Dim objMatch As Object
    Dim intReferenceCount As Integer
    Dim intIndex As Integer
    Dim booIsReference As Boolean
    Dim objName As name
    Dim booNameFound As Boolean

    With objRegExp
        .MultiLine = True
        .Global = True
        .IgnoreCase = True
    End With

    For Each objCell In Selection.Cells
        If Left(objCell.Formula, 1) = "=" Then

            objRegExp.pattern = "\"".*\"""
            Set objStringMatches = objRegExp.Execute(objCell.Formula)

            objRegExp.pattern = "(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
            & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)?" _
            & "(\$?[a-z]{1,3}\$?[0-9]{1,7}(\:\$?[a-z]{1,3}\$?[0-9]{1,7})?" _
            & "|\$[a-z]{1,3}\:\$[a-z]{1,3}" _
            & "|[a-z]{1,3}\:[a-z]{1,3}" _
            & "|\$[0-9]{1,7}\:\$[0-9]{1,7}" _
            & "|[0-9]{1,7}\:[0-9]{1,7}" _
            & "|[a-z_\\][a-z0-9_\.]{0,254})"
            Set objReferenceMatches = objRegExp.Execute(objCell.Formula)

            intReferenceCount = 0
            For Each objMatch In objReferenceMatches
                intReferenceCount = intReferenceCount + 1
            Next

            Debug.Print objCell.Formula
            For intIndex = intReferenceCount - 1 To 0 Step -1
                booIsReference = True
                For Each objMatch In objStringMatches
                    If objReferenceMatches(intIndex).FirstIndex > objMatch.FirstIndex _
                    And objReferenceMatches(intIndex).FirstIndex < objMatch.FirstIndex + objMatch.Length Then
                        booIsReference = False
                        Exit For
                    End If
                Next

                If booIsReference Then
                    objRegExp.pattern = "(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
                    & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)?" _
                    & "(\$?[a-z]{1,3}\$?[0-9]{1,7}(\:\$?[a-z]{1,3}\$?[0-9]{1,7})?" _
                    & "|\$[a-z]{1,3}\:\$[a-z]{1,3}" _
                    & "|[a-z]{1,3}\:[a-z]{1,3}" _
                    & "|\$[0-9]{1,7}\:\$[0-9]{1,7}" _
                    & "|[0-9]{1,7}\:[0-9]{1,7})"
                    If Not objRegExp.Test(objReferenceMatches(intIndex).value) Then 'reference is not A1
                        objRegExp.pattern = "^(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
                        & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)" _
                        & "[a-z_\\][a-z0-9_\.]{0,254}$"
                        If Not objRegExp.Test(objReferenceMatches(intIndex).value) Then 'name is not external
                            booNameFound = False
                            For Each objName In objCell.Worksheet.Parent.Names
                                If objReferenceMatches(intIndex).value = objName.name Then
                                    booNameFound = True
                                    Exit For
                                End If
                            Next
                            If Not booNameFound Then
                                objRegExp.pattern = "^(\'.*(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\'\!" _
                                & "|(\[.*\])?([^\:\\\/\?\*\[\]]{1,31}\:)?[^\:\\\/\?\*\[\]]{1,31}\!)"
                                For Each objName In objCell.Worksheet.Names
                                    If objReferenceMatches(intIndex).value = objRegExp.Replace(objName.name, "") Then
                                        booNameFound = True
                                        Exit For
                                    End If
                                Next
                            End If
                            booIsReference = booNameFound
                        End If
                    End If
                End If

                If booIsReference Then
                    Debug.Print "  " & objReferenceMatches(intIndex).value _
                    & " (" & objReferenceMatches(intIndex).FirstIndex & ", " _
                    & objReferenceMatches(intIndex).Length & ")"
                End If
            Next intIndex
            Debug.Print

        End If
    Next

    Set objRegExp = Nothing
    Set objStringMatches = Nothing
    Set objReferenceMatches = Nothing
    Set objMatch = Nothing
    Set objCell = Nothing
    Set objName = Nothing

End Sub
 Sub replaceAllAC()
 Dim r As Range
 Dim f As String
 For Each r In Selection.Cells
 Call replaceAll(r)
 f = getFormula(r)
 Debug.Print f
r.Formula = REGPLACE(f, "1.215", "gККЗ")

 Next

 End Sub
 
Sub replaceAll(Optional r As Range)
If r Is Nothing Then Set r = ActiveCell
Dim rStr As String
Dim v As Variant
Dim strInput As String
Dim testExpression As String
Dim newz
Dim z As String
testExpression = CStr(r.Formula)
 rStr = CellReflist(r)
  v = Split(rStr, ",")
 For I = 0 To UBound(v)
 strInput = v(I)
 z = Val(regex(strInput, "\d+")) + 0
 leter = regex(strInput, "[A-Z]+")
 newz = z + 65
 Debug.Print v(I)
testExpression = REGPLACEFORMULA(testExpression, z, CStr(newz))
Debug.Print testExpression
Next I
r.Formula = testExpression
End Sub
Function BankersRound(num, precision)
BankersRound = Round(num, precision)
End Function
Function JOIN_Range(r As Range, Optional delim = "_")
'????????? ????????? ????
  
Dim myArr As Variant
Dim Val As String
  v = r
   Dim rngRectangle As Range, rngRows As Range, rngColumns As Range
    Set rngRectangle = r 'Selection
'    ‘ Определяет вертикальный вектор массива
    Set rngRows = rngRectangle.Resize(, 1)
'    ‘ Определяет горизонтальный вектор массива
    Set rngColumns = rngRectangle.Resize(1)
'    rngRectangle = Evaluate("IF(ROW(" & rngRows.Address & "), _
        IF(COLUMN(" & rngColumns.Address &"),UPPER(" & rngRectangle.Address & ")))")
myArrH = Application.Transpose(Application.Transpose(r))
myArrV = Application.Transpose(r)
If r.Rows.Count > r.Columns.Count Then
 Val = Join(myArrV, delim)
Else
 Val = Join(myArrH, delim)
 End If
JOIN_Range = Val
  
End Function
Sub Join_example()
'????????? ????????? ????
  
Dim valArr(1 To 3)
Dim Val As String
  
 valArr(1) = "???"
 valArr(2) = "MS"
 valArr(3) = "Excel"
   
' Val = Join(valArr, "_")

  Val = REGPLACE(REGPLACE(JOIN_Range(Selection), "_{1,}", "_"), "_$", "")
'  Val = JOIN_Range(Selection)
MsgBox Val
  
End Sub
Sub RectangularUpper()
'    ‘ Преобразует все ячейки в выделенном диапазоне в верхний регистр
    Dim rngRectangle As Range, rngRows As Range, rngColumns As Range
    Set rngRectangle = Selection
'    ‘ Определяет вертикальный вектор массива
    Set rngRows = rngRectangle.Resize(, 1)
'    ‘ Определяет горизонтальный вектор массива
    Set rngColumns = rngRectangle.Resize(1)
'    rngRectangle = Evaluate("IF(ROW(" & rngRows.Address & "), _
        IF(COLUMN(" & rngColumns.Address &"),UPPER(" & rngRectangle.Address & ")))")
End Sub

Sub AddNamesOffsetRange()
 Dim r As Range
 Dim NameMe As String
 For Each r In Selection.Cells
 If r <> "" Then
' r.value
 NameMe = r.value
 ActiveWorkbook.Names.Add name = NameMe, _
                      RefersTo:="'" & r.Offset(0, r.Offset(0, 1)).Parent.name & "'!" & r.Offset(0, r.Offset(0, 1)).address
 
 End If

 Next

 End Sub
Const LateBind = True
Function RegExpSubstitute(ReplaceIn, _
        ReplaceWhat As String, ReplaceWith As String)
    #If Not LateBind Then
    Dim RE As RegExp
    Set RE = New RegExp
    #Else
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
        #End If
    RE.pattern = ReplaceWhat
    RE.Global = True
    RegExpSubstitute = RE.Replace(ReplaceIn, ReplaceWith)
    End Function

'Sub createRanges()
''//https://itsalocke.com/blog/dynamic-named-range-generator/
''Таблица необработанных данных с 20 столбцами займет много времени для создания именованных диапазонов, учитывая, что я хочу:
''
''Динамический диапазон, охватывающий заголовки тоже для сводных таблиц
''Динамический диапазон без заголовков для vlookups
''Динамический диапазон для каждого столбца без заголовков
''Я использую макрос, назначенный симпатичной кнопке на моей ленте, чтобы генерировать все соответствующие диапазоны.
''
''Какие особые соображения?
''
''Структура - таблицы необработанных данных ВСЕГДА должны быть настроены особым образом - с первичным ключом слева и всегда заполненным, без пустых строк или столбцов
''
''Специальные символы - имена диапазонов не могут содержать специальные символы. VBA использует функциональность RegEx, чтобы удалить их.
''
''Числа - имена диапазонов также не могут иметь номера. Мы не можем просто вырезать числа, как специальные символы, потому что они могут быть важны, как Grade1, Grade2 и Grade3, и сворачивание их всех в имя Grade будет проблемой. Вместо этого макрос преобразует все числа в соответствующие буквы в алфавите.
''
''Сколько данных вырастет? По умолчанию я установил, что макрос будет использовать 10-кратное количество записей, присутствующих при запуске макроса - если он уже больше 25 тыс. Строк, это число нужно будет уменьшить, а если я не думаю, что в 10 раз число будет адекватный, я увеличу количество.
''
''' Specify some upfront variables
'rCol = ActiveSheet.UsedRange.Columns(1).Column
'rRow = ActiveSheet.UsedRange.Rows(1).Row
'sName = "'" & ActiveSheet.name & "'!"
'' This is where the row count gets multiplied to allow for growth
'LastRow = (ActiveSheet.UsedRange.Rows.Count - 1) * 10
'LastColumn = ActiveSheet.UsedRange.Columns.Count
'
'' Build a cleansed sheetname for use in naming the raw data tables
'SheetName = ActiveSheet.name
'SheetName = RegExpSubstitute(SheetName, "[^w+]", "")
'sheetname = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sheetname, "0", "a"), "1", "b"), "2", "c"), "3", "d"), "4", "e"), "5", "f"), "6", "g")
', "7", "h"), "8", "i"), "9", "j"), "|", "")
'
'' Build the headered raw data range
'ActiveWorkbook.Names.Add name:=SheetName, _
'        RefersTo:="=Offset(" & sName & Cells(rRow, rCol).address & ",0,0,counta(" _
'        & sName & Cells(rRow, rCol).address & ":" & Cells(LastRow, rCol).address _
'        & "),counta(" & sName & Cells(rRow, rCol).address & ":" & Cells(rRow, LastColumn * 3).address & "))"
'
'' Build the headerless raw data range
'ActiveWorkbook.Names.Add name:=SheetName & "HEADERLESS", _
'        RefersTo:="=Offset(" & sName & Cells(rRow + 1, rCol).address & ",0,0,counta(" _
'        & sName & Cells(rRow + 1, rCol).address & ":" & Cells(LastRow, rCol).address _
'        & "),counta(" & sName & Cells(rRow, rCol).address & ":" & Cells(rRow, LastColumn * 3).address & "))"
'
'' Create individual columns ranges
'While rCol <= LastColumn
'RangeName = Replace(Cells(rRow, rCol).value, " ", "")
'RangeName = RegExpSubstitute(RangeName, "[^w+]", "")
'rangeName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(rangeName, "0", "a"), "1", "b"), "2", "c"), "3", "d"), "4", "e"), "5", "f"), "6", "g")
', "7", "h"), "8", "i"), "9", "j"), "|", "")
'ActiveWorkbook.Names.Add name:=RangeName, _
'    RefersTo:="=Offset(" & sName & Cells(rRow + 1, rCol).address & ",0,0,counta(" & sName & Cells(rRow + 1, ActiveSheet.UsedRange.Columns(1).Column).address & ":" & Cells(LastRow, ActiveSheet.UsedRange.Columns(1).Column).address & "))"
'rCol = rCol + 1
'Wend
'
'End Sub

'Заменить именованный диапазон формулой
Public Sub replaceFormula()
    For Each ws In ActiveWorkbook.Worksheets
        For Each Rng In ws.UsedRange
            If Rng.HasFormula Then
                For Each namerng In Names
                    If InStr(CStr(Rng.Formula), namerng.name) > 0 Then
                        Rng.Formula = Replace(Rng.Formula, namerng.name, Replace(namerng.RefersTo, "=", ""))
                        Exit For
                    End If
                Next namerng
            End If
        Next Rng
    Next ws
End Sub
'http://access-excel.tips/find-external-links-broken-links/
'Код VBA - найти все ячейки, содержащие именованный диапазон
'Обратите внимание, что этот макрос имеет ограничение, вы не можете давать именованным диапазонам одинаковые имена.
Public Sub find_namedRng()
    Sheets.Add
    shtName = ActiveSheet.name
    Set summaryWS = ActiveWorkbook.Worksheets(shtName)
    summaryWS.Range("A1") = "Worksheet"
    summaryWS.Range("B1") = "Cell"
    summaryWS.Range("C1") = "Formula"
    summaryWS.Range("D1") = "Named Range"
    summaryWS.Range("E1") = "Refers To"
    
    For Each ws In ActiveWorkbook.Worksheets
        For Each Rng In ws.UsedRange
            If Rng.HasFormula Then
                For Each namerng In Names
                    If InStr(CStr(Rng.Formula), namerng.name) > 0 Then
                        nextrow = summaryWS.Range("A" & Rows.Count).End(xlUp).Row + 1
                        summaryWS.Range("A" & nextrow) = ws.name
                        summaryWS.Range("B" & nextrow) = Replace(Rng.address, "$", "")
                        summaryWS.Hyperlinks.Add Anchor:=summaryWS.Range("B" & nextrow), address:="", SubAddress:="'" & ws.name & "'!" & Rng.address
                        summaryWS.Range("C" & nextrow) = "'" & Rng.Formula
                        summaryWS.Range("D" & nextrow) = namerng.name
                        summaryWS.Range("E" & nextrow) = "'" & namerng.RefersTo
                        
                    End If
                Next namerng
            End If
        Next Rng
    Next ws
    
    Columns("A:E").EntireColumn.AutoFit
    LastRow = summaryWS.Range("A" & Rows.Count).End(xlUp).Row
    
    If LastRow = 1 Then
        MsgBox ("No Named Range found")
        Application.DisplayAlerts = False
        summaryWS.Delete
            Application.DisplayAlerts = True
    End If
End Sub
'Код VBA - найти все внешние ссылки и неработающие ссылки в книге
Sub listLinks()
    alinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(alinks) Then
        Sheets.Add
        shtName = ActiveSheet.name
        Set summaryWS = ActiveWorkbook.Worksheets(shtName)
        summaryWS.Range("A1") = "Worksheet"
        summaryWS.Range("B1") = "Cell"
        summaryWS.Range("C1") = "Formula"
        summaryWS.Range("D1") = "Workbook"
        summaryWS.Range("E1") = "Link Status"
        For Each ws In ActiveWorkbook.Worksheets
            If ws.name <> summaryWS.name Then
                For Each Rng In ws.UsedRange
                    If Rng.HasFormula Then
                        For j = LBound(alinks) To UBound(alinks)
                            FilePath = alinks(j)   'LinkSrouces returns full file path with file name
                            Filename = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))   'extract just the file name
                            filePath2 = Left(alinks(j), InStrRev(alinks(j), "\")) & "[" & Filename & "]"  'file path with brackets
                 
                            If InStr(Rng.Formula, FilePath) Or InStr(Rng.Formula, filePath2) Then
                                nextrow = summaryWS.Range("A" & Rows.Count).End(xlUp).Row + 1
                                summaryWS.Range("A" & nextrow) = ws.name
                                summaryWS.Range("B" & nextrow) = Replace(Rng.address, "$", "")
                                summaryWS.Hyperlinks.Add Anchor:=summaryWS.Range("B" & nextrow), address:="", SubAddress:="'" & ws.name & "'!" & Rng.address
                                summaryWS.Range("C" & nextrow) = "'" & Rng.Formula
                                summaryWS.Range("D" & nextrow) = FilePath
                                summaryWS.Range("E" & nextrow) = linkStatusDescr(ActiveWorkbook.LinkInfo(CStr(FilePath), xlLinkInfoStatus))
                                Exit For
                            End If
                        Next j
                            
                        For Each namedRng In Names
                            If InStr(Rng.Formula, namedRng.name) Then
                                FilePath = Replace(Split(Right(namedRng.RefersTo, Len(namedRng.RefersTo) - 2), "]")(0), "[", "") 'remove =' and range in the file path
                                nextrow = summaryWS.Range("A" & Rows.Count).End(xlUp).Row + 1
                                summaryWS.Range("A" & nextrow) = ws.name
                                summaryWS.Range("B" & nextrow) = Replace(Rng.address, "$", "")
                                summaryWS.Hyperlinks.Add Anchor:=summaryWS.Range("B" & nextrow), address:="", SubAddress:="'" & ws.name & "'!" & Rng.address
                                summaryWS.Range("C" & nextrow) = "'" & Rng.Formula
                                summaryWS.Range("D" & nextrow) = FilePath
                                summaryWS.Range("E" & nextrow) = linkStatusDescr(ActiveWorkbook.LinkInfo(CStr(FilePath), xlLinkInfoStatus))
                                Exit For
                            End If
                        Next namedRng
                    End If
                Next Rng
            End If
        Next
        Columns("A:E").EntireColumn.AutoFit
        LastRow = summaryWS.Range("A" & Rows.Count).End(xlUp).Row
        
        For r = 2 To LastRow
            If ActiveSheet.Range("E" & r).value = "File missing" Then
                countBroken = countBroken + 1
            End If
        Next
        
        If countBroken > 0 Then
            sInput = MsgBox("Do you want to remove broken links of status 'File missing'?", vbOKCancel + vbExclamation, "Warning")
            If sInput = vbOK Then
                For r = 2 To LastRow
                    If ActiveSheet.Range("E" & r).value = "File missing" Then
                        Sheets(Range("A" & r).value).Range(Range("B" & r).value).ClearContents
                        dummy = MsgBox(countBroken & " broken links removed", vbInformation)
                    End If
                Next
            End If
        End If
    Else
        MsgBox "No external links"
    End If
End Sub
Public Function linkStatusDescr(statusCode)
           Select Case statusCode
                Case xlLinkStatusCopiedValues
                    linkStatusDescr = "Copied values"
                Case xlLinkStatusIndeterminate
                    linkStatusDescr = "Unable to determine status"
                Case xlLinkStatusInvalidName
                    linkStatusDescr = "Invalid name"
                Case xlLinkStatusMissingFile
                    linkStatusDescr = "File missing"
                Case xlLinkStatusMissingSheet
                    linkStatusDescr = "Sheet missing"
                Case xlLinkStatusNotStarted
                    linkStatusDescr = "Not started"
                Case xlLinkStatusOK
                    linkStatusDescr = "No errors"
                Case xlLinkStatusOld
                    linkStatusDescr = "Status may be out of date"
                Case xlLinkStatusSourceNotCalculated
                    linkStatusDescr = "Source not calculated yet"
                Case xlLinkStatusSourceNotOpen
                    linkStatusDescr = "Source not open"
                Case xlLinkStatusSourceOpen
                    linkStatusDescr = "Source open"
                Case Else
                    linkStatusDescr = "Unknown status"
            End Select
End Function
'Как удалить переносы строк (возвраты каретки) из ячеек в Excel 2013, 2010 и 2007
'http://office-guru.ru/excel/kak-udalit-perenosy-strok-vozvraty-karetki-iz-jacheek-v-excel-2013-2010-i-2007-437.html
'VBA
Sub RemoveCarriageReturns()
    Dim MyRange As Range
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    delim = "_"
    pref = "g_th_"
' ActiveSheet.UsedRange
Set MyRanges = Selection
    For Each MyRange In MyRanges
        If 0 < InStr(MyRange, Chr(10)) Then
            MyRange = Replace(MyRange, Chr(10), delim)
            MyRange = pref & MyRange
        End If
    Next
 
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub '
Sub ApplyAllNames()
'http://qaru.site/questions/11166702/replace-coordinate-references-with-named-ranges
    Dim D As New Dictionary
    Dim C As Collection
    Dim ws As Worksheet, sh As Worksheet
    Dim A As Variant, v As Variant
    Dim nm As name, I As Long, n As Long, ref As String
    Dim r As Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    For Each ws In Worksheets
        Set C = New Collection
        D.Add ws.name, C
    Next ws
    For Each nm In Names
        ref = Split(nm.RefersTo, "!")(0) '=sheet name of ref
        ref = Mid(ref, 2) 'get rid of "="
        D(ref).Add nm
    Next nm

    'replace each collection of names
    'by an array sorted in order of descending length
    Set sh = Worksheets.Add
    For Each ws In Worksheets
        If ws.name <> sh.name Then
            Set C = D(ws.name)
            n = C.Count
            If n = 0 Then
                D(ws.name) = Array()
            Else
                ReDim A(1 To n, 1 To 2)
                For I = 1 To n
                    A(I, 1) = C(I).name
                    A(I, 2) = Len(C(I).RefersTo)
                Next I
                Set r = sh.Range(sh.Cells(1, 1), sh.Cells(n, 2))
                r.value = A
                r.Sort key1:=Range("B1:B" & n), order1:=xlDescending, Header:=xlNo
                A = r.value
                D(ws.name) = A
            End If
        End If
    Next ws
    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True

    'now loop over sheets and name array
    For Each ws In Sheets
        For Each sh In Sheets
            A = D(sh.name)
            If ws.name = sh.name Then
                On Error Resume Next
                    For I = 1 To UBound(A)
                        ws.Cells.ApplyNames A(I, 1)
                    Next I
                On Error GoTo 0
            Else
                For I = 1 To UBound(A)
                    Set v = Names(A(I, 1))
                    ref = Mid(v.RefersTo, 2) 'name with "=" removed
                    ws.Cells.Replace ref, v.name
                    ref = Replace(ref, "$", "")
                    ws.Cells.Replace ref, v.name
                Next I
            End If
            Debug.Print ws.name & " <- " & sh.name
            DoEvents
        Next sh
    Next ws
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
'https://www.extendoffice.com/documents/excel/3138-excel-find-where-named-range-is-used.html
'Теперь в первом открывшемся диалоговом окне Kutools for Excel введите имя рабочего листа и нажмите кнопку ОК ;
'а затем во втором открывшемся диалоговом окне введите имя определенного именованного диапазона и нажмите кнопку ОК
'Теперь появляется третье диалоговое окно Kutools for Excel,
'в котором перечислены ячейки с использованием определенного именованного диапазона



Sub Find_namedrange_place()
Dim xRg As Range
Dim xCell As Range
Dim xSht As Worksheet
Dim xFoundAt As String
Dim xAddress As String
Dim xShName As String
Dim xSearchName As String
On Error Resume Next
xShName = Application.InputBox("Please type a sheet name you will find cells in:", "Kutools for Excel", Application.ActiveSheet.name)
Set xSht = Application.Worksheets(xShName)
Set xRg = xSht.Cells.SpecialCells(xlCellTypeFormulas)
On Error GoTo 0
If Not xRg Is Nothing Then
xSearchName = Application.InputBox("Please type the name of named range:", "Kutools for Excel")
Set xCell = xRg.Find(What:=xSearchName, LookIn:=xlFormulas, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
        If Not xCell Is Nothing Then
xAddress = xCell.address
If IsPresent(xCell.Formula, xSearchName) Then
xFoundAt = xCell.address
End If
            Do
Set xCell = xRg.FindNext(xCell)
If Not xCell Is Nothing Then
If xCell.address = xAddress Then Exit Do
If IsPresent(xCell.Formula, xSearchName) Then
If xFoundAt = "" Then
xFoundAt = xCell.address
Else
xFoundAt = xFoundAt & ", " & xCell.address
End If
End If
Else
Exit Do
End If
Loop
End If
If xFoundAt = "" Then
MsgBox "The Named Range was not found", , "Kutools for Excel"
Else
MsgBox "The Named Range has been found these locations: " & xFoundAt, , "Kutools for Excel"
End If
On Error Resume Next
xSht.Range(xFoundAt).Select
End If
End Sub
Private Function IsPresent(sFormula As String, sName As String) As Boolean
Dim xPos1 As Long
Dim xPos2 As Long
Dim xLen As Long
Dim I As Long
xLen = Len(sFormula)
xPos2 = 1
Do
xPos1 = InStr(xPos2, sFormula, sName) - 1
If xPos1 < 1 Then Exit Do
IsPresent = IsVaildChar(sFormula, xPos1)
xPos2 = xPos1 + Len(sName) + 1
If IsPresent Then
If xPos2 <= xLen Then
IsPresent = IsVaildChar(sFormula, xPos2)
End If
End If
Loop
End Function
Private Function IsVaildChar(sFormula As String, Pos As Long) As Boolean
Dim I As Long
IsVaildChar = True
For I = 65 To 90
If UCase(Mid(sFormula, Pos, 1)) = Chr(I) Then
IsVaildChar = False
Exit For
End If
Next I
If IsVaildChar = True Then
If UCase(Mid(sFormula, Pos, 1)) = Chr(34) Then
IsVaildChar = False
End If
End If
If IsVaildChar = True Then
If UCase(Mid(sFormula, Pos, 1)) = Chr(95) Then
IsVaildChar = False
End If
End If
End Function

Sub getHeaderNameRange()
name = "gКошторис"
Dim r As Range
Dim nd As Range
Dim wb As Workbook
Dim nd_Header As Range
Set nd = ThisWorkbook.Names(name).RefersToRange
'Комплект
Set wb = mywbBook("test.xlsm", ThisWorkbook.Path & "\")
If wb Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & "Комплект.xlsm")
Set ndc = wb.Names(name).RefersToRange
For Each r In ndc.Cells
If Not r.Formula Like "=*" Then
'If r.value <> "" Or r.value <> Empty Then
ndc.Parent.Cells.Cells(r.Row, r.Column).value = nd.Parent.Cells(r.Row, r.Column).value
'End If
Else
'r.Cells(r.Row, r.Column).Formula = nd.Cells(r.Row, r.Column).Formula
End If
Next
'ThisWorkbook.Save
'ThisWorkbook.Close

End Sub
Function LeterАсtiveSheet(Optional name)
If IsMissing(name) Then name = ActiveSheet.name
Select Case name

Case "РГК"
LeterАсtiveSheet = "g"
Case "РД"
LeterАсtiveSheet = "d"
Case "Молодняки"
LeterАсtiveSheet = "m"
End Select

End Function
