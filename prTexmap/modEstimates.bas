Attribute VB_Name = "modEstimates"

'========================================================
'modEstimates -Кошторис
'========================================================
'UDF
'Function NameEstimates(TypeOfFelling As String)
'Function NamedRangeEstimates(wb As Workbook, TypeOfFelling As String)
'========================================================
Sub getNameEstimatesActiveListObjectActiveRow()
'Получение имени кошториса по активной строке ListObject
Dim loj As clsmListObjs
Dim los As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loJSON  As ListColumn
Dim twb As Workbook
Dim rl As ListRow
Dim ndrКомплект As name
Dim r As Range
Dim v As Range
Dim http As Object, json As Object, i As Integer, ndr As name
Dim w As Range
Dim wb As Workbook
Dim ИмяКошторис

Set twb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize twb
loName = "Техкарта"
lcName = "Вид рубки"
Set lou = loj.items(loName)
Set rl = .ActiveListObjectActiveRow
Set lofc = lou.ListColumns(lcName)
Set v = Intersect(rl.Range, lofc.DataBodyRange)

Set loJSON = lou.ListColumns("jsonKoshtoris")
Set rJSON = Intersect(rl.Range, loJSON.DataBodyRange)
Set json = ParseJson(rJSON.value)


ndrName = NameEstimates(v.value)
End With
End Sub


'Public Sub JSONВКошторис(json As Object, v)
Public Sub setJsonToEstimates(json As Object, TypeOfFelling As String)
'modEstimates.setJsonToEstimates
    On Error GoTo EH
Dim wbTexmap As Workbook
Dim wbCompl As Workbook
Dim strComplFileName As String
    
Dim loj As clsmListObjs
Dim los As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loJSON  As ListColumn

Dim rl As ListRow
Dim ndrКомплект As name
Dim r As Range

Dim http As Object, i As Integer, ndr As name
Dim w As Range
'Dim wb As Workbook
'Dim twb As Workbook


Dim ИмяКошторис
strComplFileName = "Комплект.xlsm"
'Open wbCompl
Set wbTexmap = ThisWorkbook
Set wbCompl = mywbBook(strComplFileName, wbTexmap.Path & "\")
If wbCompl Is Nothing Then MsgBox ("Файл " & TwbTexmap.Path & "\" & strComplFileName & "не обнаружен по пути")
'DisplayError Err.Source, Err.Description, "modEstimates.setJsonToEstimates"



ndrName = NameEstimates(TypeOfFelling)


'Set JSON = ParseJson(rJSON.value)
'Set wb = mywbBook("Комплект.xlsm", ThisWorkbook.Path & "\")
'If wb Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & "Комплект.xlsm")
Set loj = New clsmListObjs
With loj
.Initialize wb

Set lou = loj.items("Таблица1")


'Set rl = .ActiveListObjectActiveRow
'Set lofc = lou.ListColumns("Вид рубки")
'Set v = Intersect(rl.Range, lofc.DataBodyRange)
'Set loJSON = lou.ListColumns("jsonKoshtoris")
'Set rJSON = Intersect(rl.Range, loJSON.DataBodyRange)

Set ndrКомплект = wb.Names(ndrName)
i = 2
For Each Item In json
For Each jtem In Item.Keys
Dim f As Variant
If Not IsEmpty(jtem) Then
f = Split(jtem, "_")
If param = "offset" Then
Set fistRangeКошторис = ndrКомплект.RefersToRange.Cells(1, 1)
'Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
Set w = fistRangeКошторис.Offset(Val(f(0)), Val(f(1)))

End If
If param = "address" Then
Set w = Sheets("Лист8").Range(jtem)
End If
'fixme
'If param = "NameColumn" Then myitem(firstRange.Offset(0, j - 1).address) = firstRange.Offset(i - 1, j - 1).value


'Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
'Sheets("Лист8").Range(Jtem).value = Item(Jtem)
 If Not w.Formula Like "=*" Then w.value = Item(jtem)
End If


Next
i = i + 1
Next
MsgBox ("complete")
End With
Done:
    Exit Sub
EH:
    DisplayError Err.Source, Err.Description, "Module1.Topmost"
End Sub
'Public Sub JSONВКошторис(json As Object, v)
Public Sub JSONВКошторис(json As Object, v)
Dim loj As clsmListObjs
Dim los As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loJSON  As ListColumn
Dim twb As Workbook
Dim rl As ListRow
Dim ndrКомплект As name
Dim r As Range

Dim http As Object, i As Integer, ndr As name
Dim w As Range
Dim wb As Workbook
Dim ИмяКошторис

Set twb = ThisWorkbook




ndrName = LeterSheetShablon(v.value) & v.value & "_кошторис"

'Set http = CreateObject("MSXML2.XMLHTTP")
'http.Open "GET", "http://jsonplaceholder.typicode.com/users", False
'http.Send
param = "offset"
'Set JSON = ParseJson(rJSON.value)
Set wb = mywbBook("Комплект.xlsm", ThisWorkbook.Path & "\")
If wb Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & "Комплект.xlsm")
Set loj = New clsmListObjs
With loj
.Initialize wb

Set lou = loj.items("Таблица1")


'Set rl = .ActiveListObjectActiveRow
'Set lofc = lou.ListColumns("Вид рубки")
'Set v = Intersect(rl.Range, lofc.DataBodyRange)
'Set loJSON = lou.ListColumns("jsonKoshtoris")
'Set rJSON = Intersect(rl.Range, loJSON.DataBodyRange)

Set ndrКомплект = wb.Names(ndrName)
i = 2
For Each Item In json
For Each jtem In Item.Keys
Dim f As Variant
If Not IsEmpty(jtem) Then
f = Split(jtem, "_")
If param = "offset" Then
Set fistRangeКошторис = ndrКомплект.RefersToRange.Cells(1, 1)
'Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
Set w = fistRangeКошторис.Offset(Val(f(0)), Val(f(1)))

End If
If param = "address" Then
Set w = Sheets("Лист8").Range(jtem)
End If
'fixme
'If param = "NameColumn" Then myitem(firstRange.Offset(0, j - 1).address) = firstRange.Offset(i - 1, j - 1).value


'Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
'Sheets("Лист8").Range(Jtem).value = Item(Jtem)
 If Not w.Formula Like "=*" Then w.value = Item(jtem)
End If


Next
i = i + 1
Next
MsgBox ("complete")
End With
End Sub

Function NameEstimates(TypeOfFelling As String)
'TypeOfFelling-вид рубки
prefix = "_кошторис"
NameEstimates = LeterSheetShablon(TypeOfFelling) & TypeOfFelling & prefix
End Function
Function NamedRangeEstimates(wb As Workbook, TypeOfFelling As String)
'TypeOfFelling-вид рубки
Dim name
name = NameEstimates(TypeOfFelling)
Set NamedRangeEstimates = wb.Names(name)
End Function


Function getNameEstimatesActiveListObjectActiveRow1() As Range
'Получение имени кошториса по активной строке ListObject
Dim loj As clsmListObjs
Dim los As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loJSON  As ListColumn
Dim twb As Workbook
Dim rl As ListRow
Dim ndrКомплект As name
Dim r As Range
Dim rTypeOfFelling As Range
Dim http As Object, json As Object, i As Integer, ndr As name
Dim w As Range
Dim wb As Workbook
Dim ИмяКошторис
loName = "Техкарта"
lcName = "Вид рубки"
Set twb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize twb
Set lou = loj.items(loName)
Set rl = .ActiveListObjectActiveRow
Set lofc = lou.ListColumns(lcName)
Set rTypeOfFelling = Intersect(rl.Range, lofc.DataBodyRange)
End With
End Function
'==================================================
'modNamedRanges

'Sub setJsonToNamedRange(wb As Workbook, name As String, strJson As String)
'Function getJsonToNamedRange(wb As Workbook, name As String)

'==================================================
Sub testopenThisWorkBookPathByName()

Dim wbTexmap As Workbook
Dim wbCompl As Workbook
Dim TypeOfFelling As String
Dim strComplFileName As String
Dim strJson As String
'On Error GoTo ErH
call Outro
Call mMacros.Intro
Set wbTexmap = ThisWorkbook
TypeOfFelling = ActiveSheet.name
strJson = getJsonToNamedRange(wbTexmap, NameEstimates(TypeOfFelling))

strComplFileName = "Комплект.xlsm"
Set wbCompl = openThisWorkBookPathByName(strComplFileName)
Call setJsonToNamedRange(wbCompl, NameEstimates(TypeOfFelling), strJson)
Call Outro
'Done:
End Sub
'ErH:
'    DisplayError Err.Source, Err.description, "Module1.testopenThisWorkBookPathByName"
'End Sub

Function openThisWorkBookPathByName(strComplFileName As String, Optional wbTexmap As Workbook)
On Error GoTo EH
'strComplFileName = "Комплект.xlsm"
'Open wbCompl by strComplFileName
'Dim wbTexmap As Workbook
'Dim wbCompl As Workbook
Set wbTexmap = ThisWorkbook
Set wbCompl = mywbBook(strComplFileName, wbTexmap.Path & "\")
If wbCompl Is Nothing Then MsgBox ("Файл " & TwbTexmap.Path & "\" & strComplFileName & "не обнаружен по пути")
Done:
Set openThisWorkBookPathByName = wbCompl
    Exit Function
EH:
    DisplayError Err.Source, Err.Description, "Module1.openThisWorkBookPathByName"
End Function
Sub WorksheetByNameRangeActivate()
Dim wbTexmap As Workbook
Dim lou As ListObject
'On Error GoTo EH
Call Outro
Name = "Техкарта"
Set wbTexmap = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize wbTexmap
Set lou = loj.items(Name)
lou.parent.Activate
End With
EH:
End Sub
