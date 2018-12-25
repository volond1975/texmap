Attribute VB_Name = "importexport"

'=====================================================
'Sub ВТеккарту()
'Sub ИзТехкарты()
'Sub PrintSelectionTexkart()
'==================================================
'UDF
'Function LeterSheetShablon(shablon)


'=====================================================



Sub ВТеккарту()


Dim loj As clsmListObjs
Dim json As Object
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loft  As ListColumn
Dim rl As ListRow
Dim r As Range
Dim rr As Range
Dim v As Range
Dim wb As Workbook
Dim TypeOfFelling As String '-вид рубки
Dim fTypeOfFelling As Range '-вид рубки range Форма
Dim tTypeOfFelling As Range '-вид рубки range Техкарта
 Call Outro
Set wb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize wb
Set lof = loj.items("Форма")
Set lou = loj.items("Техкарта")
Set lofu = loj.items("Шаблон")
Set lofc = lou.ListColumns("Сведено")
Set fTypeOfFelling = .ValueListObject("Форма", "Параметр", "Значение", "Рубка Лист")
fTypeOfFelling.value = ActiveSheet.name



Set z = .ValueListObject("Шаблон", "Наименование", "Имя", .ValueListObject("Форма", "Параметр", "Значение", "Рубка Лист"))
Set Сведено = Range(.ValueListObject("Форма", "Параметр", "Адрес", "Сведено"))
Set rr = lofc.DataBodyRange.Find(What:=Сведено.value, LookAt:=xlWhole)
If Not rr Is Nothing Then
Set rl = lou.ListRows(rr.Row - 1)
Else
Set rl = lou.ListRows.Add
End If

 For Each r In lof.ListColumns("Техкарта").DataBodyRange.Cells
 If r.value <> "" Then
Set loft = lou.ListColumns(r.value)
Set v = Intersect(rl.Range, loft.Range)
If (Not (v.FormulaLocal Like "=*" Or lof.ListColumns("Адрес").DataBodyRange.Cells(r.Row - 1, 1) = "")) Then  'Or r.value Like "Витрати*"
v.value = Range(lof.ListColumns("Адрес").DataBodyRange.Cells(r.Row - 1, 1))
End If
 End If
 Next
'Сохраним как JSON:
'Значения из листа Расчет
'Значения из именованного диапазона Function NameEstimates {Буква листа шаблона}&{Имя шаблона}&"_кошторис"
Dim rJSON As Range
Set rJSON = Intersect(rl.Range, lou.ListColumns("jsonKoshtoris").DataBodyRange)
rJSON.value = modNamedRanges.setJsonToNamedRange(wb, NameEstimates(fTypeOfFelling.value))

'v.value = URLDecode(ConvertListObjectToJson("jsonTexzavdannya"))
'Set json = ParseJson(ПоИмениJSON(LeterSheetShablon(fTypeOfFelling.value) & Y.value & "_кошторис"))
'Set json = ParseJson(ПоИмениJSON(NameEstimates(fTypeOfFelling.value)))
'==================================================
'modNamedRanges
'Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'Sub setJsonToNamedRange(wb As Workbook, name As String, strJson As String)
'Function getJsonToNamedRange(wb As Workbook, name As String, strJson As String)
'Function getHeadersRange(wb As Workbook, name)
'Function ПоИмениJSON(ndrName) FIXME
'==================================================
Call JSONВКошторис(json, fTypeOfFelling)
'v.value = ПоИмениJSON(LeterSheetShablon(fTypeOfFelling.value) & fTypeOfFelling.value & "_кошторис")
Set v = Intersect(rl.Range, lou.ListColumns("JSONРасчет").DataBodyRange)
v.value = URLDecode(ConvertListObjectToJson("rachtet"))
End With

End Sub





Sub ИзТехкарты()


Dim loj As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loft  As ListColumn
Dim rl As ListRow
Dim r As Range
Dim v As Range
 Dim wb As Workbook
 Set wb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize wb
Set lof = loj.items("Форма")
Set lofu = loj.items("Шаблон")
Set lou = loj.items("Техкарта")
Set lofc = lou.ListColumns("Сведено")
'Set r = .Columns("AE:AE").Find(What:=Form.cbNTK, LookAt:=xlWhole)
Set Сведено = Range(.ValueListObject("Форма", "Параметр", "Адрес", "Сведено"))
Set rl = .ActiveListObjectActiveRow
Set r = lofc.Range.Find(What:=Сведено.value, LookAt:=xlWhole)
Set z = .ValueListObject("Форма", "Параметр", "Значение", "Рубка Лист")
z.value = rl.Range.Cells(1).value

 For Each r In lof.ListColumns("Техкарта").DataBodyRange.Cells
 If r.value <> "" Then
Set loft = lou.ListColumns(r.value)
Set v = Intersect(rl.Range, loft.Range)
If Not Range(lof.ListColumns("Адрес").DataBodyRange.Cells(r.Row - 1, 1)).Formula Like "=*" Then Range(lof.ListColumns("Адрес").DataBodyRange.Cells(r.Row - 1, 1)) = v.value
 End If
 Next
Set v = Intersect(rl.Range, lou.ListColumns("JSONРасчет").DataBodyRange)

End With
Worksheets(z.value).Activate
'Range(lof.ListColumns("Адрес").DataBodyRange.Cells(r.Row - 1, 1))
End Sub

Sub PrintSelectionTexkart()
Dim r As Range

Dim loj As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loft  As ListColumn
Dim rl As ListRow
Dim rr As Range
Dim v As Range
Call Outro
 Dim wb As Workbook
 Set wb = ThisWorkbook
 Set r = Selection
Set loj = New clsmListObjs
With loj
.Initialize wb
Set lou = loj.items("Техкарта")


For Each rr In r.Cells
lou.parent.Activate
rr.Activate
Call ИзТехкарты

'Application.Calculate
Set rl = lou.ListRows(rr.Row - 1)
Set loft = lou.ListColumns("Шаблон2")
Set v = Intersect(rl.Range, loft.Range)
wb.Worksheets(v.value).Activate
Call ВТеккарту

ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        
Next
End With
End Sub

Sub copy()
Dim wb As Workbook
Dim twb As Workbook
If Not IsBookOpen("G:\Dropbox\Чат\Комплект.xlsm") Then
Set wb = Workbooks.Open("G:\Dropbox\Чат\Комплект.xlsm", ReadOnly:=True)
Else
Set wb = Workbooks("Комплект.xlsm")
End If

Set twb = ThisWorkbook
twb.Save
twb.Sheets("Техкарта").copy Before:=wb.Sheets(1)
'wb.Close False
End Sub
Sub ВКомплект()


Dim loj As clsmListObjs
Dim los As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loft  As ListColumn
Dim rl As ListRow
Dim r As Range
Dim v As Range
 Dim wb As Workbook
Dim twb As Workbook

Dim B As Workbook
Set twb = ThisWorkbook

Set wb = mywbBook("Комплект.xlsm", ThisWorkbook.Path & "\")
If wb Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & "Комплект.xlsm")
Set loj = New clsmListObjs
With loj
.Initialize twb

Set lou = loj.items("Техкарта")
lou.DataBodyRange.copy

End With
Set los = New clsmListObjs
With los

wb.Activate
.Initialize wb
Set lof = los.items("Таблица1")
lof.parent.Activate
lof.DataBodyRange.Cells.ClearContents
lof.parent.Range(lou.DataBodyRange.Address).Select
lof.parent.Range(lou.DataBodyRange.Address).value = lou.DataBodyRange.value
'ActiveSheet.Paste
End With
Call РасчетВКомплект
ThisWorkbook.Save
ThisWorkbook.Close



End Sub
Sub РасчетВКомплект()


Dim loj As clsmListObjs
Dim los As clsmListObjs
Dim lof As ListObject
Dim lou As ListObject
Dim lofu As ListObject
Dim lofc  As ListColumn
Dim loft  As ListColumn
Dim rl As ListRow
Dim r As Range
Dim v As Range
 Dim wb As Workbook
Dim twb As Workbook

Dim B As Workbook
Set twb = ThisWorkbook

Set wb = mywbBook("Комплект.xlsm", ThisWorkbook.Path & "\")
If wb Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & "Комплект.xlsm")
Set loj = New clsmListObjs
With loj
.Initialize twb

Set lou = loj.items("rachtet")
lou.DataBodyRange.copy

End With
Set los = New clsmListObjs
With los

wb.Activate
.Initialize wb
Set lof = los.items("rachtet")
lof.parent.Activate
lof.parent.Range(lou.DataBodyRange.Address).Select
lof.DataBodyRange.value = lou.DataBodyRange.value

End With

End Sub


Public Sub КошторисВКомплект()
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
Dim http As Object, JSON As Object, i As Integer, ndr As name
Dim w As Range
Dim wb As Workbook
Dim ИмяКошторис

Set twb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize twb

Set lou = loj.items("Техкарта")


Set rl = .ActiveListObjectActiveRow
Set lofc = lou.ListColumns("Вид рубки")
Set v = Intersect(rl.Range, lofc.DataBodyRange)
Set loJSON = lou.ListColumns("jsonKoshtoris")
Set rJSON = Intersect(rl.Range, loJSON.DataBodyRange)




ndrName = LeterSheetShablon(v.value) & v.value & "_кошторис"


param = "offset"
Set JSON = ParseJson(rJSON.value)
Set wb = mywbBook("Комплект.xlsm", ThisWorkbook.Path & "\")
If wb Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & "Комплект.xlsm")
Set ndrКомплект = wb.Names(ndrName)
i = 2
For Each Item In JSON
For Each Jtem In Item.Keys
Dim f As Variant
If Not IsEmpty(Jtem) Then
f = Split(Jtem, "_")
If param = "offset" Then
Set fistRangeКошторис = ndrКомплект.RefersToRange.Cells(1, 1)
'Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
Set w = fistRangeКошторис.Offset(Val(f(0)), Val(f(1)))

End If
If param = "address" Then
Set w = Sheets("Лист8").Range(Jtem)
End If
'fixme
'If param = "NameColumn" Then myitem(firstRange.Offset(0, j - 1).address) = firstRange.Offset(i - 1, j - 1).value


'Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
'Sheets("Лист8").Range(Jtem).value = Item(Jtem)
 If Not w.Formula Like "=*" Then w.value = Item(Jtem)
End If


Next
i = i + 1
Next
'MsgBox ("complete")
End With
End Sub
Function ПоИмениJSON(ndrName)
'
' ПоВыделению2 Макрос
'
Dim rng As Range, items As New Collection, myitem As New Dictionary, i As Integer, cell As Variant



'
   Dim r As Range
   Dim firstRange As Range
   Dim rRange As Range
   Dim rDataRange As Range
   Arr = ThisWorkbook.Names(ndrName).RefersToRange
   param = "offset"
'   rowCol = UBound(v)
   Dim rHeader As Range
 Set rHeader = ThisWorkbook.Names(ndrName).RefersToRange.Rows(1)
 Set rRange = ThisWorkbook.Names(ndrName).RefersToRange
 Set firstRange = rRange.Cells(1, 1)
' set rDataRange=range(cells(
For i = LBound(Arr, 1) To UBound(Arr, 1)
  For j = LBound(Arr, 2) To UBound(Arr, 2)
'  Debug.Print firstRange.Offset(0, j - 1)
'  myitem(firstRange.Offset(i - 1, j - 1).r) = firstRange.Offset(i - 1, j - 1).value
If firstRange.Offset(i - 1, j - 1).value <> "" Then
myitem(i - 1 & "_" & j - 1) = firstRange.Offset(i - 1, j - 1).value
If param = "offset" Then myitem(i - 1 & "_" & j - 1) = firstRange.Offset(i - 1, j - 1).value
If param = "address" Then myitem(firstRange.Offset(i - 1, j - 1).Address) = firstRange.Offset(i - 1, j - 1).value
If param = "NameColumn" Then
myitem(firstRange.Offset(0, j - 1).Address) = firstRange.Offset(i - 1, j - 1).value

End If

Else

End If
  Next j
  items.Add myitem
Set myitem = Nothing

Next i
Dim strJson As String

strJson = URLDecode(RussianStringToURLEncode_New(ConvertToJson(items)))

Debug.Print URLDecode(RussianStringToURLEncode_New(strJson))
'Set JSON = ParseJson(strJson)
ПоИмениJSON = strJson

End Function

Sub getNameEstimates()
'Получение имени кошториса
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
Dim http As Object, JSON As Object, i As Integer, ndr As name
Dim w As Range
Dim wb As Workbook
Dim ИмяКошторис

Set twb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize twb

Set lou = loj.items("Техкарта")
Set rl = .ActiveListObjectActiveRow
Set lofc = lou.ListColumns("Вид рубки")
Set v = Intersect(rl.Range, lofc.DataBodyRange)

Set loJSON = lou.ListColumns("jsonKoshtoris")
Set rJSON = Intersect(rl.Range, loJSON.DataBodyRange)
Set JSON = ParseJson(rJSON.value)


ndrName = NameEstimates(v.value)
End With
End Sub

Function NameEstimates(TypeOfFelling As String)
'TypeOfFelling-вид рубки
prefix = "_кошторис"
NameEstimates = LeterSheetShablon(v.value) & v.value & prefix
End Function
Function NamedRangeEstimates(wb As Workbook, TypeOfFelling As String)
'TypeOfFelling-вид рубки
Dim name
name = NameEstimates(TypeOfFelling)
Set NamedRangeEstimates = wb.Names(name)
End Function



Function LeterSheetShablon(shablon)
'With Form
Select Case shablon

Case "РГК"
LeterSheetShablon = "g"
Case "РД"
LeterSheetShablon = "d"
Case "Молодняки"
LeterSheetShablon = "m"
End Select
'End With
End Function
