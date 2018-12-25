Attribute VB_Name = "modNamedRanges"
Sub NameRange_Add()
'PURPOSE: Various ways to create a Named Range
'SOURCE: www.TheSpreadsheetGuru.com

Dim cell As Range
Dim rng As Range
Dim RangeName As String
Dim CellName As String

'Single Cell Reference (Workbook Scope)
  RangeName = "Price"
  CellName = "D7"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add name:=RangeName, RefersTo:=cell


'Single Cell Reference (Worksheet Scope)
  RangeName = "Year"
  CellName = "A2"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  Worksheets("Sheet1").Names.Add name:=RangeName, RefersTo:=cell


'Range of Cells Reference (Workbook Scope)
  RangeName = "myData"
  CellName = "F9:J18"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add name:=RangeName, RefersTo:=cell


'Secret Named Range (doesn't show up in Name Manager)
  RangeName = "Username"
  CellName = "L45"
  
  Set cell = Worksheets("Sheet1").Range(CellName)
  ThisWorkbook.Names.Add name:=RangeName, RefersTo:=cell, Visible:=False

End Sub
Sub NamedRange_DeleteAll()
'PURPOSE: Delete all Named Ranges in the ActiveWorkbook (Print Areas optional)
'SOURCE: www.TheSpreadsheetGuru.com

Dim nm As name
Dim DeleteCount As Long

'Delete PrintAreas as well?
  UserAnswer = MsgBox("Do you want to skip over Print Areas?", vbYesNoCancel)
    If UserAnswer = vbYes Then SkipPrintAreas = True
    If UserAnswer = vbCancel Then Exit Sub

'Loop through each name and delete
  For Each nm In ActiveWorkbook.Names
    On Error GoTo Skip
    
    If SkipPrintAreas = True And Right(nm.name, 10) = "Print_Area" Then GoTo Skip
    
    nm.Delete
    DeleteCount = DeleteCount + 1

Skip:
    
  Next
  
'Reset Error Handler
  On Error GoTo 0
     
'Report Result
  If DeleteCount = 1 Then
    MsgBox "[1] name was removed from this workbook."
  Else
    MsgBox "[" & DeleteCount & "] names were removed from this workbook."
  End If

End Sub
Sub NamedRange_DeleteErrors()
'PURPOSE: Delete all Named Ranges with #REF error in the ActiveWorkbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim nm As name
Dim DeleteCount As Long

'Loop through each name and delete
  For Each nm In ActiveWorkbook.Names
    On Error GoTo Skip
    If InStr(1, nm.RefersTo, "#REF!") > 0 Then
      nm.Delete
      DeleteCount = DeleteCount + 1
    End If
Skip:
    
  Next
  
'Reset Error Handler
  On Error GoTo 0
   
'Report Result
  If DeleteCount = 1 Then
    MsgBox "[1] errorant name was removed from this workbook."
  Else
    MsgBox "[" & DeleteCount & "] errorant names were removed from this workbook."
  End If
  
End Sub

Sub NamedRange_DeleteReferBook()
'PURPOSE: Delete all Named Ranges with #REF error in the ActiveWorkbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim nm As name
Dim DeleteCount As Long

'Loop through each name and delete
  For Each nm In ActiveWorkbook.Names
    On Error GoTo Skip
    If InStr(1, nm.RefersTo, ":\") > 0 Then
      nm.Delete
      DeleteCount = DeleteCount + 1
    End If
Skip:
    
  Next
  
'Reset Error Handler
  On Error GoTo 0
   
'Report Result
  If DeleteCount = 1 Then
    MsgBox "[1] errorant name was removed from this workbook."
  Else
    MsgBox "[" & DeleteCount & "] errorant names were removed from this workbook."
  End If
  
End Sub
Public Sub makeNameJunkVisible()
'https://stackoverflow.com/questions/34384066/xlfn-iferror-excel2013-deletion
'Удаление диапазонов с ошибкой .RefersTo "= # NAME?" и имя "_xlfn.IFERROR"
'так как он невидимый в диспечере Имен
' 1.Выполните код
' 2.Откройте Диспетчер имен, который находится на вкладке "Формула" в меню ленты.
'  Теперь в Диспетчере имен должен отображаться неисправный именованный диапазон,
'  и вы сможете удалить его.
Dim n As name
  For Each n In ThisWorkbook.Names
    If n.RefersTo = "=#NAME?" Then
      n.Visible = True
    End If
  Next n
End Sub
c


Sub rty()
Dim n As name
For Each n In ThisWorkbook.Names
 n.Visible = True
 Debug.Print n.name, n.RefersTo, n.Comment
Next
End Sub





Sub testNamedRange_Loop()
Call NamedRange_Loop
End Sub

Sub NamedRange_Loop(Optional wb As Workbook, Optional sh As Worksheet)
'PURPOSE: Delete all Named Ranges in the Active Workbook
'SOURCE: www.TheSpreadsheetGuru.com
'Желательно перед єтим віполнить makeNameJunkVisible
Dim nm As name



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
Dim lr As ListRow
'On Error Resume Next
Dim ИмяКошторис
If wb Is Nothing Then Set wb = ActiveWorkbook
Set twb = wb
Set loj = New clsmListObjs
With loj
.Initialize twb
Set lou = loj.items("NamedRanges")

'Loop through each named range in workbook
  For Each nm In wb.Names
  On Error GoTo Skip
  If Not nm.name = "NamedRange" Then
Debug.Print nm.name, nm.RefersTo, nm.Comment
Set findRange = .ValueListObject("NamedRanges", "name", "name", nm.name)
If Not findRange Is Nothing Then
For Each lr In lou.ListRows
        If Not Intersect(findRange, lr.Range) Is Nothing Then
          Set rl = lr
        Exit For
        End If
        Next
'Set rl = Intersect(findRange.Row, lou.ListRows)
Else
Set rl = .ActiveListObjectRowAdd(lou)
End If

Set lofc = lou.ListColumns("name")
Set v = Intersect(rl.Range, lofc.DataBodyRange)
 v = nm.name
 Set v = Nothing
Set lofc = lou.ListColumns("RefersTo")
Set v = Intersect(rl.Range, lofc.DataBodyRange)
  v = nm.RefersTo
  Set v = Nothing
Set lofc = lou.ListColumns("Comment")
Set v = Intersect(rl.Range, lofc.DataBodyRange)
  v = nm.Comment
  Set v = Nothing
    End If
Skip:
  Next nm
  End With
'Loop through each named range scoped to a specific worksheet
'  For Each nm In Worksheets("Sheet1").Names
'    Debug.Print nm.name, nm.RefersTo
'  Next nm

End Sub
Function NameExist(NameName As String)

Dim rRangeCheck As Range
    
    On Error Resume Next
    Set rRangeCheck = Range(NameName)
    On Error GoTo 0
    If rRangeCheck Is Nothing Then
       NameExist = False
    Else
       NameExist = True
    End If

End Function

Sub ggggjh()
Dim locls As clsmListObjs
Dim lo_forma As ListObject
Dim loc As ListColumn
Dim lo As ListObject
Dim r As Range
Dim AR As Range
Dim n As Range
Dim wb As Workbook
Set wb = ThisWorkbook
Set locls = New clsmListObjs
With locls
.Initialize wb
Set lo_forma = .items("Форма")
Set w = .ValueListObject("Форма", "Параметр", "Значение", "Рубка Лист")
Set loc = lo_forma.ListColumns("Имя")
For Each r In loc.DataBodyRange.Cells
Set AR = .ValueListObject("Форма", "Имя", "Адрес", r.value)

If Not IsEmpty(r.value) Then

If AR.value Like "* *" Then
Y = "'" & AR.value
Y = VBA.Replace(Y, "!", "'!")
Else

Y = AR.value
End If

If Not NameExist(r.value) And r.value <> "" Then
'Range(AR.value).name
Set ARК = .ValueListObject("Форма", "Имя", w.value, r.value)
If ARК.value = 1 Then

Set z = wb.Names.Add(r.value, Range(Y))
End If

Else
Set n = Range(r.value)
n.name.Delete
Set z = wb.Names.Add(r.value, Range(Y))
End If
End If
Next

End With
End Sub
Sub EnumActiveSheetNames()
    Dim i As Long
    Dim sName As String
    Dim sNameList As String
    Dim NameArray As Variant
    Dim sr As String, s As String
    Dim n As name
    sName = ActiveSheet.name
    sNameList = vbNullString
    For Each n In Names
       sr = n.RefersTo
       If Mid$(sr, 2, 1) = "'" Then
           i = InStr(4, sr, "'!")
           If i > 0 Then
               If Mid$(sr, 3, i - 3) = sName Then
                   sNameList = sNameList & n.name & ","
               End If
           End If
       Else
           i = InStr(2, sr, "!")
           If i > 0 Then
               If Mid$(sr, 2, i - 2) = sName Then
                   sNameList = sNameList & n.name & ","
               End If
           End If
       End If
    Next n
    If Len(sNameList) > 0 Then
        NameArray = Split(sNameList, ",")
        Debug.Print "?????????? ???? ?? ?????: " & UBound(NameArray)
        For i = LBound(NameArray) To UBound(NameArray) - 1
           Debug.Print NameArray(i)
        Next i
    Else
        Debug.Print "??? ???? ???????????? ?? ?????"
    End If
End Sub



'==================================================
'modNamedRanges
'Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'Sub setJsonToNamedRange(wb As Workbook, name As String, strJson As String)
'Function getJsonToNamedRange(wb As Workbook, name As String)
'Function getHeadersRange(wb As Workbook, name)
'Function ПоИмениJSON(ndrName) FIXME
'==================================================

Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'aki indexOf
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check where a value is in an array
Dim i As Long
For i = LBound(arr1) To UBound(arr1)
    If arr1(i) = vFind Then
        WhereInArray = i
        Exit Function
    End If
Next i
'if you get here, vFind was not in the array. Set to null
WhereInArray = Null
End Function

Sub setJsonToNamedRange(wb As Workbook, name As String, strJson As String)
'Set wb = ThisWorkbook
'name = "testRange1"
'strJson = ActiveCell.value
Dim json As Object
Dim n As name
Dim nr As Range
Dim nrh As Range
Dim b()
Dim rng As Range, items As New Collection, myitem As New Dictionary, cell As Variant


Set json = ParseJson(strJson)
Set n = wb.Names(name)
Set nr = n.RefersToRange
Arr = nr
Set firstRange = nr.Cells(1, 1)
b = getHeadersRange(wb, name)
i = 1
For Each Item In json
For Each jtem In Item.Keys
ind = WhereInArray(b, jtem)
Set w = firstRange.Offset(i, ind)
If Not w.Formula Like "=*" Then
w.value = Item(jtem)
End If
Next
i = i + 1
Next

 End Sub

Function getJsonToNamedRange(wb As Workbook, name As String)

Dim strJson As String
Dim n As name
Dim nr As Range
Dim nrh As Range
Dim b()
Dim rng As Range, items As New Collection, myitem As New Dictionary, cell As Variant
'Set wb = ThisWorkbook
'name = "testRange1"
'strJson = ActiveCell.value

Set n = wb.Names(name)
Set nr = n.RefersToRange
Arr = nr
Set firstRange = nr.Cells(1, 1)
b = getHeadersRange(wb, name)
For i = LBound(Arr, 1) + 1 To UBound(Arr, 1)
  For j = LBound(Arr, 2) To UBound(Arr, 2)
myitem(b(j - 1)) = firstRange.Offset(i - 1, j - 1).value
  Next j
  items.Add myitem
Set myitem = Nothing

Next i


strJson = URLDecode(RussianStringToURLEncode_New(ConvertToJson(items)))
Debug.Print strJson
getJsonToNamedRange = strJson
'Set json = ParseJson(strJson)
'nr.parent.Range("A1") = strJson

End Function


Function getHeadersRange(wb As Workbook, name)
Dim n As name
Dim nr As Range
Dim nrh As Range
Dim b()
Dim rng As Range, items As New Collection, myitem As New Dictionary, cell As Variant

Set n = wb.Names(name)
Set nr = n.RefersToRange
Arr = nr
Set firstRange = nr.Cells(1, 1)
Set nrh = nr.Rows(1)
Dim i
i = 1
j = 0
ReDim b(nrh.Cells.Count)
For Each r In nrh.Cells

If r = "" Then
b(j) = "'" & Space(i)
i = i + 1
Else
b(j) = r.value

End If
j = j + 1

Next
getHeadersRange = b
End Function
Function ПоИмениJSON(ndrName)
'
' FIXME
'
   Dim rng As Range, items As New Collection, myitem As New Dictionary, i As Integer, cell As Variant
   Dim strJson As String
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




Else

End If
  Next j
  items.Add myitem
Set myitem = Nothing

Next i


strJson = URLDecode(RussianStringToURLEncode_New(ConvertToJson(items, Whitespace:=2)))
'Debug.Print strJson

ПоИмениJSON = strJson

End Function
