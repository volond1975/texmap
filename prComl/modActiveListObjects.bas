Attribute VB_Name = "modActiveListObjects"

'---------------------------------------------------------------------------------------
' Procedure : ActiveListObject
' Author    : ВАЛЕРА
' Date      : 19.01.2016
' Purpose   :Возвращает ссылку на Активную Умную Таблицу ( Умная таблица в которой находится Активная ячейка)
' если в процедуру не переданы параметры
'---------------------------------------------------------------------------------------
'
Function ActiveListObject(Optional wb, Optional ShName, Optional r)
'Возвращает ссылку на Активную Умную Таблицу ( Умная таблица в которой находится Активная ячейка)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(wb) Then Set wb = ActiveWorkbook
If IsMissing(r) Then Set r = Range(ActiveCell.address)

If IsMissing(ShName) Then ShName = r.Parent.name


For Each MyTable In wb.Worksheets(ShName).ListObjects
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
        Set ActiveListObject = MyTable
        Exit Function
        Else
        Set ActiveListObject = Nothing
        
        End If
        Next
End Function

Function ActiveListObjectHeaderRowRangeRange(Optional wb, Optional ShName, Optional r)
'Возвращает ссылку на  регион заголовков  Активной Умной Таблицы
'(Активной Умной Таблицы в котором находится Активная ячейка)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(wb) Then Set wb = ActiveWorkbook
If IsMissing(r) Then Set r = Range(ActiveCell.address)

If IsMissing(ShName) Then ShName = r.Parent.name


For Each MyTable In wb.Worksheets(ShName).ListObjects
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
       Set ActiveListObjectHeaderRowRangeRange = MyTable.HeaderRowRange
        
        Exit Function
        Else
        Set ActiveListObjectHeaderRowRangeRange = Nothing
        
        End If
        Next
End Function

Function ActiveListObjectRow(Optional wb, Optional ShName, Optional r)
'Возвращает ссылку на строку Активной Умной Таблицы( строка  Активной Умной Таблицы в котором находится Активная ячейка)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(wb) Then Set wb = ActiveWorkbook
If IsMissing(r) Then Set r = Range(ActiveCell.address)

If IsMissing(ShName) Then ShName = r.Parent.name


For Each MyTable In wb.Worksheets(ShName).ListObjects
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
       Set ActiveListObjectRow = Application.Intersect(wb.Worksheets(ShName).Rows(r.Row), MyTable.Range)
        
        Exit Function
        Else
        Set ActiveListObjectRow = Nothing
        
        End If
        Next
End Function

Function ActiveListObjectColumnName(Optional wb, Optional ShName, Optional r, Optional name)
'Возвращает ссылку на столбец Активной Умной Таблицы( столбец  Активной Умной Таблицы в котором находится Активная ячейка)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(wb) Then Set wb = ActiveWorkbook
If IsMissing(r) Then Set r = Range(ActiveCell.address)

If IsMissing(ShName) Then ShName = r.Parent.name
'If IsMissing(Name) Then Name = "Отбор"

For Each MyTable In wb.Worksheets(ShName).ListObjects
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
        Set isect = Application.Intersect(wb.Worksheets(ShName).Columns(r.Column), MyTable.Range)
   If IsMissing(name) Then name = isect.Cells(1).value
       Set ActiveListObjectColumnName = MyTable.ListColumns(name)
        
        Exit Function
        Else
        Set ActiveListObjectColumnName = Nothing
        
        End If
        Next
End Function

Function ActiveListObjectColumnDataBodyRange(Optional wb, Optional ShName, Optional r)
'Возвращает ссылку на столбец региона данных Активной Умной Таблицы
'( столбец региона данных Активной Умной Таблицы в котором находится Активная ячейка)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(wb) Then Set wb = ActiveWorkbook
If IsMissing(r) Then Set r = Range(ActiveCell.address)

If IsMissing(ShName) Then ShName = r.Parent.name


For Each MyTable In wb.Worksheets(ShName).ListObjects
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
       Set ActiveListObjectColumnDataBodyRange = Application.Intersect(wb.Worksheets(ShName).Columns(r.Column), MyTable.DataBodyRange)
        
        Exit Function
        Else
        Set ActiveListObjectColumnDataBodyRange = Nothing
        
        End If
        Next
End Function
Function ActiveListRangeNoTotal(Optional wb, Optional ShName, Optional r)
'Возвращает ссылку на регион без итогов Активной Умной Таблицы( регион без итогов  Активной Умной Таблицы в котором находится Активная ячейка)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(wb) Then Set wb = ActiveWorkbook
If IsMissing(r) Then Set r = Range(ActiveCell.address)

If IsMissing(ShName) Then ShName = r.Parent.name


For Each MyTable In wb.Worksheets(ShName).ListObjects
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
       
       Set ActiveListRangeNoTotal = Union(MyTable.HeaderRowRange, MyTable.DataBodyRange)
        
        Exit Function
        Else
        Set ActiveListRangeNoTotal = Nothing
        
        End If
        Next
End Function
Function ActiveListObjectColumnConstans(Optional wb, Optional ShName, Optional q)
Dim New_tbl_Columns As Range
Dim r As ListColumn
Dim r_constasns
Dim v()
Dim New_tbl As ListObject
Dim isect As Range
On Error Resume Next
Set New_tbl = ActiveListObject()
If Not (New_tbl Is Nothing) Then
'Set New_tbl_Columns = New_tbl.HeaderRowRangeRange.Range
k = 0
For Each r In New_tbl.ListColumns
Set isect = Application.Intersect(r.Range, New_tbl.DataBodyRange)

Set r_constans = isect.SpecialCells(xlCellTypeFormulas)

If Err = 1004 Then

ReDim Preserve v(k)
v(k) = IndexColumn(ActiveSheet.name, New_tbl.name, r.name)
Err = 0

Else
Err = 0

End If
k = k + 1
Next
End If
ActiveListObjectColumnConstans = v
End Function

Sub ActiveListObjectColumnExcelFormula()
Dim lsc As ListColumn
Set lsc = ActiveListObjectColumnName()
myActiveListObjectColumnExcelFormula = "=" & lsc.Parent.name & "[" & lsc.name & "]"
MsgBox myActiveListObjectColumnExcelFormula
End Sub
Function ActiveListObjectColumnExcelFormulaДВССЫЛ(Optional lsc)

If IsMissing(lsc) Then Set lsc = ActiveListObjectColumnName()
ActiveListObjectColumnExcelFormulaДВССЫЛ = "=INDIRECT(" & """" & lsc.name & "[" & lsc.name & "]" & """" & ")"

End Function
Sub vvvv()
MsgBox FormulaДВССЫЛSource()
End Sub

Function FormulaДВССЫЛSource(Optional r)
Dim cl As clsVBScriptRegExp
Dim str As String
Set cl = New clsVBScriptRegExp
If IsMissing(lsc) Then Set lsc = ActiveListObjectColumnName()
lapki = Chr(34)
cl.myPattern = "[^" & lapki & "]+"


cl.myGlobal = True
If IsMissing(r) Then Set r = ActiveCell
str = ValidationFormula(r)
Set objMatches = cl.objRegExpReplaceExecute(str)
Set objMatch = objMatches.Item(1)
    FormulaДВССЫЛSource = objMatch.value '& ", " & "FirstIndex=" & objMatch.FirstIndex & ", " & "Length=" & objMatch.Length


End Function



Sub ActiveListObjectColumnAddValidadationList()
myFormula = ActiveListObjectColumnExcelFormulaДВССЫЛ()
Call ValidationListAdd(ActiveCell, myFormula)
End Sub





Sub DetermineActiveTable()

Dim SelectedCell As Range
Dim TableName As String
Dim ActiveTable As ListObject

Set SelectedCell = ActiveCell

'Determine if ActiveCell is inside a Table
  On Error GoTo NoTableSelected
    TableName = SelectedCell.ListObject.name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
  On Error GoTo 0

'Do something with your table variable (ie Add a row to the bottom of the ActiveTable)
  ActiveTable.ListRows.Add AlwaysInsert:=True
  
Exit Sub

'Error Handling
NoTableSelected:
  MsgBox "There is no Table currently selected!", vbCritical

End Sub

Sub fofof()
Dim New_tbl
Set New_tbl = ActiveListObjectHeaderRowRangeRange()
If Not (New_tbl Is Nothing) Then New_tbl.Select
End Sub

Sub myListObjectАddSelection()
Dim lo As ListObject
Dim sh As Worksheet, shlog As Worksheet
Dim LoCol As ListColumn
Call myListObjectaddIdColumn(Selection)
End Sub
Sub myListObjectaddIdColumn(HederRange As Range)
Dim lo As ListObject
Dim sh As Worksheet, shlog As Worksheet
Dim LoCol As ListColumn
Dim r As Range
Dim myNameTab As String
Set shlog = SheetExistBookCreate(ActiveWorkbook, "log", True)
shlog.Cells(1, 1) = "Имя"
shlog.Cells(1, 2) = "Успех"
k = 2
For Each r In HederRange.Cells
myNameTab = r.value
If myNameTab = "" Then
shlog.Cells(k, 1) = myNameTab
shlog.Cells(k, 2) = "Имя столбца не может быть пустым"
myCreate = False
Else
myCreate = True
End If
 If RangeExists(myNameTab) Then
 shlog.Cells(k, 1) = myNameTab
 If shlog.Cells(k, 2) <> "" Then shlog.Cells(k, 2) = shlog.Cells(k, 2) & ","
shlog.Cells(k, 2) = shlog.Cells(k, 2) & "Именованый диапазон или формула с таким именем уже существует в книге"
myCreate = False
Else
myCreate = True
End If
 
If SheetExist(myNameTab) Then
shlog.Cells(k, 1) = myNameTab
If shlog.Cells(k, 2) <> "" Then shlog.Cells(k, 2) = shlog.Cells(k, 2) & ","
shlog.Cells(k, 2) = shlog.Cells(k, 2) & "Лист с таким именем уже существует в книге"
myCreate = False
Else
myCreate = True
End If
If myCreate Then
Set sh = SheetExistBookCreate(ActiveWorkbook, myNameTab, True)
End If
If ListObjectExist(ActiveWorkbook, myNameTab) Then
shlog.Cells(k, 1) = myNameTab
If shlog.Cells(k, 2) <> "" Then shlog.Cells(k, 2) = shlog.Cells(k, 2) & ","
shlog.Cells(k, 2) = shlog.Cells(k, 2) & "Таблица с таким именем уже существует в книге"
myCreateTab = False
Else
myCreateTab = True
End If
If myCreate And myCreateTab Then
Set lo = myListObjectadd(ActiveWorkbook, myNameTab, myNameTab, [A1], True)
lo.ListColumns(1).name = "Код"
Set LoCol = ListObjectColumnAdd(lo, myNameTab)
shlog.Cells(k, 1) = myNameTab
If shlog.Cells(k, 2) <> "" Then shlog.Cells(k, 2) = shlog.Cells(k, 2) & ","
shlog.Cells(k, 2) = shlog.Cells(k, 2) & "Таблица успешно созданы"




End If
k = k + 1
Next
shlog.Activate
End Sub






Sub myListObjectaddActiveSheetName()
Dim lo As ListObject
Dim sh As Worksheet
Dim LoCol As ListColumn
Dim mySelect As Range
myName = ActiveCell.value
Set sh = ActiveSheet

TableName = VBA.Replace(VBA.Trim(sh.name), " ", "_")
If ListObjectExist(TableName) Then
Exit Sub
End If
Set mySelect = Selection
mySelectCount = mySelect.Cells.Count
If mySelectCount = 1 Then
Set startRange = mySelect

Set EndRange = mySelect.CurrentRegion.Cells(mySelect.CurrentRegion.Cells.Count).Offset(rowoffset:=1)
Else
Set startRange = mySelect.Cells(1)
Set EndRange = mySelect.Cells(mySelectCount)
End If
Set v = Range(startRange, EndRange)


Set lo = myListObjectadd(ActiveWorkbook, sh.name, TableName, v, xlYes)

End Sub
Function ListObjectExist(wb As Workbook, name)
    
    Dim ws As Worksheet
    Dim lstList As ListObject
    ListObjectExist = False
    For Each ws In wb.Worksheets
    
    For Each lstList In ws.ListObjects
        If lstList.name = name Then
          ListObjectExist = True
            Exit For
        End If
    Next
    Next
    
    
End Function
Function RangeExists(s As String) As Boolean
    On Error GoTo Nope
    RangeExists = Range(s).Count > 0
Nope:
End Function
