Attribute VB_Name = "Module4"
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









Sub NamedRange_Loop(Optional wb As Workbook, Optional sh As Worksheet)
'PURPOSE: Delete all Named Ranges in the Active Workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim nm As name
If Not wb Is missing Then wb = ActiveWorkbook
'Loop through each named range in workbook
  For Each nm In wb.Names
    Debug.Print nm.name, nm.RefersTo
  Next nm
  
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
