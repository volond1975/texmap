Attribute VB_Name = "Module2"




Public Sub exceljson()
Dim http As Object, JSON As Object, i As Integer
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "http://jsonplaceholder.typicode.com/users", False
http.Send
Set JSON = ParseJson(http.ResponseText)
i = 2
For Each Item In JSON
Sheets(1).Cells(i, 1).value = Item("id")
Sheets(1).Cells(i, 2).value = Item("name")
Sheets(1).Cells(i, 3).value = Item("username")
Sheets(1).Cells(i, 4).value = Item("email")
Sheets(1).Cells(i, 5).value = Item("address")("city")
Sheets(1).Cells(i, 6).value = Item("phone")
Sheets(1).Cells(i, 7).value = Item("website")
Sheets(1).Cells(i, 8).value = Item("company")("name")
i = i + 1
Next
MsgBox ("complete")
End Sub
Public Sub exceljson1()
Dim http As Object, JSON As Object, i As Integer, ndr As name
Dim w As Range
Dim wb As Workbook
'Set http = CreateObject("MSXML2.XMLHTTP")
'http.Open "GET", "http://jsonplaceholder.typicode.com/users", False
'http.Send
param = "offset"
ndrName = "gРГК_кошторис"
Set JSON = ParseJson(ПоИмениJSON(ndrName))
Set wb = mywbBook("Комплект.xlsm", ThisWorkbook.Path & "\")
If wb Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & "Комплект.xlsm")

i = 2
For Each Item In JSON
For Each Jtem In Item.Keys
Dim f As Variant
If Not IsEmpty(Jtem) Then
f = Split(Jtem, "_")
If param = "offset" Then Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
If param = "address" Then Set w = Sheets("Лист8").Range(Jtem)
'fixme
'If param = "NameColumn" Then myitem(firstRange.Offset(0, j - 1).address) = firstRange.Offset(i - 1, j - 1).value


Set w = Sheets("Лист8").Range("A4").Offset(Val(f(0)), Val(f(1)))
'Sheets("Лист8").Range(Jtem).value = Item(Jtem)
w.value = Item(Jtem)
End If
'Sheets(1).Cells(i, 2).value = Item("name")
'Sheets(1).Cells(i, 3).value = Item("username")
'Sheets(1).Cells(i, 4).value = Item("email")
'Sheets(1).Cells(i, 5).value = Item("address")("city")
'Sheets(1).Cells(i, 6).value = Item("phone")
'Sheets(1).Cells(i, 7).value = Item("website")
'Sheets(1).Cells(i, 8).value = Item("company")("name")

Next
i = i + 1
Next
MsgBox ("complete")
End Sub

Public Sub exceltojson()

Dim rng As Range, items As New Collection, myitem As New Dictionary, i As Integer, cell As Variant

Set rng = Range("A2:A3")

'Set rng = Range(Sheets(2).Range("A2"), Sheets(2).Range("A2").End(xlDown)) use this for dynamic range

i = 0

For Each cell In rng

Debug.Print (cell.value)
myitem("name") = cell.value

myitem("email") = cell.Offset(0, 1).value

myitem("phone") = cell.Offset(0, 2).value

items.Add myitem

Set myitem = Nothing

i = i + 1

Next

Sheets(1).Range("A4").value = ConvertToJson(items, Whitespace:=2)

End Sub
Public Function ToJSON(rng As Range) As String
'http://niraula.com/blog/convert-excel-data-json-format-using-vba/
    ' Make sure there are two columns in the range
    If rng.Columns.Count < 2 Then
        ToJSON = CVErr(xlErrNA)
        Exit Function
    End If
 
    Dim dataLoop, headerLoop As Long
    ' Get the first row of the range as a header range
    Dim headerRange As Range: Set headerRange = Range(rng.Rows(1).Address)
    
    ' We need to know how many columns are there
    Dim colCount As Long: colCount = headerRange.Columns.Count
    
    Dim JSON As String: JSON = "["
    
    For dataLoop = 1 To rng.Rows.Count
        ' Skip the first row as it's been used as a header
        If dataLoop > 1 Then
            ' Start data row
            Dim rowJson As String: rowJson = "{"
            
            ' Loop through each column and combine with the header
            For headerLoop = 1 To colCount
                rowJson = rowJson & """" & headerRange.Value2(1, headerLoop) & """" & ":"
                rowJson = rowJson & """" & rng.Value2(dataLoop, headerLoop) & """"
                rowJson = rowJson & ","
            Next headerLoop
            
            ' Strip out the last comma
            rowJson = Left(rowJson, Len(rowJson) - 1)
            
            ' End data row
            JSON = JSON & rowJson & "},"
        End If
    Next
    
    ' Strip out the last comma
    JSON = Left(JSON, Len(JSON) - 1)
    
    JSON = JSON & "]"
    
    ToJSON = JSON
End Function
Option Explicit
Public Sub GetInfoFromSheet()
    Dim jsonStr As String
    jsonStr = [A1]                               '<== read in from sheet
    Dim json As Object
    Set json = JsonConverter.ParseJson(jsonStr)

    Dim i As Long, j As Long, key As Variant
    For i = 1 To json.Count
        For Each key In json(i).Keys
            Select Case key
            Case "name", "type"
                Debug.Print key & " " & json(i)(key)
            Case Else
                Select Case TypeName(json(i)(key))
                Case "Dictionary"
                    Dim key2 As Variant
                    For Each key2 In json(i)(key)
                        Select Case TypeName(json(i)(key)(key2))
                        Case "Collection"
                            Dim k As Long
                            For k = 1 To json(i)(key)(key2).Count
                                Debug.Print key & " " & key2 & " " & json(i)(key)(key2)(k)
                            Next k
                        Case Else
                            Debug.Print key & " " & key2 & " " & json(i)(key)(key2)
                        End Select
                    Next key2
                Case "Collection"
                    For j = 1 To json(i)(key).Count '<== "actions"
                        Dim key3 As Variant
                        For Each key3 In json(i)(key)(j).Keys
                            Select Case TypeName(json(i)(key)(j)(key3))
                            Case "String", "Boolean", "Double"
                                Debug.Print key & " " & key3 & " " & json(i)(key)(j)(key3)
                            Case Else
                                Dim key4 As Variant
                                For Each key4 In json(i)(key)(j)(key3).Keys
                                    Debug.Print key & " " & key3 & " " & key4 & " " & json(i)(key)(j)(key3)(key4)
                                Next key4
                            End Select
                        Next key3
                    Next j
                Case Else
                    Debug.Print key & " " & json(i)(key)
                End Select
            End Select
        Next key
    Next i
End Sub
