Attribute VB_Name = "JSONListObject"
Function JsonTextСтрока(ztxt)
Attribute JsonTextСтрока.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос4 Макрос
'
Dim wb As Workbook
Dim sh As Worksheet
Dim lo As ListObject
Dim txt As String, JsonText As String
Dim r As Range
Dim lrow  As ListRow
Dim locol As ListColumn
Set wb = ThisWorkbook
Set lo = Worksheets("Наряд").ListObjects("tableJsonEOD")

Dim dict As New Dictionary
dict.CompareMode = CompareMethod.TextCompare
spztxt = Split(ztxt, ",")
'Set Dict("0") = New Dictionary
'Dict("0").Add "Number", spztxt(0)
'Dict("0").Add "Date", spztxt(1)
For i = 1 To lo.ListRows.Count
If Intersect(lo.ListRows(i).Range, lo.ListColumns("№").DataBodyRange).value <> "" Then
Set dict(i) = New Dictionary
For Each r In lo.HeaderRowRange.Cells

dict(i).Add r.value, Intersect(lo.ListRows(i).Range, lo.ListColumns(r.value).DataBodyRange).value

Next
End If
Next i
JsonText = ConvertToJson(dict)

Debug.Print URLDecode(JsonText)
JsonTextСтрока = JsonText
'Dict ("A") ' -> Empty
'Dict("A") = 123
'Dict ("A") ' -> = Dict.Item("A") = 123
'Dict.Exists "A" ' -> True

'Dict.Add "A", 456
' -> Throws 457: This key is already associated with an element of this collection

' Both Set and Let work
'Set Dict("B") = New Dictionary
'Dict("B").Add "Inner", "Value"
'Dict("B")("Inner") ' -> "Value"
'
'UBound(Dict.Keys) ' -> 1
'UBound(Dict.Items) ' -> 1
'
'' Rename key
'Dict.Key("B") = "C"
'Dict.Exists "B" ' -> False
'Dict("C")("Inner") ' -> "Value"
'
'' Trying to remove non-existant key throws 32811
'Dict.Remove "B"
'' -> Throws 32811: Application-defined or object-defined error
'
'' Trying to change CompareMode when there are items in the Dictionary throws 5
'Dict.CompareMode = CompareMethod.BinaryCompare
'' -> Throws 5: Invalid procedure call or argument
'
'Dict.Remove "A"
'Dict.RemoveAll
'
'Dict.Exists "A" ' -> False
'Dict ("C") ' -> Empty
'
End Function
Sub TestConvertListObjectToJson()
Dim jsonFileObject As New FileSystemObject
Dim jsonFileExport As TextStream
Set dict = New Dictionary
'Set Temp = ConvertListObjectToJson("fulllishoz")
Set dict("fulllishoz") = ConvertListObjectToDict("fulllishoz")("fulllishoz")
Set dict("fulllink_small") = ConvertListObjectToDict("fulllink_small")("fulllink_small")
mText = URLDecode(RussianStringToURLEncode_New(UTF8_Decode(ConvertToJson(dict))))
'mText = URLDecode(RussianStringToURLEncode_New(UTF8_Decode(ConvertToJson(ConvertListObjectToJson("fulllishoz")))))
'MsgBox Mid(mText, 2, Len(mText) - 2)
'MsgBox mText
'mText2 = URLDecode(RussianStringToURLEncode_New(UTF8_Decode(ConvertToJson(ConvertListObjectToJson("fulllink_small")))))
'zText = UTF8_Decode(mText)
'SetClipboard (URLDecode(zText))
'MsgBox URLDecode(RussianStringToURLEncode_New(zText))
Set wbMy = ThisWorkbook
sFileName = Replace(wbMy.name, ".xlsm", ".json")
sPath = wbMy.Path & "\" & sFileName
'Open sPath For Output As #1
'
'Print #1, URLDecode(zText)
'
'Close #1
Set jsonFileExport = jsonFileObject.CreateTextFile(sPath, True)
'jsonFileExport.WriteLine (Mid(URLDecode(zText), 2, Len(URLDecode(zText)) - 2))
jsonFileExport.WriteLine (mText)
End Sub





Function ConvertListObjectToJson(Optional ListObjectName, Optional wb As Workbook)
'
' Макрос4 Макрос
', Optional KeyColumnName = "ID"
'Dim wb As Workbook
Dim sh As Worksheet
Dim lo As ListObject
Dim lj As New clsmListObjs
Dim txt As String, JsonText As String
Dim r As Range
Dim lrow  As ListRow
Dim locol As ListColumn
If Not IsMissing(wb) Then Set wb = ActiveWorkbook
With lj
.Initialize wb

If IsMissing(ListObjectName) Then
Set lo = .ActiveListObject
ListObjectName = lo.name
Else
Set lo = .items(ListObjectName)
End If
Set Nastr = ParseJson(lo.Comment)
Dim dict As New Dictionary
dict.CompareMode = CompareMethod.TextCompare
Set dict(ListObjectName) = New Dictionary
For i = 1 To lo.ListRows.Count
z = Nastr("keycolumnname")
KeyName = Intersect(lo.ListRows(i).Range, lo.ListColumns(URLDecode(RussianStringToURLEncode_New(Nastr("keycolumnname")))).DataBodyRange).value
Set dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)) = New Dictionary
For Each r In lo.HeaderRowRange.Cells
Set zn = Intersect(lo.ListRows(i).Range, lo.ListColumns(r.value).DataBodyRange)
If IsEmpty(zn.value) Then q = Null Else q = zn.value
dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)).Add r.value, q

Next
Next i
JsonText = ConvertToJson(dict)

ConvertListObjectToJson = JsonText

End With
End Function

Function ConvertListObjectToDict(Optional ListObjectName, Optional wb As Workbook) As Object
'
' Макрос4 Макрос
', Optional KeyColumnName = "ID"
'Dim wb As Workbook
Dim sh As Worksheet
Dim lo As ListObject
Dim lj As New clsmListObjs
Dim txt As String, JsonText As String
Dim r As Range
Dim lrow  As ListRow
Dim locol As ListColumn
If Not IsMissing(wb) Then Set wb = ActiveWorkbook
With lj
.Initialize wb

If IsMissing(ListObjectName) Then
Set lo = .ActiveListObject
ListObjectName = lo.name
Else
Set lo = .items(ListObjectName)
End If
Set Nastr = ParseJson(lo.Comment)
Dim dict As New Dictionary
dict.CompareMode = CompareMethod.TextCompare
'spztxt = Split(ztxt, ",")
'Set Dict("0") = New Dictionary
'Dict("0").Add "Number", spztxt(0)
'Dict("0").Add "Date", spztxt(1)
Set dict(ListObjectName) = New Dictionary
For i = 1 To lo.ListRows.Count
'If Intersect(lo.ListRows(i).Range, lo.ListColumns("№").DataBodyRange).value <> "" Then
z = Nastr("keycolumnname")
KeyName = Intersect(lo.ListRows(i).Range, lo.ListColumns(URLDecode(RussianStringToURLEncode_New(Nastr("keycolumnname")))).DataBodyRange).value
Set dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)) = New Dictionary
For Each r In lo.HeaderRowRange.Cells
'Set KeyColumn = lo.ListColumns(r.value).DataBodyRange
Set zn = Intersect(lo.ListRows(i).Range, lo.ListColumns(r.value).DataBodyRange)
If IsEmpty(zn.value) Then q = Null Else q = zn.value
dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)).Add r.value, q

Next
'End If
Next i

End With
Set ConvertListObjectToDict = dict
End Function

Function ConvertListRowsToDict(Optional ListObjectName, Optional wb As Workbook)
'
' Макрос4 Макрос
', Optional KeyColumnName = "ID"
'Dim wb As Workbook
Dim sh As Worksheet
Dim lo As ListObject
Dim lj As New clsmListObjs
Dim txt As String, JsonText As String
Dim r As Range
Dim lrow  As ListRow
Dim locol As ListColumn
If Not IsMissing(wb) Then Set wb = ActiveWorkbook
With lj
.Initialize wb

If IsMissing(ListObjectName) Then
Set lo = .ActiveListObject
ListObjectName = lo.name
Else
Set lo = .items(ListObjectName)
End If
Set Nastr = ParseJson(lo.Comment)
z = Nastr("keycolumnname")


Dim dict As New Dictionary
dict.CompareMode = CompareMethod.TextCompare
'spztxt = Split(ztxt, ",")
'Set Dict("0") = New Dictionary
'Dict("0").Add "Number", spztxt(0)
'Dict("0").Add "Date", spztxt(1)
Set dict(ListObjectName) = New Dictionary
For i = 1 To lo.ListRows.Count
'If Intersect(lo.ListRows(i).Range, lo.ListColumns("№").DataBodyRange).value <> "" Then

KeyName = Intersect(lo.ListRows(i).Range, lo.ListColumns(URLDecode(RussianStringToURLEncode_New(Nastr("keycolumnname")))).DataBodyRange).value
Set dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)) = New Dictionary
For Each r In lo.HeaderRowRange.Cells
'Set KeyColumn = lo.ListColumns(r.value).DataBodyRange
Set zn = Intersect(lo.ListRows(i).Range, lo.ListColumns(r.value).DataBodyRange)
If IsEmpty(zn.value) Then q = Null Else q = zn.value
dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)).Add r.value, q

Next
'End If
Next i

End With
End Function















Sub Перечесление()
Attribute Перечесление.VB_ProcData.VB_Invoke_Func = " \n14"
'Требует JsonConverter
' Макрос6 Макрос
'
Dim v As Collection
Set v = ParseJson(ActiveCell)
q = collectionToArray(v)
For i = 0 To UBound(q)
'If i = v.Count Then
TextList = TextList & q(i) & ";"
TextList = Mid(TextList, 1, Len(TextList) - 1)
Next
End Sub
Sub МокраяПечатьПодложка()
Attribute МокраяПечатьПодложка.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос7 Макрос
Set sh = Worksheets("Наряд")
If sh.PageSetup.LeftFooterPicture.fileName = "" Then
sh.PageSetup.LeftFooterPicture.fileName = _
       "E:\1S_bases\Зарплата Электронный Облик\НАРЯДИ\Печать.png"
       sh.PageSetup.LeftFooter = "&G"
       sh.PrintPreview
'  sh.Shapes("TextBox 40").ShapeRange(1).TextFrame2.TextRange.Characters(1, 11).Font.Fill.ForeColor.RGB = RGB(112, 48, 160)
 With sh.Shapes("TextBox 40").TextFrame2.TextRange.Characters(1, 11).Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
        .Solid
    End With
      
       
       Else
     sh.PageSetup.LeftFooterPicture.fileName = ""
     sh.PageSetup.LeftFooter = ""
'     sh.Shapes("TextBox 40").TextFrame2.TextRange.Characters(1, 11).Font.Fill
     With sh.Shapes("TextBox 40").TextFrame2.TextRange.Characters(1, 11).Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
       End If
'
'    ActiveSheet.PageSetup.LeftFooterPicture.Filename = _
'        "E:\1S_bases\Зарплата Электронный Облик\НАРЯДИ\Печать.png"
'    Application.PrintCommunication = False
   
'ActiveSheet.PageSetup.LeftFooter = "&G"
'        .LeftHeader = ""
'        .CenterHeader = ""
'        .RightHeader = ""
'
'        .CenterFooter = ""
'        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0.590551181102362)
'        .RightMargin = Application.InchesToPoints(0.196850393700787)
'        .TopMargin = Application.InchesToPoints(0.196850393700787)
'        .BottomMargin = Application.InchesToPoints(0.196850393700787)
'        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
'        .FooterMargin = Application.InchesToPoints(0.511811023622047)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .CenterHorizontally = False
'        .CenterVertically = False
'        .Orientation = xlLandscape
'        .Draft = False
'        .PaperSize = xlPaperA4
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = True
'        .Zoom = 80
'        .PrintErrors = xlPrintErrorsDisplayed
'        .OddAndEvenPagesHeaderFooter = False
'        .DifferentFirstPageHeaderFooter = False
'        .ScaleWithDocHeaderFooter = True
'        .AlignMarginsHeaderFooter = False
'        .EvenPage.LeftHeader.Text = ""
'        .EvenPage.CenterHeader.Text = ""
'        .EvenPage.RightHeader.Text = ""
'        .EvenPage.LeftFooter.Text = ""
'        .EvenPage.CenterFooter.Text = ""
'        .EvenPage.RightFooter.Text = ""
'        .FirstPage.LeftHeader.Text = ""
'        .FirstPage.CenterHeader.Text = ""
'        .FirstPage.RightHeader.Text = ""
'        .FirstPage.LeftFooter.Text = ""
'        .FirstPage.CenterFooter.Text = ""
'        .FirstPage.RightFooter.Text = ""
'    End With
'    Application.PrintCommunication = True
End Sub
Sub Макрос8()
Attribute Макрос8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос8 Макрос
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$A$1:$W$54"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.590551181102362)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.196850393700787)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = True
        .Zoom = 80
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
End Sub
Sub Макрос9()
Attribute Макрос9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос9 Макрос
'

'
    With ActiveSheet.Shapes.Range(Array("TextBox 40")).ShapeRange(1).TextFrame2.TextRange.Characters(1, 11).Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
        .Solid
    End With
End Sub
Sub Макрос10()
Attribute Макрос10.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос10 Макрос
'

'
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 11).Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
End Sub
Sub ConvertRasxtnToJson()
Dim jsonFileObject As New FileSystemObject
Dim jsonFileExport As TextStream
Set dict = New Dictionary
temp = URLDecode(ConvertListObjectToJson("rachtet"))
'Set dict("rachtet") = ConvertListObjectToDict("rachtet")("rachtet")
'Set dict("fulllink_small") = ConvertListObjectToDict("fulllink_small")("fulllink_small")
mText = URLDecode(RussianStringToURLEncode_New(UTF8_Decode(ConvertToJson(dict))))
'mText = URLDecode(RussianStringToURLEncode_New(UTF8_Decode(ConvertToJson(ConvertListObjectToJson("fulllishoz")))))
'MsgBox Mid(mText, 2, Len(mText) - 2)
MsgBox temp
'mText2 = URLDecode(RussianStringToURLEncode_New(UTF8_Decode(ConvertToJson(ConvertListObjectToJson("fulllink_small")))))
'zText = UTF8_Decode(mText)
'SetClipboard (URLDecode(zText))
'MsgBox URLDecode(RussianStringToURLEncode_New(zText))
'Set wbMy = ThisWorkbook
'sFileName = Replace(wbMy.Name, ".xlsm", ".json")
'sPath = wbMy.Path & "\" & sFileName
''Open sPath For Output As #1
''
''Print #1, URLDecode(zText)
''
''Close #1
'Set jsonFileExport = jsonFileObject.CreateTextFile(sPath, True)
''jsonFileExport.WriteLine (Mid(URLDecode(zText), 2, Len(URLDecode(zText)) - 2))
'jsonFileExport.WriteLine (mText)
End Sub
Sub TestConvertJsonToListObject()
t = ConvertJsonToListObject(ActiveCell.value, "rachtet")
End Sub
Function ConvertJsonToListObject(JsonText, Optional ListObjectName, Optional wb As Workbook)
'
' Макрос4 Макрос
', Optional KeyColumnName = "ID"
'Dim wb As Workbook
Dim sh As Worksheet
Dim lo As ListObject
Dim lj As New clsmListObjs
Dim txt As String
Dim r As Range
Dim lrow  As ListRow
Dim locol As ListColumn
If Not IsMissing(wb) Then Set wb = ActiveWorkbook
With lj
.Initialize wb
Set dict = ParseJson(JsonText)
If IsMissing(ListObjectName) Then
Set lo = .ActiveListObject
ListObjectName = lo.name
Else
Set lo = .items(ListObjectName)
End If
Set Nastr = ParseJson(lo.Comment)
z = Nastr("keycolumnname")
'Dim dict As New Dictionary
'dict.CompareMode = CompareMethod.TextCompare
'spztxt = Split(ztxt, ",")
'Set Dict("0") = New Dictionary
'Dict("0").Add "Number", spztxt(0)
'Dict("0").Add "Date", spztxt(1)
'Set dict(ListObjectName) = New Dictionary

For Each t In dict(ListObjectName).Keys
For Each w In dict(ListObjectName)(t).Keys
Set p = .ValueListObject(ListObjectName, z, w, t)
lo.parent.Range(Replace(p.FormulaLocal, "=", "")).value = dict(ListObjectName)(t)(w)
Next
'Set r = dict(ListObjectName)(t)
Next
'For i = 1 To lo.ListRows.Count
'
''If Intersect(lo.ListRows(i).Range, lo.ListColumns("№").DataBodyRange).value <> "" Then
'z = Nastr("keycolumnname")
'KeyName = Intersect(lo.ListRows(i).Range, lo.ListColumns(URLDecode(RussianStringToURLEncode_New(Nastr("keycolumnname")))).DataBodyRange).value
'Set dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)) = New Dictionary
'For Each r In lo.HeaderRowRange.Cells
''Set KeyColumn = lo.ListColumns(r.value).DataBodyRange
'Set zn = Intersect(lo.ListRows(i).Range, lo.ListColumns(r.value).DataBodyRange)
'If IsEmpty(zn.value) Then q = Null Else q = zn.value
'dict(ListObjectName)(RussianStringToURLEncode_New(KeyName)).Add r.value, q
'
'Next
''End If
'Next i
''JsonText = ConvertToJson(dict)
'
''Debug.Print URLDecode(JsonText)
'ConvertListObjectToJson = JsonText

End With
End Function
