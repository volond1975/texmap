Attribute VB_Name = "������������"

'����� �������
    'myListObjectadd-��������� � ���������� ������ �� ����� �������
'������� ����� �������
    'ListObjectColumnAdd-��������� ������� � ������ LCName � ����� �������
    'ListObjectColumnRENAME-��������������� ������� ����� �������
'������� � ������� ����� �������
'    ListObjectColumnFormulaLocal
'    ListObjectColumnFormulaR1C1
'    ListObjectColumnFormula-�������� �������� �� ������� �������� ������ ���
'    MAXListTableColumn()-���������� ������������ �������� ������� ����� �������


'HeaderRowRangeStilNone-������� ������� ��������� ����� ������� � ���� ������ �������




'----------------------------------------------------------------------------------------------------
'����� �������
Function myListObjectadd(wb, NameSheet, NameListObject, r, Optional zag) As ListObject

Dim lo As ListObject
If IsMissing(zag) Then zag = 1
Set lo = wb.Worksheets(NameSheet).ListObjects.Add(xlSrcRange, r, , zag)
'ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$E$2"), , xlYes)
'Else
'Set lo = wb.Worksheets(NameSheet).ListObjects.Add(xlSrcRange, r, , xlNo)
'End If

lo.name = _
        NameListObject
', ListObjectTableStyle
'        LO.TableStyle = ListObjectTableStyle
        Set myListObjectadd = lo
End Function





Function myListObject(wb As Workbook, NameSheet, NameListObject)
Dim lo As ListObject
On Error Resume Next
Set lo = wb.Worksheets(NameSheet).ListObjects(NameListObject)
If VBA.Err = 0 Then
Set myListObject = lo
Else
Set myListObject = Nothing
End If

End Function
Function ListObjectExist(wb As Workbook, NameSheet, NameListObject)
Dim lo As ListObject
On Error Resume Next
Set lo = wb.Worksheets(NameSheet).ListObjects(NameListObject)
If VBA.Err = 0 Then ListObjectExist = True _
Else: ListObjectExist = False
End Function

'-----------------------------------------------------------------
'������� ����� �������
'-------------------------------------------------------------------
Function ListObjectColumnAdd(lo As ListObject, LCName)
'��������� ������� � ������ LCName � ����� �������
Dim lcs As ListColumn
Set lcs = lo.ListColumns.Add
     lo.HeaderRowRange(lcs.Index) = LCName
Set ListObjectColumnAdd = lcs
End Function
Function ListObjectColumnCount(��������, ����������)
'��������� ������� � ������ LCName � ����� �������
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)


   ListObjectColumnCount = ls.HeaderRowRange.Columns.Count
End Function



Function ListObjectColumnAddFormulaOrValue(lo As ListObject, LCName As String, formul As String, form_at As String, v As Boolean) As ListColumn
'��������� ������� � ������ LCName � ����� �������,��������� ������� Formul,����������� form_at,� ��� v As true ���������� � ��������
Dim lcs As ListColumn
Set lcs = ListObjectColumnAdd(lo, LCName)
    Call ListObjectColumnFormula(lcs, formul, form_at)
  If v Then
  For Each r In lcs.Range.Cells
  r.value = r.value
  Next
End If
Set ListObjectColumnAddFormulaOrValue = lcs
End Function



Function ListObjectColumnRENAME(lo As ListObject, LCName, NewName)
'��������������� ������� ����� �������
Dim lcs As ListColumn
Set lcs = lo.ListColumns(LCName)
     lo.HeaderRowRange(lcs.Index) = NewName
Set ListObjectColumnRENAME = lcs
End Function

'--------------------------------------------------------------------------------------

'������� � ������� ����� �������
'---------------------------------------------------------------------------------------
Sub ListObjectColumnFormulaLocal(lcs As ListColumn, myFormulaLocal As String, Optional NumbFormat)
Dim lo As ListObject
Dim r As Range
Set lo = lcs.Parent
Set r = Intersect(lo.DataBodyRange, lcs.Range)
       r.FormulaLocal = myFormulaLocal
       r.NumberFormat = NumbFormat
       
End Sub
Sub ListObjectColumnFormulaR1C1(lcs As ListColumn, myFormulaR1C1, Optional NumbFormat)
Dim lo As ListObject
Dim r As Range
Set lo = lcs.Parent
Set r = Intersect(lo.DataBodyRange, lcs.Range)
       r.FormulaR1C1 = myFormulaR1C1
       r.NumberFormat = NumbFormat
       
End Sub

Sub ListObjectColumnFormula(lcs As ListColumn, myFormula, Optional NumbFormat)
Dim lo As ListObject
Dim r As Range
Set lo = lcs.Parent
Set r = Intersect(lo.DataBodyRange, lcs.Range)
       r.Formula = myFormula
       r.NumberFormat = NumbFormat
       
End Sub

Function ListObjectColumnExcelFormula(lcs As ListColumn)
ListObjectColumnFormula "=" & lsc.Parent.name & "[" & lsc.name & "]"
End Function

'---------------------------------------------------------------------------------------------------

Function MAXListTableColumn(��������, ����������, ����������)
'���������� ������������ �������� ������� ����� �������

MAXListTableColumn = Application.WorksheetFunction.max(���������������������������������(��������, ����������, ����������))
End Function

Function ���������������������������������(��������, ����������, ����������) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������)
Set ��������������������������������� = Range(lsc.Range.Cells(2), lsc.Range.Cells(lsc.Range.Cells.Count))


End Function
Function ��������������������������������������������(��������, ����������, ����������) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������)
Set r = lsc.DataBodyRange
Set �������������������������������������������� = r


End Function














Function ��������������������(��������, ����������, ����������������, ������������������, ��������������) As Range
'��������, ����������, ����������������, ������������������,��������������
'"�����","�����","�����","�������� � �������","Գ�����"
' ����� ������������ �������
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������������)
Set lscf = ls.ListColumns(������������������)
Set n = lsc.Range.Find(��������������, LookIn:=xlValues, LookAt:=xlWhole)
If Not n Is Nothing Then
Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set �������������������� = r
Else
Set �������������������� = Nothing
End If
End Function
Function ��������������������V(��������, ����������, ����������������, ������������������, ��������������)
'��������, ����������, ����������������, ������������������,��������������
'"�����","�����","�����","�������� � �������","Գ�����"
' ����� ������������ �������
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������������)
v = ������������������
Set lscf = ls.ListColumns(v)
Set n = lsc.Range.Find(��������������, LookIn:=xlValues, LookAt:=xlWhole)
If Not n Is Nothing Then
Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
��������������������V = r.value
Else
��������������������V = ""
End If
End Function














'Sub ���������������������Substring(��������, ����������, ����������������, ������������������, ��������������, Delimiter, n)
'Dim txt
'Set txt = ��������������������(��������, ����������, ����������������, ������������������, ��������������)
'���������������������Substring = Substring(txt, Delimiter, n)
''����� �������� ������� �� �������
'� = ���������������������Substring("���������", ���������, ����������������, ������������������, ��������������, Delimiter, n)
'
'
'm = Split(r.Offset(columnoffset:=1).value, ",")
'Glav.ComboBox_���.Text = VBA.Trim(m(1))
''��������� �������� ������� �� �������
'Glav.cbB.Text = m(2)
'm = Split(r.Offset(columnoffset:=1).value, " ")
''���� �������� ������� �� �������
'Glav.TextBox_������.Text = m(3)
''�������  �� �������
'Glav.ComboBox_������.Text = r.Offset(columnoffset:=2).value
'End Sub
















Function �������������������������(��������, ����������, ����������������, ������������������, ��������������, �������� As Boolean) As Range
'��������, ����������, ����������������, ������������������,��������������
'"�����","�����","�����","�������� � �������","Գ�����"
' ����� ������������ �������
'1 �������� �����������
'2 msgbox

Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������������)
Set lscf = ls.ListColumns(������������������)
Set n = lsc.Range.Find(��������������, LookAt:=xlWhole)

If Not n Is Nothing Then
Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set ������������������������� = r
Else
    If �������� Then
    Set n = ������������������(��������, ����������, ����������������, ��������������)
    Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set ������������������������� = r
    Else
    Select Case MsgBox("�������� �� ������� ! ������ � ����������", vbOKCancel Or vbCritical Or vbDefaultButton1, "�������� �� �������")
    
        Case vbOK
     Set n = ������������������(��������, ����������, ����������������, ��������������)
    Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lscf.Range, lsr.Range)
Set ������������������������� = r
        Case vbCancel
    Set ������������������������� = Nothing
    End Select
    End If

End If
End Function
Function ��������������������������������������(��������, ����������, lsr As ListRow, ����������, ��������) As Range
'��������, ����������, ����������������, ������������������,��������������
'"�����","�����","�����","�������� � �������","Գ�����"
' ����� ������������ �������
'1 �������� �����������
'2 msgbox

Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
'Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������)

'Set lsr = ls.ListRows(n.Row - 1)
Set r = Application.Intersect(lsc.Range, lsr.Range)
r = ��������
Set �������������������������������������� = r

    
 
End Function



Function ��������������������������(��������, ����������, Optional ���������, Optional AlwaysInserta As Boolean = True) As ListRow
'��������� ����� ������ � �������, �������������� ��������� ListObject
'���������-���������� ������������� ��������� ����� ������
' ���������, ������� �� ������ ���������� ������ � ������ ���� ��������� ������ �������,
'����� ����������� ����� ������, ���������� �� ����, ������ ���� ������� ����.
'���� True , ��� ���� ������� ������ ����� �������� �� ���� ������ ����.
'���� False , ���� ������ ���� ������� ����, ������� ����� �����������, ����� ������ ��� ������
'��� �������� ������ ��� ���; �� ���� ������ ���� ������� ���������� ������,
'��� ������ ����� �������� ����, ����� ����������� ����� ������.


Dim wb As Workbook
Dim ws As Worksheet
Dim lo As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set lo = ws.ListObjects(����������)
If IsMissing(���������) Then
Set lsr = lo.ListRows.Add
Else
Set lsr = lo.ListRows.Add(���������, AlwaysInsert)
End If
Set �������������������������� = lsr
End Function











Function ������������������(��������, ����������, ����������������, ��������������) As ListRow
'��������, ����������, ����������������, ������������������,��������������
'"�����","�����","�����","�������� � �������","Գ�����"
' ����� ������������ �������
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������������)
'Set lscf = ls.ListColumns(������������������)
Set n = lsc.Range.Find(��������������, LookAt:=xlWhole)
Set lsr = ls.ListRows(n.Row - 1)
'Set r = Application.Intersect(lscf.Range, lsr.Range)
Set ������������������ = lsr
End Function

Sub �������������������������(��������, ����������, ����������������, ��������������)
Dim lsr As ListRow
Set lsr = ������������������(��������, ����������, ����������������, ��������������)
'lsr.ClearContents
lsr.Delete
End Sub
Function ���������������������������(��������, ����������, ������) As ListRow
'��������, ����������, ����������������, ������������������,��������������
'"�����","�����","�����","�������� � �������","Գ�����"
' ����� ������������ �������
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)


Set lsr = ls.ListRows(������)

Set ��������������������������� = lsr
End Function




Function �����������������������������������������������(��������, ����������) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
On Error Resume Next
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)

Set r = ls.DataBodyRange
Set ����������������������������������������������� = r.SpecialCells(xlCellTypeVisible)


End Function
Function �����������������������������������(��������, ����������) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)

Set r = ls.Range
Set ����������������������������������� = r.SpecialCells(xlCellTypeVisible)


End Function
Function ������������������������������������������(��������, ����������, ����������) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
On Error Resume Next
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������)
Dim rr As Range
Set rr = ls.DataBodyRange
Set www = Intersect(lsc.Range, rr)
Set r = www.SpecialCells(xlCellTypeVisible)
'Set r = lsc.Range.SpecialCells(xlCellTypeVisible)
Set ������������������������������������������ = r


End Function
Sub �����������������������������(��������, ����������, ����������)
'
' ������29 ������
'

'
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������)
   ls.Sort.SortFields _
        .Clear
    ls.Sort.SortFields _
        .Add Key:=lsc.Range, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers
    With ls.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub







Sub �����������������(br)
Dim Findrange As Range
Dim f As Range
Set Findrange = ���������������������������������("��������", "��������", "��������")
Set f = FindAll(SearchRange:=Findrange.Cells, FindWhat:=br, LookAt:=xlPart)

If f Is Nothing Then

Set r = ������������������("��������", "��������", "��������", br)
End If
End Sub
Sub cvfd()
�������� = "������"
���������� = "������"
ist = Array("˳�������_ĳ�����", "��������", "�����", "����� ���")
br = Array("���������� 15 �� (4 ���) 0 ��.", "����4", "4", "6")
Set rr = ��������Masiv(��������, ����������, ist, br)

End Sub

Function ��������Masiv(��������, ����������, ist As Variant, br As Variant)
Dim Findrange As Range
Dim lsr As ListRow
Dim f As Range

Set Findrange = ���������������������������������(��������, ����������, ist(0))


Set f = FindAll(SearchRange:=Findrange.Cells, FindWhat:=br(0), LookAt:=xlPart)

If f Is Nothing Then

Set rr = ������������������(��������, ����������, ist(0), br(0))
For I = 0 To UBound(ist)
q = IndexColumn(��������, ����������, ist(I))
rr.Cells(q).value = br(I)

Next I

Else
Set lsr = ������������������(��������, ����������, ist(0), br(0))
Set rr = lsr
For I = 1 To UBound(ist)
q = IndexColumn(��������, ����������, ist(I))
rr.Range.Cells(q).value = br(I)

Next I
End If
'For i = 1 To UBound(ist)
'q = IndexColumn(��������, ����������, ist(i))
'rr.Range.Cells(q).Value = br(i)
'
'Next i
Set ��������Masiv = rr
End Function

 Function IndexColumn(��������, ����������, ����������)
 Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set zagrow = ls.HeaderRowRange
IndexColumn = Application.WorksheetFunction.Match(����������, zagrow, 0)


 End Function
 Sub DeleteColumn(��������, ����������, ����������)
 On Error Resume Next
 Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set zagrow = ls.HeaderRowRange
mIndexColumn = Application.WorksheetFunction.Match(����������, zagrow, 0)
ls.ListColumns(mIndexColumn).Delete

 End Sub
'---------------------------------------------------------------------------------------
' Procedure : �������������������������
' Author    : ������
' Date      : 26.11.2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
 Sub �������������������������(��������, ����������, ����������, Optional myCriteria As String)
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Dim z As Integer
   On Error GoTo �������������������������_Error

Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
z = IndexColumn(��������, ����������, ����������)
   If Not myCriteria = "" Then
ls.Range.AutoFilter Field:=z, Criteria1:= _
        myCriteria
        Else
    ls.Range.AutoFilter Field:=z
    End If

   On Error GoTo 0
   Exit Sub

�������������������������_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ������������������������� of Module ������������"
End Sub
Function ������������������(��������, ����������, ����������, ���) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim cc As Range
Dim rr As Range
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set zagrow = ls.HeaderRowRange
q = Application.WorksheetFunction.Match(����������, zagrow.Cells, 0)
Set lsr = ls.ListRows.Add(AlwaysInsert:=True)
'Set lsc = ls.ListColumns(����������)
Set rr = ls.Range.Rows(ls.Range.Rows.Count)
'Set cc = lsc.DataBodyRange
Set r = lsr.Range.Cells(q)
r.value = ���
Set ������������������ = rr


End Function
Function ��������������������(��������, ����������, ParamArray ��������������()) As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set lsc = ls.ListColumns(����������)
Set zagrow = ls.HeaderRowRange
q = Application.WorksheetFunction.Match(����������, zagrow.Cells, 0)
Set n = lsc.Range.Find(��������������, LookAt:=xlWhole)
Set lsr = ls.ListRows.Add(AlwaysInsert:=True)
'Set lsc = ls.ListColumns(����������)
Set rr = ls.Range.Rows(ls.Range.Rows.Count)
'Set cc = lsc.DataBodyRange
Set r = lsr.Range.Cells(q)
r.value = ���
Set �������������������� = rr


End Function


Sub CopyFilteredRowsAndHeadingsList1()
Dim wsL As Worksheet
Dim ws As Worksheet
Dim Rng As Range
Dim rng2 As Range
Dim Lst As ListObject

Application.ScreenUpdating = False
Set wsL = ActiveSheet
Set Lst = wsL.ListObjects(1)

With Lst.AutoFilter.Range
 On Error Resume Next
   Set rng2 = .Offset(1, 0).Resize(.Rows.Count - 1, 1) _
       .SpecialCells(xlCellTypeVisible)
 On Error GoTo 0
End With
If rng2 Is Nothing Then
   MsgBox "No data to copy"
Else
   Set ws = Sheets.Add
   Set Rng = Lst.AutoFilter.Range
   'copy rows with headings
   Rng.SpecialCells(xlCellTypeVisible).Copy _
     destination:=ws.Range("A1")
End If
   
Application.ScreenUpdating = True

End Sub
Function ���������������(��������, ����������)
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)

��������������� = ls.DataBodyRange.Rows.Count


End Function
Function ��������������������(�������� As Workbook, �������� As String, ����������) As ListObject
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range

If Not IsBookOpen(��������.name) Then Exit Function
If Not SheetExistBook(��������, ��������) Then Exit Function
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
Set �������������������� = ls
End Function
Function ������������()
Dim ls1 As ListObject
Dim ls2 As ListObject


Set ls1 = ��������������������(��������1, ��������1, ����������1)
Set ls2 = ��������������������(��������2, ��������2, ����������2)






End Function



Sub ������������(��������, ����������, ����������, Criteria, Optional ������ As Boolean)
'
' ������18 ������
Dim wb As Workbook
Dim ws As Worksheet
Dim ls As ListObject
Dim lsc As ListColumn
Dim lscf As ListColumn
Dim lsr As ListRow
Dim r As Range
Set ws = ThisWorkbook.Worksheets(��������)
Set ls = ws.ListObjects(����������)
If ������ Then
ws.ListObjects(����������).Range.AutoFilter Field:=IndexColumn(��������, ����������, ����������)
Else
    ws.ListObjects(����������).Range.AutoFilter Field:=IndexColumn(��������, ����������, ����������), Criteria1:= _
        Criteria
        End If
End Sub
 
Function MyTableRangeID(MyTable As ListObject, r As Range)
   Dim isect As Range
    
        Set isect = Application.Intersect(r, MyTable.Range)
        If Not (isect Is Nothing) Then
            MyTableRangeID = 1 ' "������ ����������� ������� " & MyTable.Name
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
               MyTableRangeID = 2 '������ � ���������"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.DataBodyRange)
            If Not (isect Is Nothing) Then
               MyTableRangeID = 3 '������ � ������� ������"
            End If
            Set isect = Nothing
            On Error Resume Next
            If MyTable.ShowTotals Then
                If Not Intersect(r, MyTable.TotalsRowRange) Is Nothing Then
                   MyTableRangeID = 5 '"������ � ������ ������" & vbNewLine
                Else
                    Temp = Temp & "������ ������ ����, �� ������ �� � ���" & vbNewLine
                End If
 
            Else
                If Not Intersect(r, MyTable.Range.Rows(lo.Range.Rows.Count)) Is Nothing Then
                   MyTableRangeID = 4 ' "������ ������ ���, ������ � ��������� ������ �������" & vbNewLine
                End If
            End If
            On Error GoTo 0
            Else
           MyTableRangeID = 0
        End If
   
End Function
Sub HeaderRowRangeStilNone(lo As ListObject)
'
' ������� ������� ��������� ����� ������� � ���� ������ �������
'

'
   
    With lo.HeaderRowRange
    .Interior.pattern = xlNone
    .Font.ColorIndex = xlAutomatic
      
    End With
End Sub
Function fgaAuthorList(��������, ����������, ����������)
count_tab = ���������������(��������, ����������)
 Set v = ������������������������������������������(��������, ����������, ����������)
 
lngLoopCount = 0
ReDim mgaAuthorList(count_tab - 1)
For I = 1 To count_tab
mgaAuthorList(lngLoopCount) = v(I)
lngLoopCount = lngLoopCount + 1

Next I
fgaAuthorList = mgaAuthorList
End Function
Function fgaAuthorListUnikum(��������, ����������, ����������, Optional Unikum)
count_tab = ���������������(��������, ����������)
 Set v = ������������������������������������������(��������, ����������, ����������)
 If Unikum Then Set v = mMacros.UnicumRange(v)
lngLoopCount = 0
ReDim mgaAuthorList(count_tab - 1)
For I = 1 To count_tab
mgaAuthorList(lngLoopCount) = v(I)
lngLoopCount = lngLoopCount + 1

Next I
fgaAuthorListUnikum = mgaAuthorList
End Function
Function fgaAuthorListVisible(��������, ����������, ����������)
 Set v = ������������������������������������������(��������, ����������, ����������)
 count_tab = v.Cells.Count

lngLoopCount = 0
ReDim mgaAuthorList(count_tab - 1)
For I = 1 To count_tab
mgaAuthorList(lngLoopCount) = v(I)
lngLoopCount = lngLoopCount + 1

Next I
fgaAuthorListVisible = mgaAuthorList
End Function
Function ListVisible(��������, ����������)
 Set v = �����������������������������������������������(��������, ����������)
 If v Is Nothing Then Exit Function
 
 count_tab = ListObjectColumnCount(��������, ����������)
 count_row = v.Cells.Count / count_tab

lngLoopCount = 1
ReDim mgaAuthorList(1 To count_row, 1 To count_tab)
For I = 1 To count_row
For j = 1 To count_tab
mgaAuthorList(I, j) = v(lngLoopCount)
lngLoopCount = lngLoopCount + 1
Next j
Next I
ListVisible = mgaAuthorList

End Function



'Sub ListControlListobjectColumn(forma As UserForm, ��������, ����������, ����������, �����������, ������������ As Boolean)
'Dim contr
'
'Set contr = forma.Controls(�����������)
'
'If ������������ Then
'v = fgaAuthorListUnikum(��������, ����������, ����������, ������������)
'Else
'v = fgaAuthorListVisible(��������, ����������, ����������)
'End If
'contr.List = v
'
'End Sub
Sub Test_Cell()
'��������� � �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
'� ����� ����� ������ ������
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
            Set lsr = ��������������������������(ActiveSheet.name, MyTable.name)
            
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ���������"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ������"
            End If
            On Error GoTo 0
        End If
    Next
End Sub
Sub Test_Cell_TypE(Optional v)
'��������� � �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
'� ����� ����� ������ ������ � � ������ ������ ���������
'�������� ������ � ������� ���������� �������� ������
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        Set isect = Application.Intersect(ActiveSheet.Rows(ActiveCell.Row), MyTable.Range)
            Set lsr = ��������������������������(ActiveSheet.name, MyTable.name)
       If IsMissing(v) Then v = Array(1)
       For I = 0 To UBound(v)
      lsr.Range.Cells(v(I)) = isect.Cells(v(I))
      Next I
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ���������"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ������"
            End If
            On Error GoTo 0
        End If
    Next
End Sub
Sub DEl_ROw_TypE()
'������� ��  �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
'
'������ � ������� ���������� �������� ������
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
      
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ���������"
                Exit Sub
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ������"
                Exit Sub
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveSheet.Rows(ActiveCell.Row), MyTable.Range)
            isect.Delete
            On Error GoTo 0
        End If
    Next
End Sub

Sub AddShowTotal()
'���������( ��� ������� ���� ��� ����)������ ������ � �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
       If MyTable.ShowTotals = True Then MyTable.ShowTotals = False Else MyTable.ShowTotals = True
            
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ���������"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ������"
                
                
            End If
            On Error GoTo 0
        End If
    Next
End Sub


Sub AddShowAutoFilterDropDown()
'���������( ��� ������� ���� ��� ����)������ �������� � �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
    Dim MyTable As ListObject, isect As Range
    Dim lsr As ListRow
    For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
        
       If MyTable.ShowAutoFilterDropDown = True Then MyTable.ShowAutoFilterDropDown = False Else MyTable.ShowAutoFilterDropDown = True
            
            Set isect = Application.Intersect(ActiveCell, MyTable.HeaderRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ���������"
            End If
            Set isect = Nothing
            On Error Resume Next
            Set isect = Application.Intersect(ActiveCell, MyTable.TotalsRowRange)
            If Not (isect Is Nothing) Then
                MsgBox "ActiveCell � ������"
                
                
            End If
            On Error GoTo 0
        End If
    Next
End Sub
Sub RemoveTableBodyData()
'������� ��  �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
'��� ��������
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then


'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
    MyTable.DataBodyRange.Delete
    
    
  End If
End If
    Next
End Sub
Sub Sort�olumnAscending()
' ���������� �� ����������� � ������ �������� ����� ������� �� ��������� �������� ������
Call Sort�olumn(1)
End Sub
Sub Sort�olumnDescending()
' ���������� �� �������� � ������ �������� ����� ������� �� ��������� �������� ������
Call Sort�olumn(2)
End Sub

Sub Sort�olumn(myOrder)
'
' ����������  � ������ �������� ����� ������� �� ��������� �������� ������
'1-xlAscending-�� ����������
'2-xlDescending-�� ��������
 Dim lo As ListObject
 Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim ActiveColumnListObject As Range
Dim ActiveColumnListObjectName As Range
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
'        �������� ������� �������� ����� �������
 Set ActiveColumnListObject = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
' ����� ��������� ������� �������� ����� �������
 Set ActiveColumnListObjectName = ActiveColumnListObject.Cells(1)
 KeyrangeStruktureAddress = MyTable.name & "[[#Headers],[#Data],[" & ActiveColumnListObjectName.value & "]]"
    MyTable.Sort. _
        SortFields.Clear
'  Range("��������[[#Headers],[#Data],[���]]")
    MyTable.Sort. _
        SortFields.Add Key:=Range(KeyrangeStruktureAddress), SortOn:= _
        xlSortOnValues, Order:=myOrder, DataOption:=xlSortTextAsNumbers
      v = MyTable.Sort.SortFields(1).Order
    With MyTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End If
    Next
End Sub



Sub SortFieldsClear()
'
' ���������� �� ����������� � ������ �������� ����� ������� �� ��������� �������� ������
'

 Dim lo As ListObject
 Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim ActiveColumnListObject As Range
Dim ActiveColumnListObjectName As Range
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
'        �������� ������� �������� ����� �������
 Set ActiveColumnListObject = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
' ����� ��������� ������� �������� ����� �������
 Set ActiveColumnListObjectName = ActiveColumnListObject.Cells(1)
 KeyrangeStruktureAddress = MyTable.name & "[[#Headers],[#Data],[" & ActiveColumnListObjectName.value & "]]"
    MyTable.Sort. _
        SortFields.Clear
''  Range("��������[[#Headers],[#Data],[���]]")
'    MyTable.Sort. _
'        SortFields.Add Key:=Range(KeyrangeStruktureAddress), SortOn:= _
'        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'    With MyTable.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With


'ActiveWorkbook.Worksheets("����������2").ListObjects("��������").Sort. _
'        SortFields.Add Key:=Range("��������[[#Headers],[#Data],[�����]]"), SortOn:= _
'        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortTextAsNumbers
'    With ActiveWorkbook.Worksheets("����������2").ListObjects("��������").Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With







    End If
    Next
End Sub







Sub ����������������������()
'������� ��  �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
'��� ��������
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then

Set isect = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
    Call �������������������������(ActiveSheet.name, MyTable.name, isect.Cells(1), ActiveCell.value)
 
  End If
End If
    Next
End Sub

Sub �����������������������������()
'������� ��  �������� ����� �������(���� �������� ������ ��������� � ����� ����� �������)
'��� ��������
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then

Set isect = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
    Call �������������������������(ActiveSheet.name, MyTable.name, isect.Cells(1))
 
  End If
End If
    Next
End Sub
Sub ��������������������������(Optional wb, Optional ShName, Optional loName, Optional r, Optional v)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(ShName) Then ShName = ActiveCell.value
If IsMissing(loName) Then loName = ActiveCell.value
If IsMissing(wb) Then Set wb = ActiveWorkbook
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then
If ActiveCell.value = "" Then
Call MsgBox("������� ����� ����������� � �������� ������", vbCritical, "�������� �����������")
Exit Sub

End If
Set isect = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
'Delete Table's Body Data
  If MyTable.ListRows.Count >= 1 Then
  
 Set sh = SheetExistBookCreate(wb, ShName, True)
 If IsMissing(r) Then Set r = sh.Cells(1, 1)
 
   Set New_tbl = myListObjectadd(wb, ShName, loName, r)
    If IsMissing(v) Then v = Array("�����", "����")
    For I = 0 To UBound(v)
    If I = 0 Then
 Call ListObjectColumnRENAME(New_tbl, 1, v(I))
 Else
 Call ListObjectColumnAdd(New_tbl, v(I))
 End If
 Next I
  End If
End If
    Next
End Sub
Sub ������������������(Optional wb, Optional ShName, Optional loName, Optional r, Optional v)

Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
If IsMissing(ShName) Then ShName = ActiveCell.value
If IsMissing(loName) Then loName = ActiveCell.value
If IsMissing(wb) Then Set wb = ActiveWorkbook

If loName = "" Then
Call MsgBox("������� ����� ����������� � �������� ������", vbCritical, "�������� �����������")
Exit Sub

End If

  
  
 Set sh = SheetExistBookCreate(wb, ShName, True)
 If IsMissing(r) Then Set r = sh.Cells(1, 1)
 
   Set New_tbl = myListObjectadd(wb, ShName, loName, r)
    If IsMissing(v) Then v = Array("�����", "����")
    For I = 0 To UBound(v)
    If I = 0 Then
 Call ListObjectColumnRENAME(New_tbl, 1, v(I))
 Else
 Call ListObjectColumnAdd(New_tbl, v(I))
 End If
 Next I
 

End Sub


Sub AutoFillFormulasInListsRevers()
'��� ���� ������ �� �������� listobject.
'���� ���� �������� �������, �������� "=890", ��� ���������� �� ����� �������. ���� ����� ������ ������, ���������� ���������.
'
'� ��� �������� �������������� � ��������, ���� � �������� ������� ����� vba? ����� �� ��������� ��� �������� �� ������ �������, ����� ����� ��������?


'�������������� ListObject
If Application.AutoCorrect.AutoFillFormulasInLists Then

Application.AutoCorrect.AutoFillFormulasInLists = False ' ��������� �������������� ListObject
Else
    Application.AutoCorrect.AutoFillFormulasInLists = True  ' �������� �������������� ListObject
    End If
End Sub


Sub ShowAllRecordsList()
'�������� �� ������

'��������� ������ ������ �� ������ � �������� ����� �������  �� ��������� ������, ���� ������ ��� ������������.
Dim Lst As ListObject
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then


 If MyTable.AutoFilter.FilterMode Then
    MyTable.AutoFilter.ShowAllData
  End If
End If
    Next


End Sub

Sub HideArrowsList1()
'hides all arrows except list 1 column 2
'��������� ��� ������ ���������� ������, ��� ������

'���� ����, �� ������, ��� ����������� ������ ����� ���� �� �������� � ������ 1. �������� ��������� Excel ���������� VBA ������� ������ ��� ��� ��������, ��� ����� ������� � ������ 1
Dim MyTable As ListObject, isect As Range
Dim tbl As ListObject
Dim New_tbl As ListObject
Dim r As Range
Application.ScreenUpdating = False
For Each MyTable In ActiveSheet.ListObjects
        Set isect = Application.Intersect(ActiveCell, MyTable.Range)
        If Not (isect Is Nothing) Then

Set ActiveColumnListObject = Application.Intersect(ActiveSheet.Columns(ActiveCell.Column), MyTable.Range)
I = 1

For Each C In MyTable.HeaderRowRange
 If I <> ActiveColumnListObject.Column Then
    MyTable.Range.AutoFilter Field:=I, _
      VisibleDropDown:=False
 Else
     MyTable.Range.AutoFilter Field:=I, _
      VisibleDropDown:=True
 End If
 I = I + 1
Next
End If
    Next

Application.ScreenUpdating = True



End Sub
Sub AddParametr(��������������, ��������)
�������� = "���������"
���������� = "Parameters"
���������������� = "Parameter"
������������������ = "Value"
Set r = ��������������������(��������, ����������, ����������������, ������������������, ��������������)
r.value = ��������
End Sub
 
Function Parametr(��������������)
�������� = "���������"
���������� = "Parameters"
���������������� = "Parameter"
������������������ = "Value"
Set r = ��������������������(��������, ����������, ����������������, ������������������, ��������������)
Set Parametr = r
End Function
 
Sub NullValueListObjects(��������, ����������, ����������, v, ��������������)
Dim r As Range
For I = 0 To UBound(v)
Set r = ��������������������(��������, ����������, ����������, v(I), ��������������)
If Not (r Is Nothing) Then r.value = ""
Next I
