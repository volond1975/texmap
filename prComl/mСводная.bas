Attribute VB_Name = "m�������"

Sub ������������������()
Attribute ������������������.VB_Description = "������ ������� 12.05.2011 (��������)"
Attribute ������������������.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ABS_WB As Workbook
Dim Ost_lst As Worksheet
Dim SV_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Dim pi As PivotItem
Set ABS_WB = ThisWorkbook


Call New_Multi_Table_Pivot
Call CellAutoFilterVisible("3")
Set pc = ABS_WB.PivotCaches(1)
Set Ost_lst = ABS_WB.Worksheets("�������")
Set SV_lst = ABS_WB.Worksheets("�������")
Set pt = SV_lst.PivotTables("CV_Ob")

        Dim r As Range
        Set r = Worksheets("����������").Range("B1")
    Set EndPeriodM = r.End(xlDown)
    Set EndPeriodY = EndPeriodM.Offset(columnoffset:=1)
    With pt.PivotFields("���")
    For Each pi In .PivotItems
      
       pi.Visible = True

 Next
    For Each pi In .PivotItems
       If pi.value = EndPeriodY.value Then
       pi.Visible = True
Else
pi.Visible = False
End If

        
        Next
    End With
    
    With pt.PivotFields("�����")
     For Each pi In .PivotItems
     
       pi.Visible = True

Next
       For Each pi In .PivotItems
       If pi.value = Val(EndPeriodM.value) Then
       pi.Visible = True
Else
pi.Visible = False
End If
Next
    End With
' �all lst_clear_cell(Ost_lst)
 Ost_lst.Cells.Clear
''Ost_lst.Cells.Font.ColorIndex = xlNone
'Ost_lst.Cells.Interior.ColorIndex = xlNone
 
 
 SV_lst.Cells.Copy
'    Sheets("�������").Select
'    Cells.Select
'    Selection.Copy
     
   Ost_lst.Select
   Cells.Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'    Rows("1:4").Select
'    Range("A4").Activate
    Application.CutCopyMode = False
'    Selection.Delete Shift:=xlUp
    Rows("1:4").Delete Shift:=xlUp
End Sub
Sub PivotFieldOrientation(PF As PivotField)
Attribute PivotFieldOrientation.VB_Description = "������ ������� 12.05.2011 (��������)"
Attribute PivotFieldOrientation.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������3 ������
' ������ ������� 12.05.2011 (��������)
'
Select Case PF.Orientation
    Case xlHidden
        MsgBox "Hidden field"
    Case xlRowField
        MsgBox "Row field"
    Case xlColumnField
        MsgBox "Column field"
    Case xlPageField
        MsgBox "Page field"
    Case xlDataField
        MsgBox "Data field"
End Select

  
End Sub
Sub ColumnsAutoFit(r As Range)
Attribute ColumnsAutoFit.VB_Description = "������ ������� 12.05.2011 (��������)"
Attribute ColumnsAutoFit.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����������������
'

    r.Columns.AutoFit
   
   
End Sub
Sub RowsAutoFit(r As Range)
'
' ����������������
'

    r.Rows.AutoFit
 
   
End Sub
Sub ���������������������()
Dim r As Range
Set r = Selection
If r.Columns.Count < r.Rows.Count Then
Call ColumnsAutoFit(r)
Else
Call RowsAutoFit(Selection)
End If
End Sub
Sub �����������������������������������(pt As PivotTable, ���������������� As String, ����������������� As String)
On Error Resume Next
Dim pi As PivotItem

Call Intro
pt.ManualUpdate = True
k = 0

    With pt.PivotFields(����������������)
    For Each pi In .PivotItems
    Select Case �����������������
    Case ""
    pi.Visible = True
    Case Else
       If pi.value = ����������������� Then
       pi.Visible = True
       k = k + 1
Else
pi.Visible = False
'If Err = 1004 Then MsgBox Err
End If
End Select
Next
Call Outro
pt.ManualUpdate = False
End With
End Sub

Sub ���������������������������������������(���������������� As String, ����������������� As String)
Dim ABS_WB As Workbook
Dim O_lst As Worksheet
Dim SV_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Dim pi As PivotItem

Set ABS_WB = ThisWorkbook
Call Intro
���������������� = "������������"
����������������� = "���������� 2% 25�� ������ �-�+"
Set SV_lst = ABS_WB.Worksheets("�������")
Set pc = ABS_WB.PivotCaches(1)
Set pt = SV_lst.PivotTables("CV_Ob")
Cells(1, 1).Select
ActiveWindow.SmallScroll Up:=65000

     pt.PivotSelection = ���������������� & "[" & ����������������� & "]"
    
    ActiveWindow.SmallScroll Down:=ActiveCell.Row - 6
Call Outro

End Sub


Sub PVTableRefresh()
Dim ABS_WB As Workbook
Dim O_lst As Worksheet
Dim SV_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Set ABS_WB = ThisWorkbook

Set O_lst = ABS_WB.Worksheets("����� ����")
Set SV_lst = ABS_WB.Worksheets("�������")
lr = LastRow(O_lst.name)
col = 12
sd = Chr(39) & "" & O_lst.name & "'!R1C1:R" & lr & "C" & col
Set pc = ABS_WB.PivotCaches(1)
'pc.SourceData = "''" & O_lst.Name & "'!" & O_lst.Range("A1").CurrentRegion.Address
pc.SourceData = sd
Debug.Print pc.SourceData
Debug.Print sd
Set pt = SV_lst.PivotTables("CV_Ob")


End Sub
Sub �����������������������������������������(pt As PivotTable, zn As Boolean)
'
' ������������������������������ ������
' ������ ������� 16.05.2011 (��������)
'

'
    With pt
        .ColumnGrand = zn
        
    End With
End Sub
Sub ����������������������������������������(pt As PivotTable, zn As Boolean)
'
' ������������������������������ ������
' ������ ������� 16.05.2011 (��������)
'

'
    With pt
                .RowGrand = zn
    End With
End Sub
Sub New_Multi_Table_Pivot()
ResultSheetName = f�������.ComboBox1.value
ResultPivotTableName = f�������.ComboBox1.value & "_ALL"
������ = "������1�"
Call New_Multi_Table_Pivot1(ResultSheetName, ResultPivotTableName, ������)
End Sub


Sub New_Multi_Table_Pivot1(ResultSheetName, ResultPivotTableName, ������)
    Dim I As Long
    Dim arSQL() As String
    Dim objPivotCache As PivotCache
    Dim objRS As Object
'    Dim ResultSheetName As String
    Dim SheetsNames As String
    Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim NEW_lst As Worksheet
Dim �K_lst As Worksheet
Dim �K_Cell As Range
Dim Find_Cell As Range

Set ABS_WB = ThisWorkbook
Set �K_Cell = ActiveCell
'Set O_lst = ABS_WB.Worksheets("����� ����")

  
    '��� �����, ���� ����� ���������� �������������� �������
'    ResultSheetName = "��_������������"
'    ResultPivotTableName = "��_������������"
'    '������ ���� ������ � ��������� ���������
'For Each NEW_lst In ABS_WB.Worksheets
'If NEW_lst.Name Like "*_*" Then
'SheetsNames = SheetsNames & "SELECT * FROM [" & NEW_lst.Name & "$] UNION ALL "
'
'End If
'    Next
'  SheetsNames = Trim(SheetsNames)
'  SheetsNames = Left(SheetsNames, Len(SheetsNames) - 10)
'������ = "�������"
  SheetsNames = ArraySheetName(ABS_WB, ������)
    '��������� ��� �� �������� � ������ �� SheetsNames
    With ABS_WB
       Set objRS = CreateObject("ADODB.Recordset")
objRS.Open SheetsNames, _
Join$(Array("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=", _
.FullName, ";Extended Properties=""Excel 8.0;"""), vbNullString)
End With
 '������� ������ ���� ��� ������ �������������� ������� �������
On Error Resume Next
Application.DisplayAlerts = False
Worksheets(ResultSheetName).Delete
Set wsPivot = SheetExistBookCreate(ThisWorkbook, ResultSheetName, False)

 
  
   
    
  
    '������� �� ���� ���� ������� �� ��������������� ����
    Set objPivotCache = ActiveWorkbook.PivotCaches.Add(xlExternal)
    Set objPivotCache.Recordset = objRS
    Set objRS = Nothing
    With wsPivot
        objPivotCache.CreatePivotTable TableDestination:=wsPivot.Range("A3")
       
Set pt = .PivotTables(1)
pt.name = ResultPivotTableName
        Set objPivotCache = Nothing
        Range("A3").Select
    End With
    k = 1
    For Each PF In pt.PivotFields
'PF.Name = O_lst.Cells(1, k)
'k = k + 1

Debug.Print PF.name
Next
'  Call CellAutoFilterVisible(1)
    
    
Application.DisplayAlerts = True
End Sub

Sub PVTableAddABC()
Dim ABS_WB As Workbook
Dim O_lst As Worksheet
Dim SV_lst As Worksheet
Dim ABC_lst As Worksheet
Dim pt As PivotTable
Dim pc As PivotCache
Dim PF As PivotField
Set ABS_WB = ThisWorkbook

Set O_lst = ABS_WB.Worksheets("����� ����")
Set SV_lst = ABS_WB.Worksheets("�������")
Set ABC_lst = ABS_WB.Worksheets("���")

'ABC_lst.Cells.Clear
Set pc = ABS_WB.PivotCaches(1)
Set pt = SV_lst.PivotTables("CV_Ob")
k = 1
For Each PF In pt.PivotFields
 PF.name = O_lst.Cells(1, k)
 k = k + 1
' pf.Position
Next
'lr = LastRow(O_lst.Name)
'col = 12
'sd = Chr(39) & "" & O_lst.Name & "'!R1C1:R" & lr & "C" & col
'Set pc = ABS_WB.PivotCaches(1)
''pc.SourceData = "''" & O_lst.Name & "'!" & O_lst.Range("A1").CurrentRegion.Address
'pc.SourceData = sd
'Debug.Print pc.SourceData
'Debug.Print sd
'Set pt = SV_lst.PivotTables("CV_Ob")
'pt.DataFields (1)

End Sub

Sub CellAutoFilterVisible(k)
Dim sh As Worksheet
Dim pt As PivotTable
Dim PF As PivotField
Dim PF� As PivotField
Set sh = Worksheets("�������������")
Dim sh_cv As Worksheet
'On Error Resume Next
z = 0
With sh
.Rows(1).AutoFilter Field:=1, Criteria1:=k
     If .AutoFilterMode = True And .FilterMode = True Then
        With .AutoFilter.Range.Columns(1)
             Set iFilterRange = _
             .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlVisible)
             cm = 2
             For Each iCell In iFilterRange
             
             Set sh_cv = SheetExistBookCreate(ThisWorkbook, iCell.Offset(columnoffset:=1).value, False)
             Set pt = sh_cv.PivotTables(iCell.Offset(columnoffset:=2).value)
              pt.DisplayErrorString = True
             pt.ErrorString = ""
             pt.ColumnGrand = iCell.Offset(columnoffset:=2).Font.Bold
             
           If iCell.Offset(columnoffset:=3).Font.Underline = 2 Then
pt.RowGrand = True
Else
pt.RowGrand = False
      End If
             If r = 0 Then
             For Each PF In pt.PivotFields
             If Not PF.Orientation = 0 Then PF.Orientation = xlHidden
             Next
             r = 1
             End If
             
             If iCell.Offset(columnoffset:=10).value <> "" Then
             '�������� � ������� �������
             pt.PivotFields(iCell.Offset(columnoffset:=10).value).Orientation = xlPageField

 If CommentExist(iCell.Offset(columnoffset:=10)) Then
        '������������ ��������� ������������ � ���������� ���� ����
        comentar = CommentTEXT(iCell.Offset(columnoffset:=10))
         pt.PivotFields(iCell.Offset(columnoffset:=10).value).CurrentPage = _
         "" & comentar & ""
         End If
        End If
    
             If iCell.Offset(columnoffset:=3).value <> "" Then
            '�������� � ������� �����
             pt.PivotFields(iCell.Offset(columnoffset:=3).value).Orientation = xlRowField
        If Not iCell.Offset(columnoffset:=3).Font.Bold Then
        pt.PivotFields(iCell.Offset(columnoffset:=3).value).Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
        End If
             End If
             If iCell.Offset(columnoffset:=4).value <> "" Then
              '�������� � ������� ��������
             pt.PivotFields(iCell.Offset(columnoffset:=4).value).Orientation = xlColumnField
             If Not iCell.Offset(columnoffset:=4).Font.Bold Then
        pt.PivotFields(iCell.Offset(columnoffset:=4).value).Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
        End If
             
             End If
             
           If iCell.Offset(columnoffset:=16).value <> "" Then
         Formula = "=" & iCell.Offset(columnoffset:=15).value
'      Formula = "=����/'��''��'"
 Set pfc = pt.CalculatedFields.Add(iCell.Offset(columnoffset:=14).value _
        , Formula, True)
    pt.PivotFields(iCell.Offset(columnoffset:=14).value). _
        Orientation = ����������������(iCell.Offset(columnoffset:=16).value)
         
         
         
         
         
         
           
             End If
             
             
             
             
             
             If iCell.Offset(columnoffset:=5).value <> "" Then
             If z = 0 Then
             For I = 1 To pt.DataFields.Count
             pt.DataFields(1).Orientation = xlHidden
            
             Next I
             z = 1
             End If
        
             
             
             
             
             
              
             If iCell.Offset(columnoffset:=5).value Like "*����� ��*.*" Then
             l = InStr(1, iCell.Offset(columnoffset:=5).value, ".")
             Dim t As String
             t = iCell.Offset(columnoffset:=5).Text
            Mid(t, l, 1) = "#"
            Set PF = pt.PivotFields(t)
             Else
             Set PF = pt.PivotFields(iCell.Offset(columnoffset:=5).value)
             End If
            
              
             
             
             
             Call AddDataFildConsolidac(pt, PF, iCell.Offset(columnoffset:=6).value, iCell.Offset(columnoffset:=7).value)
             
           


             
             
             
             
             End If
              If iCell.Offset(columnoffset:=13).value <> "" Then
             pt.Format (iCell.Offset(columnoffset:=13).value)
             End If
                Debug.Print iCell.Row & "_" & iCell.Offset(columnoffset:=13).value
             Next
        End With
'        .ShowAllData '���������� �� - �������������
     End If
End With
End Sub
Sub AddDataFildConsolidac(pt As PivotTable, PF As PivotField, Zagolovok As String, CF As String)
'�������������� ��������
'����� xlSum
'���������� xlCount
'������� xlAverage
'�������� xlMax
'������� xlMin
'���������� �����    xlCountNums
'������������ xlProduct
'��������� ����������    xlStDev
'����������� ����������  xlStDevP
'����������� xlUnknown
'��������� ��������� xlVar
'����������� ���������   xlVarP

             
             
             Select Case CF
                Case "�����"
              pt.AddDataField PF, Zagolovok, xlSum
                Case "����������"
              pt.AddDataField PF, Zagolovok, xlCount
                Case "�������"
              pt.AddDataField PF, Zagolovok, xlAverage
                Case Else

            End Select
End Sub
  Function ArraySheetName(ABS_WB As Workbook, �������)
  Dim NEW_lst As Worksheet
  Dim SheetsNames As String
  '������ ���� ������ � ��������� ���������
For Each NEW_lst In ABS_WB.Worksheets
If NEW_lst.name Like ������� Then
SheetsNames = SheetsNames & "SELECT * FROM [" & NEW_lst.name & "$] UNION ALL "
 
End If
    Next
  SheetsNames = Trim(SheetsNames)
  SheetsNames = Left(SheetsNames, Len(SheetsNames) - 10)
  ArraySheetName = SheetsNames
  End Function
  
Sub �����������()
' ===========================================
' ������ ��� ���������� ������ ����� �������
' ===========================================
'

If ActiveWorkbook.ShowPivotTableFieldList Then
    ActiveWorkbook.ShowPivotTableFieldList = False
    Else
    ActiveWorkbook.ShowPivotTableFieldList = True
    End If
End Sub
Sub ��������������������(r As Range, col As Long)
'
' ������11 ������
' ������ ������� 25.05.2011 (��������)
Dim b As Range

Set b = LastColumn(ActiveSheet.name, 1)
Col_r = r.Column
NameCol = Cells(1, Col_r)
NamPol = "����� �� " & NameCol



lr = LastRow(ActiveSheet.name)
    Cells(1, col).FormulaR1C1 = NamPol
    Set z = ����������������(ThisWorkbook, ActiveSheet.name, "����")
gABC = z.Column

formul = "=RC[-" & col - gABC & "]*RC[" & Col_r - col & "]"

    Range(Cells(2, col), Cells(lr, col)).FormulaR1C1 = formul
    
End Sub
Sub ����������������()
Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim NEW_lst As Worksheet
Dim �K_lst As Worksheet
Dim �K_Cell As Range
Dim Find_Cell As Range
Set ABS_WB = ThisWorkbook
Set �K_Cell = ActiveCell

For Each NEW_lst In ABS_WB.Worksheets

If NEW_lst.name Like "*_*" Then
NEW_lst.Activate

Call ����������
End If
    Next
End Sub
Sub ������������������()
Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim NEW_lst As Worksheet
Dim �K_lst As Worksheet
Dim �K_Cell As Range
Dim Find_Cell As Range
Set ABS_WB = ThisWorkbook
Set �K_Cell = ActiveCell

For Each NEW_lst In ABS_WB.Worksheets

If NEW_lst.name Like "*_*" Then
NEW_lst.Activate

Call ������������
End If
    Next
End Sub
Sub ����������()
Dim b As Range
Dim Y As String
'v = Array("���.���.�-��", "����.� -��", "����.�-��", "���.�-��")
v = Array(4, 5, 6, 7)
lr = LastRow(ActiveSheet.name)
For I = 0 To UBound(v)
'y = v(i)
'Set z = ����������������(ThisWorkbook, ActiveSheet.Name, y)
'gABC = z.Column
gABC = v(I)
Set r = Range(Cells(2, gABC), Cells(lr, gABC))
q = r
For W = 1 To UBound(q)
If q(W, 1) = 0 Then q(W, 1) = ""
Next W
r.value = q
Next I




    
End Sub

Sub ������������()
Dim b As Range
Dim Y As String
'v = Array("���.���.�-��", "����.� -��", "����.�-��", "���.�-��")
v = Array(4, 5, 6, 7)
lr = LastRow(ActiveSheet.name)
For I = 0 To UBound(v)
'y = v(i)
'Set z = ����������������(ThisWorkbook, ActiveSheet.Name, y)
'gABC = z.Column
gABC = v(I)
Set r = Range(Cells(2, gABC), Cells(lr, gABC))
q = r
For W = 1 To UBound(q)
If q(W, 1) = "" Then q(W, 1) = 0
Next W
r.value = q
Next I




    
End Sub



Sub ��������_���()
SheetName = "���"
RowsDelete = 4
Call ��������_����(SheetName, RowsDelete)
End Sub
Sub �������RC�_���()
SheetName = "���"
RowsDelete = 4
Call ��������_����(SheetName, RowsDelete)
End Sub

Sub ��������_����(CVSheetName, SheetName, RowsDelete)
'
' ������11 ������
' ������ ������� 24.05.2011 (��������)
'
Dim shcopy As Worksheet
Dim shpaste As Worksheet
Dim pt As PivotTable
Dim rf As Object
Dim v() As String
Dim rr As Range
Call Intro
Set shpaste = SheetExistBookCreate(ThisWorkbook, SheetName, True)
'Set shpaste = Worksheets(SheetName)
Set shcopy = Worksheets(CVSheetName)
    shpaste.Select
    Cells.Select
    Selection.ClearContents
    shcopy.Select
    Cells.Select
    Selection.Copy
    Sheets(SheetName).Select
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Application.CutCopyMode = False
    
    Rows("1:" & RowsDelete).Delete Shift:=xlUp
'    Columns("D:D").Select
'    With Selection.Font
'        .Name = "Georgia"
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'    End With
   Selection.Interior.ColorIndex = xlNone
   Selection.Font.ColorIndex = 0
   Rows(RowsDelete - 2 & ":" & RowsDelete - 2).AutoFilter
   
  Set pt = shcopy.PivotTables(1)
' RowsDelete = pt.PivotFields
 ActiveWindow.FreezePanes = False
 Cells(2, pt.RowFields.Count + 1).Select
' ActiveWindow.FreezePanes = True
 Set rr = Cells(1, (pt.RowFields.Count))
 
 Range(rr, rr.End(xlToLeft)).Interior.ColorIndex = 9
 Range(rr, rr.End(xlToLeft)).Font.ColorIndex = 2
 Range(rr, rr.End(xlToLeft)).Font.Bold = True
 Set rr = Cells(1, (pt.RowFields.Count + 1))
 
 Range(rr, rr.End(xlToRight)).Interior.ColorIndex = 37
 Range(rr, rr.End(xlToRight)).Font.Bold = True
 ReDim v(pt.RowFields.Count - 1)
 For g = 0 To pt.RowFields.Count - 1
 v(g) = Cells(2, g + 1).value
 Next g
 lr = LastRow(shpaste.name) ' - RowsDelete
' Set b = LastColumn(sh.Name, 1)
 Set z = shpaste.Range(shpaste.Cells(2, 1), shpaste.Cells(lr, pt.RowFields.Count))
 X = z
 For I = 1 To UBound(X, 1)
 For j = 1 To UBound(X, 2)
 If X(I, j) = "" Then X(I, j) = v(j - 1)
 If X(I, j) <> v(j - 1) Then v(j - 1) = X(I, j)
 Next j
 Next I
z.value = X
 Call Outro
End Sub
Sub �����������(r As Range, col As Long)
'
' ����������� ������������ ������ � ������� ���
' ������ ������� 25.05.2011 (��������)

Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim Nel_lst As Worksheet
Dim O_lst As Worksheet
Dim ABS_lst As Worksheet
Dim nr As Range
Dim f As Range
Dim b As Range
Set ABS_WB = ThisWorkbook
Set ABS_lst = ABS_WB.ActiveSheet
Set b = LastColumn(ABS_lst.name, 1)
Set zags = ABS_lst.Range(ABS_lst.Cells(1, 1), b)
Set zag = ����������������(ABS_WB, ABS_lst.name, "����������")
If zag Is Nothing Then
b.Offset(columnoffset:=1).value = "����������"
Set b = b.Offset(columnoffset:=1)
End If
lr = LastRow(ABS_lst.name)






    Cells(1, col).FormulaR1C1 = NamPol
    Set z = ����������������(ThisWorkbook, ActiveSheet.name, "����������")
gABC = z.Column

formul = "=RC[-" & col - gABC & "]*RC[" & Col_r - col & "]"

    Range(Cells(2, col), Cells(lr, col)).FormulaR1C1 = formul
    
End Sub
Sub ���������������������������������()
'
' ������������� ������
' ������ ������� 11.07.2011 (��������)
'

Dim r As Range
Dim r_col As Range
Set r = ActiveCell
Set r_col = Cells(1, r.Column)
Dim psname As String
psname = "'" & r_col.value & "'"
psname = psname & "['" & r.value & "']"
    Worksheets("�������").PivotTables("CV_Ob").PivotSelection = psname
End Sub
Sub ���������������������������������All()
'
' ������������� ������
' ������ ������� 11.07.2011 (��������)
'

Dim r As Range
Dim r_col As Range
Set r = ActiveCell
Set r_col = Cells(1, r.Column)
Dim psname As String
psname = "'" & r_col.value & "'"
psname = psname & "[" & "ALL" & "]"
    Worksheets("�������").PivotTables("CV_Ob").PivotSelection = psname
End Sub


Sub ��������������������������������������������()
Dim wb As Workbook
Dim shNamenkl As Worksheet
Dim r As Range
Dim r_col As Range
Dim psname As String
Dim NameZag As String
Dim ColName As String
Set wb = ActiveWorkbook

Set shNamenkl = wb.Worksheets("������������")
Set r = ActiveCell
Set r_col = Cells(1, r.Column)
ColName = r_col.value
psname = PivotSelectionName(ColName, r.value)
NameZag = "�������"
NamePoz = r.value
poz_row = r.Row
Set poz_col = ����������������(wb, shNamenkl.name, NameZag)

Poz = shNamenkl.Cells(poz_row, poz_col.Column)


Call ��������������������������������������(psname, ColName, Poz, NamePoz)
End Sub

Function PivotSelectionName(ColName As String, it As String) As String
Dim psname As String
psname = "'" & ColName & "'"
'psname = ColName

psname = psname & "['" & it & "']"
PivotSelectionName = psname

End Function

Sub ��������������������������������������(psname As String, ColName As String, Poz As Variant, NamePoz)
Dim wb As Workbook
Dim shSvodnay As Worksheet
Dim pt As PivotTable
Dim PF As PivotField
Dim pi As PivotItem
On Error Resume Next
Set wb = ActiveWorkbook

Set shSvodnay = wb.Worksheets("�������")
Set pt = shSvodnay.PivotTables("CV_Ob")
Set PF = pt.PivotFields(ColName)
Set pi = PF.PivotItems(NamePoz)

Call ����������������������������(PF)
pi.Visible = True
 pt.PivotSelection = psname
 If Err = 1004 Then
 Else
 pi.Visible = True
 pt.PivotSelect "'" & NamePoz & "'", xlLabelOnly, True
 pi.Position = Poz
 
 End If

End Sub
Function ArrayFiltrVisible()



Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim Nel_lst As Worksheet

 
Dim O_lst As Worksheet
Dim ABS_lst As Worksheet
Dim nr As Range
Dim f As Range
Dim v() As Range
Set ABS_WB = ActiveWorkbook
Set ABS_lst = ABS_WB.ActiveSheet
'Set B = LastColumn(ABS_lst.Name, 1)
'Set zags = ABS_lst.Range(ABS_lst.Cells(1, 1), B)
'Set zag = ����������������(ABS_WB, ABS_lst.Name, "�������")
'If zag Is Nothing Then
'B.Offset(columnoffset:=1).Value = "�������"
'Set B = B.Offset(columnoffset:=1)
'Else
'Set B = zag
'End If
Set b = ���������������������������������(ABS_WB, ABS_lst.name, "�������")
Set m = ����������������(ABS_WB, ABS_lst.name, "������������")
NABC = m.Column




If ABS_lst.AutoFilterMode = True Then
   If ABS_lst.FilterMode = True Then
      iCountOfRows = ABS_lst.AutoFilter.Range.Columns(NABC).SpecialCells(xlVisible).Count
     
    Set f = ABS_lst.AutoFilter.Range.Columns(NABC).SpecialCells(xlVisible)
    ReDim Preserve v(f.Cells.Count - 2)
    k = 0
    For Each nr In f.Cells
     
If nr.Row <> 1 Then
Set v(k) = nr
k = k + 1
'���� �������� � ������� ������� ������ �� ��������������� ���������� ����� � �������
If Cells(nr.Row, b.Column).value = "" Then Cells(nr.Row, b.Column).value = k
  
    End If
    Next
'      End If
   End If
End If
ArrayFiltrVisible = v
End Function



Sub �������������������������������������()
Dim wb As Workbook
Dim shNamenkl As Worksheet
Dim r As Range
Dim r_col As Range
Dim psname As String
Dim NameZag As String
Dim ColName As String
Dim Counter As Long             '�������� ��� ������������
Dim TotalCells As Long          '�����������
Dim pi As PivotItem
Dim pis As PivotItems
Dim PF As PivotField
Set wb = ActiveWorkbook
Set shNamenkl = wb.Worksheets("������������")
shNamenkl.Activate
v = ArrayFiltrVisible()
Call Intro
 TotalCells = UBound(v)
'             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             '���� ���������� ����� �������, �� �������� ����� � ���������� ����������� ��������� ����������
             If TotalCells >= 10 Then

                 Application.ScreenUpdating = True
                 frmProgress.lPodpis.Caption = "����������� ���������� �������..."
                 frmProgress.LabelProgress.Width = 0
                 frmProgress.Show vbModeless
                 Application.ScreenUpdating = False
             End If
             Counter = 0
'             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'     pos = (1 / UBound(v)) * 100     '���� ����� ������� � ����� ���������� �������





For I = 0 To UBound(v)


Set r_col = Cells(1, v(I).Column)
ColName = r_col.value
psname = PivotSelectionName(ColName, v(I).value)
NameZag = "�������"
NamePoz = v(I).value
poz_row = v(I).Row
Set poz_col = ����������������(wb, shNamenkl.name, NameZag)

Poz = shNamenkl.Cells(poz_row, poz_col.Column)
Call ��������������������������������������(psname, ColName, Poz, NamePoz)


Counter = Counter + 1
                         If TotalCells >= 10 Then

                            If Counter Mod 100 <> 0 Then

                                 With frmProgress
                                     .FrameProgress.Caption = Format(Counter / TotalCells, "0%")
                                     .LabelProgress.Width = (Counter / TotalCells) * (.FrameProgress.Width - 10)
                                     .Repaint
                                 End With

                            End If

                         End If


Next I

frmProgress.Hide

Set shSvodnay = wb.Worksheets("�������")
Set PF = shSvodnay.PivotTables("CV_Ob").PivotFields("������������")
k = 1
Set pis = PF.PivotItems
 TotalCells = pis.Count - UBound(v)
  Counter = 0
'             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             '���� ���������� ����� �������, �� �������� ����� � ���������� ����������� ��������� ����������
             If TotalCells >= 10 Then

                 Application.ScreenUpdating = True
                 frmProgress.lPodpis.Caption = "����������� ���������� �������..."
                 frmProgress.LabelProgress.Width = 0
                 frmProgress.Show vbModeless
                 Application.ScreenUpdating = False
             End If
             Counter = 0
For I = 1 + UBound(v) To pis.Count
Set pi = PF.PivotItems(I)
 If pi.Visible = True Then pi.Visible = False
 Counter = Counter + 1
                         If TotalCells >= 10 Then

                            If Counter Mod 100 <> 0 Then

                                 With frmProgress
                                     .FrameProgress.Caption = Format(Counter / TotalCells, "0%")
                                     .LabelProgress.Width = (Counter / TotalCells) * (.FrameProgress.Width - 10)
                                     .Repaint
                                 End With

                            End If

                         End If
       
  Next
Call Outro
End Sub
Sub jjjjj()
MsgBox ����������������("���������������")
End Sub
Function ����������������(ip)
en = Array(xlDataField, xlColumnField, xlPageField, xlRowField)

ru = Array("�������������", "���������������", "��������������", "������������")
For I = 0 To 3
If ip = ru(I) Then ���������������� = en(I)
Next I
End Function
Sub �����������������������������(znach, ������)
Dim wb As Workbook
Dim WSName As String
Dim ColumnName As String
Dim v As Variant
Dim r As Range
Set wb = ThisWorkbook
WSName = "�������������"
ColumnName = "����"
Set b = ����������������(wb, WSName, ColumnName)
Set KodColumn = ����������������(wb, WSName, "���")
Set DelRowColumn = ����������������(wb, WSName, "������� ������")
Set PerenosSheetColumn = ����������������(wb, WSName, "���������� � ����")
Worksheets("�������������").Activate

If Worksheets("�������������").AutoFilterMode = True And Worksheets("�������������").FilterMode = True Then
Worksheets("�������������").Rows(1).AutoFilter
Worksheets("�������������").Rows(1).AutoFilter
End If
Set r = Worksheets("�������������").Columns(b.Column)
Set r1 = r.Find(What:=znach, LookAt:=xlWhole, after:=Cells(1, b.Column))
Set sm = Worksheets("�������������").Cells(r1.Row, KodColumn.Column)
Set DelRow = Worksheets("�������������").Cells(r1.Row, DelRowColumn.Column)
Set PerenosSheet = Worksheets("�������������").Cells(r1.Row, PerenosSheetColumn.Column)
ResultSheetName = znach
ResultPivotTableName = znach & "_ALL"

Call New_Multi_Table_Pivot1(ResultSheetName, ResultPivotTableName, ������)
Call CellAutoFilterVisible(sm.value)


If f�������.cbPerenos Then
SheetName = PerenosSheet.value
RowsDelete = DelRow.value
Call ��������_����(Worksheets("�������������").Cells(r1.Row, 2).value, SheetName, RowsDelete)
'If f�������.cbupr Then
'SheetName = PerenosSheet.Value
'RowsDelete = DelRow.Value
'Call ��������_����(Worksheets("�������������").Cells(r1.row, 2).Value, SheetName, RowsDelete)
'End If
End If



End Sub

Sub ��������������������������(znach, ������)
Dim wb As Workbook
Dim WSName As String
Dim ColumnName As String
Dim v As Variant
Dim r As Range
Set wb = ThisWorkbook
WSName = "�������������"
ColumnName = "����"
Set b = ����������������(wb, WSName, ColumnName)
Set KodColumn = ����������������(wb, WSName, "���")
Set DelRowColumn = ����������������(wb, WSName, "������� ������")
Set PerenosSheetColumn = ����������������(wb, WSName, "���������� � ����")
Worksheets("�������������").Activate

If Worksheets("�������������").AutoFilterMode = True And Worksheets("�������������").FilterMode = True Then
Worksheets("�������������").Rows(1).AutoFilter
Worksheets("�������������").Rows(1).AutoFilter
End If
Set r = Worksheets("�������������").Columns(b.Column)
Set r1 = r.Find(What:=znach, LookAt:=xlWhole, after:=Cells(1, b.Column))
Set sm = Worksheets("�������������").Cells(r1.Row, KodColumn.Column)
Set DelRow = Worksheets("�������������").Cells(r1.Row, DelRowColumn.Column)
Set PerenosSheet = Worksheets("�������������").Cells(r1.Row, PerenosSheetColumn.Column)
ResultSheetName = znach
ResultPivotTableName = znach & "_ALL"




SheetName = PerenosSheet.value
RowsDelete = DelRow.value
Call ��������_����(Worksheets("�������������").Cells(r1.Row, 2).value, SheetName, RowsDelete)




End Sub
