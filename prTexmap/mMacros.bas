Attribute VB_Name = "mMacros"

 Sub Intro()
 If lCountWorkbooks = 0 Then Exit Sub
     With Application
     .ScreenUpdating = False
     .EnableEvents = False
     lCalculation = .Calculation
     .Calculation = xlCalculationManual
     End With
End Sub

 Sub OptimiseAppProperties()
     If lCountWorkbooks = 0 Then Exit Sub
     With Application
     .ScreenUpdating = False
     .EnableEvents = False
     lCalculation = .Calculation
     .Calculation = xlCalculationManual
     End With
 End Sub
 Sub Outro()
 If lCountWorkbooks = 0 Then Exit Sub
     With Application
     .StatusBar = False
    .ScreenUpdating = True
    .DisplayAlerts = True
     .EnableEvents = True
     .Calculation = IIf(lCalculation = 0, xlAutomatic, lCalculation)
     End With
     
End Sub
 Sub ResetAppProperties()
     If lCountWorkbooks = 0 Then Exit Sub
     With Application
     .StatusBar = False
    .ScreenUpdating = True
    .DisplayAlerts = True
     .EnableEvents = True
     .Calculation = IIf(lCalculation = 0, xlAutomatic, lCalculation)
     End With















Sub ����������������(r As Range, col As Integer)
'
' ������1 ������
' ������ ������� 12.05.2011 (��������)
'

'
    r.Font.ColorIndex = col
End Sub
Sub FontBold12(r As Range)
'
' ������1 ������
' ������ ������� 12.05.2011 (��������)
'

'
 FontBold12 = r.Font.Bold
End Sub
Sub ����������������(NameSheet As String, r As Range, p As Boolean)
'
' ������25 ������
' ������ ������� 14.05.2011 (��������)
'

Dim sh As Worksheet
Set sh = Worksheets(NameSheet)
sh.Activate
r.Select
ActiveWindow.FreezePanes = p
    
End Sub
Sub IV()
ActiveCell.value = InversiaValue(ActiveCell)
End Sub
Sub ������������(wh As Worksheet, r As Range)
'������ ������ ������
wh.PageSetup.PrintArea = r.Address
End Sub

Sub DeleteSheet(wb As Workbook, SheetName As String)
'������� ���� � ������ SheetName, ����  ����
     On Error Resume Next
     Application.DisplayAlerts = False
     wb.Sheets(SheetName).Delete
     Application.DisplayAlerts = True
End Sub
Public Sub lst_clear_cell(lst As Worksheet)
lst.Cells.Clear
'lst.Cells.Font.ColorIndex = xlNone
'lst.Cells.Interior.ColorIndex = xlNone

End Sub




Sub ReferenceStyle_Change()
'
' �������� ����� ����������� ���������� ��������
' ������ ������� 09.06.2011 (��������)
'

'
    With Application
    If .ReferenceStyle = xlR1C1 Then
    .ReferenceStyle = xlA1
    Else
    .ReferenceStyle = xlR1C1
        End If
    End With
End Sub
Function txt_ReferenceStyle()
'
' �������� ����� ����������� ���������� ��������
' ������ ������� 09.06.2011 (��������)
'

'
    With Application
    If .ReferenceStyle = xlR1C1 Then
    txt_ReferenceStyle = "A1"
    Else
    txt_ReferenceStyle = "R1C1"
        End If
    End With
End Function
  Function CheckName(sName As String, lSheet As Long)
'���� � sName  ������������ �������,
'����������� � ������������� � ����� �����.
'���� ����� ������� ������������ - ��� ����� ������ �������,
'������ �� ���������
  
  
        Dim objRegExp As Object
     Set objRegExp = CreateObject("VBScript.RegExp")
      objRegExp.Global = True: objRegExp.IgnoreCase = True
      If lSheet = 1 Then
     objRegExp.pattern = "[:,\\,/,?,\*,\],\[]"
      Else
     objRegExp.pattern = "[:,\\,/,?,\*,\<,\>,\|,""""]"
      End If
     CheckName = objRegExp.Replace(sName, "")
  End Function
  Public Function GoodName(ByVal forBook As Workbook, ByVal testName As String) As Boolean
    Dim pSheet As Object, RegExp As Object
    GoodName = False
    If (Len(testName) <= 31) And (Len(testName) > 0) Then
        Set RegExp = CreateObject("VBScript.RegExp")
        RegExp.pattern = "[\\/\*\[\]\?:]"
        If Not RegExp.Test(testName) Then
            GoodName = True
            For Each pSheet In forBook.Sheets
                If VBA.StrComp(testName, pSheet.name, vbTextCompare) = 0 Then
                    GoodName = False: Exit For
                End If
            Next pSheet
        End If
    End If
End Function
Function RangeColumName(wb As Workbook, WSName As String, ColumnName As String)
Dim r As Range
Dim sh As Worksheet
Set sh = wb.Worksheets(WSName)
lr = LastRow("�������������")
Set B = ����������������(wb, WSName, ColumnName)
Set r = sh.Range(sh.Cells(2, B.Column), sh.Cells(lr, B.Column))
Set RangeColumName = r
End Function
Sub ColumnNameMoveNewListName(lst As Worksheet, NEW_lst As Worksheet, lstname As String, newlstname As String)
'������� ������� � ������ ����� � ������ �� ����� � ����� ������
    lst.Select
    Set B = ����������������(lst.parent, lst.name, lstname)
    Columns(B.Column).Select
    Application.CutCopyMode = False
    Selection.copy
    NEW_lst.Select
    Set B = ����������������(NEW_lst.parent, NEW_lst.name, newlstname)
    Columns(B.Column).Select
    ActiveSheet.Paste
End Sub
Sub ColumnNameMoveNewList(lst As Worksheet, NEW_lst As Worksheet, lstcolumn, newlstcolumn)
'������� ������� � ������ ����� � ������ �� ������ �������
    lst.Select
    Columns(lstcolumn).Select
    Application.CutCopyMode = False
    Selection.copy
    NEW_lst.Select
    Columns(newlstcolumn).Select
    ActiveSheet.Paste
End Sub
  Function UnicumRange(r As Range) As Variant

 '���������� ���������� ������ �� �������

 Dim v() As Variant
     s = 1
     For i = 1 To r.Cells.Count
         ReDim Preserve v(s - 1)
        
         For j = 0 To s - 1
             If r.Cells(i) = v(j) Then
             GoTo 111
             End If
         Next j
 v(s - 1) = r.Cells(i).value
    s = s + 1
111:
     Next i
   UnicumRange = v
 End Function
Sub UpdateProgress(Pct)
    With Dialog
      .FrameProgress.Caption = Format(Pct, "0%")
      .LabelProgress.Width = Pct * (.FrameProgress.Width - 10)
      .Repaint
    End With
End Sub
 Private Sub Delete_Empty_Rows_In_Table()
       Dim lLastRow As Long, li As Long
    If lCountWorkbooks = 0 Then Exit Sub
     On Error GoTo Delete_Empty_Rows_Error
     If MsgBox("��� ������ ������ � ������� ��������� �����" & vbCrLf & Space(15) & "����� �������. ����������?", vbYesNo, "��������������") = vbNo Then Exit Sub
     lLastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count
    Call OptimiseAppProperties
     For li = lLastRow To 1 Step -1
    If Rows(li).Text = "" Then Rows(li).Delete
    Next li
    Call ResetAppProperties
    Exit Sub
Delete_Empty_Rows_Error:
   sError = "������ " & Err.Number & " (" & Err.description & ") � ��������� Delete_Empty_Rows ������ Module MyMacros" & IIf(Erl <> 0, " � ������ " & Erl, "")
    frmERROR.Show
 End Sub

Sub DeleteEmptyRows(sh As Worksheet)
'    LastRow = Sh.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count    '���������� ������� �������
    Application.ScreenUpdating = False
    For r = LastRow(sh.name) To 1 Step -1           '�������� �� ��������� ������ �� ������
        If Application.CountA(Rows(r)) = 0 Then Rows(r).Delete   '���� � ������ ����� - ������� ��
    Next r
End Sub

'������ ��������� ������ ��������� ������
Sub FormulaViewOn()
    ActiveWindow.NewWindow
    ActiveWorkbook.Windows.Arrange ArrangeStyle:=xlHorizontal
    ActiveWindow.DisplayFormulas = True
End Sub

'������ ���������� ������ ��������� ������
Sub FormulaViewOff()
    If ActiveWindow.WindowNumber = 2 Then
        ActiveWindow.Close
        ActiveWindow.WindowState = xlMaximized
        ActiveWindow.DisplayFormulas = False
    End If
End Sub

'Function ��������������������(��������, ����������, ����������������, ������������������, ��������������) As Range
''��������, ����������, ����������������, ������������������,��������������
''"�����","�����","�����","�������� � �������","Գ�����"
'' ����� ������������ �������
'Dim wb As Workbook
'Dim ws As Worksheet
'Dim ls As ListObject
'Dim lsc As ListColumn
'Dim lscf As ListColumn
'Dim lsr As ListRow
'Dim r As Range
'Set ws = ThisWorkbook.Worksheets(��������)
'Set ls = ws.ListObjects(����������)
'Set lsc = ls.ListColumns(����������������)
'Set lscf = ls.ListColumns(������������������)
'Set n = lsc.Range.Find(��������������)
''Set lsr = ws(n.row)
'If Not n Is Nothing Then
'Set lsr = ls.ListRows(n.Row - 1)
'Set r = Application.Intersect(lscf.Range, lsr.Range)
'Set �������������������� = r
'Else
''Set �������������������� = "�������� �� �������-" & ��������������
'End If
'End Function

Sub �����������()
'
' ������2 ������
'

'
 
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Sub ������������()
'
' ������3 ������
'

'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Sub �������������������_GetAnotherWorkbook()
    Dim wb As Workbook
    Set wb = GetAnotherWorkbook
    If Not wb Is Nothing Then
        MsgBox "������� �����: " & wb.FullName, vbInformation
    Else
        MsgBox "����� �� �������", vbCritical: Exit Sub
    End If
    ' ��������� ������ �� ��������� �����
   X = wb.Worksheets(1).Range("a2")
    ' ...
End Sub

Function GetAnotherWorkbook() As Workbook
    ' ���� � ������ ������ ������� 2 �����, ������� ��������� ������ �������� �����
   ' ���� ������ �������, ������� ����� ����� ����� - ����� ������������ �����
   On Error Resume Next
    Dim coll As New Collection, wb As Workbook
    For Each wb In Workbooks
        If wb.name <> ThisWorkbook.name Then
            If Windows(wb.name).Visible Then coll.Add CStr(wb.name)
        End If
    Next wb
    Select Case coll.Count
        Case 0    ' ��� ������ �������� ����
           MsgBox "��� ������ �������� ����", vbCritical, "Function GetAnotherWorkbook"
        Case 1    ' ������� ��� ������ ���� ����� - � � ����������
           Set GetAnotherWorkbook = Workbooks(coll(1))
        Case Else    ' ������� ��������� ���� - ������������� �����
           For i = 1 To coll.Count
                txt = txt & i & vbTab & coll(i) & vbNewLine
            Next i
            msg = "�������� ���� �� �������� ����, � ������� � ���������� �����:" & _
                  vbNewLine & vbNewLine & txt
            res = InputBox(msg, "������� ����� ���� ����", 1)
            If IsNumeric(res) Then Set GetAnotherWorkbook = Workbooks(coll(Val(res)))
    End Select
End Function
