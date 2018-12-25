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
wh.PageSetup.PrintArea = r.address
End Sub

Sub DeleteSheet(wb As Workbook, SheetName As String)
'������� ���� � ������ SheetName, ����  ����
     On Error Resume Next
     Application.DisplayAlerts = False
     wb.Sheets(SheetName).Delete
     Application.DisplayAlerts = True
End Sub
Public Sub lst_clear_cell(Lst As Worksheet)
Lst.Cells.Clear
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
Function RangeColumName(wb As Workbook, WSName As String, ColumnName As String)
Dim r As Range
Dim sh As Worksheet
Set sh = wb.Worksheets(WSName)
lr = LastRow("�������������")
Set b = ����������������(wb, WSName, ColumnName)
Set r = sh.Range(sh.Cells(2, b.Column), sh.Cells(lr, b.Column))
Set RangeColumName = r
End Function
Sub ColumnNameMoveNewListName(Lst As Worksheet, NEW_lst As Worksheet, lstname As String, newlstname As String)
'������� ������� � ������ ����� � ������ �� ����� � ����� ������
    Lst.Select
    Set b = ����������������(Lst.Parent, Lst.name, lstname)
    Columns(b.Column).Select
    Application.CutCopyMode = False
    Selection.Copy
    NEW_lst.Select
    Set b = ����������������(NEW_lst.Parent, NEW_lst.name, newlstname)
    Columns(b.Column).Select
    ActiveSheet.Paste
End Sub
Sub ColumnNameMoveNewList(Lst As Worksheet, NEW_lst As Worksheet, lstcolumn, newlstcolumn)
'������� ������� � ������ ����� � ������ �� ������ �������
    Lst.Select
    Columns(lstcolumn).Select
    Application.CutCopyMode = False
    Selection.Copy
    NEW_lst.Select
    Columns(newlstcolumn).Select
    ActiveSheet.Paste
End Sub
  Function UnicumRange(r As Range) As Variant

 '���������� ���������� ������ �� �������

 Dim v() As Variant
     s = 1
     For I = 1 To r.Cells.Count
         ReDim Preserve v(s - 1)
        
         For j = 0 To s - 1
             If r.Cells(I) = v(j) Then
             GoTo 111
             End If
         Next j
 v(s - 1) = r.Cells(I).value
    s = s + 1
111:
     Next I
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
   sError = "������ " & Err.Number & " (" & Err.Description & ") � ��������� Delete_Empty_Rows ������ Module MyMacros" & IIf(Erl <> 0, " � ������ " & Erl, "")
    frmERROR.Show
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

Sub �������_����������(sh As Worksheet, r As Range, ����� As String, ���������� As String)
'
' ������5 ������
' ������ ������� 03.10.2012 (������������)
'
On Error Resume Next

Call �����������������
   sh.Hyperlinks.Add Anchor:=r, address:= _
        �����, TextToDisplay:=����������
End Sub

Function ����_����_����������(r As Range)
'
' ������5 ������
' ������ ������� 03.10.2012 (������������)
'

Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
A = r.Hyperlinks(r.Hyperlinks.Count).address
   If fso.FileExists(A) Then ����_����_���������� = True Else ����_����_���������� = False
End Function

Sub �����������������()
On Error Resume Next

If r.Hyperlinks.Count > 0 Then r.Hyperlinks.Delete
End Sub
Sub CalculationM()
'
' ������1 ������
' ������ ������� 05.06.2012 (������������)
'

'
    With Application
        .Calculation = xlManual
        .MaxChange = 0.001
    End With
    ActiveWorkbook.PrecisionAsDisplayed = False
End Sub
Sub CalculationA()
'
' ������3 ������
' ������ ������� 05.06.2012 (������������)
'

'
    With Application
        .Calculation = xlAutomatic
        .MaxChange = 0.001
    End With
    ActiveWorkbook.PrecisionAsDisplayed = False
End Sub
