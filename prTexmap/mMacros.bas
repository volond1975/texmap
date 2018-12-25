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















Sub ЦветТекстаЯчейки(r As Range, col As Integer)
'
' Макрос1 Макрос
' Макрос записан 12.05.2011 (Владелец)
'

'
    r.Font.ColorIndex = col
End Sub
Sub FontBold12(r As Range)
'
' Макрос1 Макрос
' Макрос записан 12.05.2011 (Владелец)
'

'
 FontBold12 = r.Font.Bold
End Sub
Sub ЗакрепитьОбласть(NameSheet As String, r As Range, p As Boolean)
'
' Макрос25 Макрос
' Макрос записан 14.05.2011 (Владелец)
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
Sub РегионПечати(wh As Worksheet, r As Range)
'Задать регион печати
wh.PageSetup.PrintArea = r.Address
End Sub

Sub DeleteSheet(wb As Workbook, SheetName As String)
'удаляем лист с именем SheetName, если  есть
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
' Изменяет стиль отображения заголовков столбцов
' Макрос записан 09.06.2011 (Владелец)
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
' Изменяет стиль отображения заголовков столбцов
' Макрос записан 09.06.2011 (Владелец)
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
'если в sName  присутствуют символы,
'запрещенные к использованию в имени книги.
'Если такие символы присутствуют - они будут просто удалены,
'ошибки не возникнет
  
  
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
lr = LastRow("ФорматСводных")
Set B = ЗаголовокСтолбца(wb, WSName, ColumnName)
Set r = sh.Range(sh.Cells(2, B.Column), sh.Cells(lr, B.Column))
Set RangeColumName = r
End Function
Sub ColumnNameMoveNewListName(lst As Worksheet, NEW_lst As Worksheet, lstname As String, newlstname As String)
'Перенос столбца с одного листа в другой по имени в конец нового
    lst.Select
    Set B = ЗаголовокСтолбца(lst.parent, lst.name, lstname)
    Columns(B.Column).Select
    Application.CutCopyMode = False
    Selection.copy
    NEW_lst.Select
    Set B = ЗаголовокСтолбца(NEW_lst.parent, NEW_lst.name, newlstname)
    Columns(B.Column).Select
    ActiveSheet.Paste
End Sub
Sub ColumnNameMoveNewList(lst As Worksheet, NEW_lst As Worksheet, lstcolumn, newlstcolumn)
'Перенос столбца с одного листа в другой по номеру столбца
    lst.Select
    Columns(lstcolumn).Select
    Application.CutCopyMode = False
    Selection.copy
    NEW_lst.Select
    Columns(newlstcolumn).Select
    ActiveSheet.Paste
End Sub
  Function UnicumRange(r As Range) As Variant

 'Возвращает уникальные данные из региона

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
     If MsgBox("Все пустые строки в таблице активного листа" & vbCrLf & Space(15) & "будут удалены. Продолжить?", vbYesNo, "Предупреждение") = vbNo Then Exit Sub
     lLastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count
    Call OptimiseAppProperties
     For li = lLastRow To 1 Step -1
    If Rows(li).Text = "" Then Rows(li).Delete
    Next li
    Call ResetAppProperties
    Exit Sub
Delete_Empty_Rows_Error:
   sError = "Ошибка " & Err.Number & " (" & Err.description & ") в процедуре Delete_Empty_Rows модуля Module MyMacros" & IIf(Erl <> 0, " в строке " & Erl, "")
    frmERROR.Show
 End Sub

Sub DeleteEmptyRows(sh As Worksheet)
'    LastRow = Sh.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count    'определяем размеры таблицы
    Application.ScreenUpdating = False
    For r = LastRow(sh.name) To 1 Step -1           'проходим от последней строки до первой
        If Application.CountA(Rows(r)) = 0 Then Rows(r).Delete   'если в строке пусто - удаляем ее
    Next r
End Sub

'макрос включения режима просмотра формул
Sub FormulaViewOn()
    ActiveWindow.NewWindow
    ActiveWorkbook.Windows.Arrange ArrangeStyle:=xlHorizontal
    ActiveWindow.DisplayFormulas = True
End Sub

'макрос выключения режима просмотра формул
Sub FormulaViewOff()
    If ActiveWindow.WindowNumber = 2 Then
        ActiveWindow.Close
        ActiveWindow.WindowState = xlMaximized
        ActiveWindow.DisplayFormulas = False
    End If
End Sub

'Function ЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска) As Range
''ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения,ЗначениеПоиска
''"Листы","Листы","Листы","Столбцов в таблице","Фінплан"
'' Можно использовать индексы
'Dim wb As Workbook
'Dim ws As Worksheet
'Dim ls As ListObject
'Dim lsc As ListColumn
'Dim lscf As ListColumn
'Dim lsr As ListRow
'Dim r As Range
'Set ws = ThisWorkbook.Worksheets(ИмяЛиста)
'Set ls = ws.ListObjects(ИмяТаблицы)
'Set lsc = ls.ListColumns(ИмяСтолбцаПоиска)
'Set lscf = ls.ListColumns(ИмяСтолбцаЗначения)
'Set n = lsc.Range.Find(ЗначениеПоиска)
''Set lsr = ws(n.row)
'If Not n Is Nothing Then
'Set lsr = ls.ListRows(n.Row - 1)
'Set r = Application.Intersect(lscf.Range, lsr.Range)
'Set ЗначениеУмнойТаблицы = r
'Else
''Set ЗначениеУмнойТаблицы = "Значение не найдено-" & ЗначениеПоиска
'End If
'End Function

Sub СЕТКАТОНКАЯ()
'
' Макрос2 Макрос
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
Sub СЕТКАТОЛСТАЯ()
'
' Макрос3 Макрос
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
Sub ПримерИспользования_GetAnotherWorkbook()
    Dim wb As Workbook
    Set wb = GetAnotherWorkbook
    If Not wb Is Nothing Then
        MsgBox "Выбрана книга: " & wb.FullName, vbInformation
    Else
        MsgBox "Книга не выбрана", vbCritical: Exit Sub
    End If
    ' обработка данных из выбранной книги
   X = wb.Worksheets(1).Range("a2")
    ' ...
End Sub

Function GetAnotherWorkbook() As Workbook
    ' если в данный момент открыто 2 книги, функция возвратит вторую открытую книгу
   ' если помимо текущей, открыто более одной книги - будет предоставлен выбор
   On Error Resume Next
    Dim coll As New Collection, wb As Workbook
    For Each wb In Workbooks
        If wb.name <> ThisWorkbook.name Then
            If Windows(wb.name).Visible Then coll.Add CStr(wb.name)
        End If
    Next wb
    Select Case coll.Count
        Case 0    ' нет других открытых книг
           MsgBox "Нет других открытых книг", vbCritical, "Function GetAnotherWorkbook"
        Case 1    ' открыта ещё только одна книга - её и возвращаем
           Set GetAnotherWorkbook = Workbooks(coll(1))
        Case Else    ' открыто несколько книг - предоставляем выбор
           For i = 1 To coll.Count
                txt = txt & i & vbTab & coll(i) & vbNewLine
            Next i
            msg = "Выберите одну из открытых книг, и введите её порядковый номер:" & _
                  vbNewLine & vbNewLine & txt
            res = InputBox(msg, "Открыто более двух книг", 1)
            If IsNumeric(res) Then Set GetAnotherWorkbook = Workbooks(coll(Val(res)))
    End Select
End Function
