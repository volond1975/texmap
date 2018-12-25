Attribute VB_Name = "mAdvanced_filter"
'расширенный фильтр (advanced filter)
'http://www.planetaexcel.ru/techniques/2/197/






Sub Расширенный_фильтр()
'Этот макрос должен быть в личной книге макросов, чтобы запускаться в любой книге Excel.
'И в центре управления безопасностью Excel должен быть открыт доступ к объектной модели проектов VBA.
''Перед применением макроса нужно выделить заголовки таблицы (все или часть), которые будут скопированы в шапку расширенного фильтра.
'Проверял работоспособность в Excel 2007 и 2010.
'Ну и повесть макрос на кнопку или горячую клавишу - это совсем просто, если нужно.


'Проверяем защищен ли проект VBA
If ActiveWorkbook.VBProject.Protection = 1 Then MsgBox "Проект защищен. Создание расширенного фильтра невозможно": Exit Sub

'Количество строк расширенного фильтра
kol = 6

'Проверяем установлин ли уже расширенный фильтр. Если да, то удаляем.
If Rows(2).Interior.Color <> 49407 Then

    If Selection.Rows.Count > 1 Then MsgBox "Нужно выделить не более одной строки", vbCritical: Exit Sub
    'Определяем границы выделенной области
    row1 = Selection.Row
    col1 = Selection.Column
    col2 = col1 + Selection.Cells.Count - 1
    'Вставляем строчки для расширенного фильтра
    For i = 1 To kol
        Range("A1").EntireRow.Insert
        If i > 1 Then Rows(1).Interior.Color = 49407
    Next i
    'Копируем шапку
    Range(Cells(row1 + kol, col1), Cells(row1 + kol, col2)).Select
    Selection.copy
    Cells(1, col1).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    'Формируем текст макроса
    MacroText = "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(13)
    MacroText = MacroText & "If Not Intersect(Target, Range(cells(2," & col1 & "),cells(" & kol - 1 & "," & col2 & "))) Is Nothing Then" & Chr(13)
    MacroText = MacroText & "On Error Resume Next" & Chr(13)
    MacroText = MacroText & "ActiveSheet.ShowAllData" & Chr(13)
    MacroText = MacroText & "cells(" & row1 + kol & "," & col1 & ").CurrentRegion.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=cells(1," & col1 & ").CurrentRegion" & Chr(13)
    MacroText = MacroText & "End If" & Chr(13)
    MacroText = MacroText & "End Sub"

    'Добавляем макрос
    ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule.AddFromString (MacroText)
    
Else:
    
    'Удаляем строки расширенного фильтра
    Range("A1:A" & kol).EntireRow.Delete
    Application.CutCopyMode = False
    
    'Удаляем макрос
    With ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule
             If .Find("Worksheet_Change", 1, 1, .CountOfLines, 1) = True Then
                iStartLine = .ProcStartLine("Worksheet_Change", 0)
                iCountLines = .ProcCountLines("Worksheet_Change", 0)
                .DeleteLines iStartLine, iCountLines
             End If
    End With
    
End If

End Sub
