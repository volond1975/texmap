Attribute VB_Name = "CreatingNewDocs"
Public Const LINK_HEADER_TABLE$ = "<ExcelTable>"
Public FileNameDoc As String
Type myHuperlink
myAddress As String
myLink As String
myCapton As String
End Type
Const wdPrintAllDocument = 0
Const wdPrintCurrentPage = 2
Const wdPrintSelection = 1
Const wdPrintRangeOfPages = 4
Const wdPrintFromTo = 3
Const wdActiveEndPageNumber = 3



Function СформироватьДоговоры()
Dim wb As Workbook
Dim sh As Worksheet
Dim WSName As String
Dim b As Range
Dim ColumnName As String
Dim ИмяФайлаШаблона ' "шаблон.dot"
Dim КоличествоОбрабатываемыхСтолбцов ' 21
Dim РасширениеСоздаваемыхФайлов ' ".doc"
Dim ИмяСтолбцаКлючи As String
Dim ИмяСтолбцаПроверки As String
Dim ИмяСтолбцаЗначений As String
Dim ЗаголовокСтолбцаКлючи As Range
Dim ЗаголовокСтолбцаПроверки As Range
Dim ЗаголовокСтолбцаЗначения As Range
Dim ИмяТаблици As String
On Error Resume Next
 'Range("cDog")

РасширениеСоздаваемыхФайлов = ".doc"
ИмяТаблици = "Форма"
ИмяСтолбцаКлючи = "Заголовки"
ИмяСтолбцаПроверки = "Заголовки"
ИмяСтолбцаЗначений = "Значения"


Set wb = ThisWorkbook
Set sh = wb.Worksheets(ИмяТаблици)
ИмяШаблона = sh.Range("fDot").value
ИмяФайлаШаблона = ИмяШаблона & ".dot"
Set ЗаголовокСтолбцаКлючи = FindAll(sh.Rows(1), ИмяСтолбцаКлючи)
Set ЗаголовокСтолбцаПроверки = FindAll(sh.Rows(1), ИмяСтолбцаПроверки)
Set ЗаголовокСтолбцаЗначения = FindAll(sh.Rows(1), ИмяСтолбцаЗначений)

    ПутьШаблона = Replace(wb.FullName, wb.name, ИмяФайлаШаблона)
    НоваяПапка = NewFolderName & Application.PathSeparator
    Dim Row As Range, pi As New ProgressIndicator
    n = 2
r = LastRow(ИмяТаблици)
    rc = r - n - 2
'    r = Cells(Rows.Count, "A").End(xlUp).row: Rc = r - 2
'    If Rc < 1 Then MsgBox "Строк для обработки не найдено", vbCritical: Exit Sub

    pi.Show "Формирование договоров": pi.ShowPercents = True: s1 = 10: s2 = 90: p = s1: A = (s2 - s1) / rc
    pi.StartNewAction , s1, "Запуск приложения Microsoft Word"

    ' Dim WA As Word.Application, WD As Word.Document: Set WA = New Word.Application    ' c подключением библиотеки Word
    Dim WA As Object, WD As Object: Set WA = CreateObject("Word.Application")    ' без подключения библиотеки Word
WA.Visible = True
WA.TrackRevisions = False



'    For Each row In ActiveSheet.Rows(N & ":" & r)
'        With row
'         ipn = Trim$(.Cells(3))
'         ipn = Trim$(.Cells(8))
'WSName = "Акты"
'ColumnName = "NДодат"
'Set b = ЗаголовокСтолбца(WB, WSName, ColumnName)
'With Worksheets("Акты")
'.Activate
'If .AutoFilterMode = True And .FilterMode = True Then
'.Rows(1).AutoFilter
'.Rows(1).AutoFilter
'End If
'Set R = .Columns(b.Column)
'
'v = CellAutoFilterVisible(ipn)
'
'  End With


' sh.Activate
''  v = Split(Range("f№Акт").Value, "\")
'КВВИД = sh.Range("fПолнаяДіл").Value
'КВВИД = Replace(КВВИД, ".", "")
'            ИмяФайла = КВВИД & "-" & sh.Range("fDot").Value & "-" & sh.Range("f№ЛК").Value ' Trim$(MonthTK(Range("cMonth")) & "_" & СокрTK(Range("cLicVa")) & "_" & Range("сN"))
            ИмяФайла = NewFaleName
'            FileNameDoc = НоваяПапка & ИмяФайла & РасширениеСоздаваемыхФайлов
FileNameDoc = NewFileFullName(НоваяПапка, ИмяФайла, РасширениеСоздаваемыхФайлов)
'            pi.StartNewAction p, p + a / 3, "Создание нового файла на основании шаблона", ФИО
            Set WD = WA.Documents.Add(ПутьШаблона): DoEvents
WD.Revisions.AcceptAll
WD.DeleteAllComments
            pi.StartNewAction p + A / 3, p + A * 2 / 3, "Замена данных ...", ИмяФайла
            For I = n To r
            If sh.Cells(I, 1).value <> "" Then
          If HasLinkToObject(sh.Cells(I, 2).Text) Then
          
      Call InsertObjectIntoDOC(WD, sh.Cells(I, 2).Text, "{" & sh.Cells(I, 1) & "}")
          Else
                FindText = "{" & sh.Cells(I, 1) & "}": ReplaceText = Trim$(sh.Cells(I, 2))

                ' так почему-то заменяет не всё (не затрагивает таблицу)
                'WA.Selection.Find.Execute FindText, , , , , , , wdFindContinue, False, ReplaceText, True

                pi.line3 = "Заменяется поле " & FindText
                With WD.Range.Find
                    .Text = FindText
                    .Replacement.Text = ReplaceText
                    .Forward = True
                    .Wrap = 1
                    .Format = False: .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=2
                  
                    
                    
                End With
            
    WD.ActiveWindow.View.ShowRevisionsAndComments = False
WD.ActiveWindow.View.Revisions.AcceptAll
WD.ActiveWindow.View.DeleteAllComments
              
               
  
              
                 
                    
                    
          End If
                
                DoEvents
  End If
 
            Next I
           
            pi.StartNewAction p + A * 2 / 3, p + A, "Сохранение файла ...", ИмяФайла, " "
            WD.SaveAs FileNameDoc: WD.Close False: DoEvents
'            p = p + a
'        End With
'
'
'
'    Next row

    pi.StartNewAction s2, , "Завершение работы приложения Microsoft Word", " ", " "
    WA.Quit False: pi.Hide
    Msg = "Сформировано " & rc + 2 - n - 1 & " договоров. Все они находятся в папке" & vbNewLine & НоваяПапка
    СформироватьДоговоры = НоваяПапка
'    MsgBox msg, vbInformation, "Готово"
End Function










Function NewFolderName() As String
Dim twb As Workbook
Dim fso As Scripting.FileSystemObject
Dim WShConst As Worksheet
Set twb = ThisWorkbook
With twb

Set fso = New Scripting.FileSystemObject
Set WShConst = .Worksheets("Форма")
v = Split(WShConst.Range("f№ДП").value, "\")
Дата = v(0)
 'WShConst.Range("cNДокумента").Value
'ВидРубки = WShConst.Range("fКвартал").Value
'Лісництво = WShConst.Range("cLicVa").Value
sl = v(1) '
КВВИД = WShConst.Range("fПолнаяДіл").value
КВВИД = Replace(КВВИД, ".", "")
p = "Техкарты"
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If
p = "Техкарты\" & Дата
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If
p = "Техкарты\" & Дата & "\" & sl
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If
 p = "Техкарты\" & Дата & "\" & sl & "\" & КВВИД
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If

    
    End With
    
'  FilenName = ThisWorkbook.Path & "\" & "Техкарты" & "\" & Range(z & "NTK") & ".doc"
    
End Function
Function NewFaleName()
Dim shp As Worksheet
Dim lo As ListObject
Dim q As Range
Set wb = ThisWorkbook
Set sh = wb.Worksheets("Форма")
ИмяШаблона = sh.Range("fDot").value
ИмяФайлаШаблона = ИмяШаблона & ".dot"
Set ЗаголовокСтолбцаКлючи = FindAll(sh.Rows(1), ИмяСтолбцаКлючи)
Set ЗаголовокСтолбцаПроверки = FindAll(sh.Rows(1), ИмяСтолбцаПроверки)
Set ЗаголовокСтолбцаЗначения = FindAll(sh.Rows(1), ИмяСтолбцаЗначений)

  
    НоваяПапка = NewFolderName & Application.PathSeparator

 'sh.Activate
'  v = Split(Range("f№Акт").Value, "\")
КВВИД = sh.Range("fПолнаяДіл").value
КВВИД = Replace(КВВИД, ".", "")
            NewFaleName = КВВИД & "-" & sh.Range("fDot").value & "-" & sh.Range("f№ЛК").value & "-" & sh.Range("fВидШаблона").value & "-" & sh.Range("fВидРубки").value ' Trim$(MonthTK(Range("cMonth")) & "_" & СокрTK(Range("cLicVa")) & "_" & Range("сN"))
End Function
Function NewFileFullName(НоваяПапка, ИмяФайла, РасширениеСоздаваемыхФайлов)
NewFileFullName = НоваяПапка & ИмяФайла & РасширениеСоздаваемыхФайлов
End Function
Sub СформироватьДоговоры1()
Dim wb As Workbook
Dim WSName As String
Dim b As Range
Dim ColumnName As String
Set wb = ThisWorkbook
    ПутьШаблона = Replace(wb.FullName, wb.name, ИмяФайлаШаблона)
    НоваяПапка = NewFolderName & Application.PathSeparator
    Dim Row As Range, pi As New ProgressIndicator
    r = Cells(Rows.Count, "A").End(xlUp).Row: rc = r - 2
    If rc < 1 Then MsgBox "Строк для обработки не найдено", vbCritical: Exit Sub

    pi.Show "Формирование договоров": pi.ShowPercents = True: s1 = 10: s2 = 90: p = s1: A = (s2 - s1) / rc
    pi.StartNewAction , s1, "Запуск приложения Microsoft Word"

    ' Dim WA As Word.Application, WD As Word.Document: Set WA = New Word.Application    ' c подключением библиотеки Word
    Dim WA As Object, WD As Object: Set WA = CreateObject("Word.Application")    ' без подключения библиотеки Word
n = ActiveCell.Row

    For Each Row In ActiveSheet.Rows(n & ":" & r)
        With Row
         ipn = Trim$(.Cells(3))
         ipn = Trim$(.Cells(8))
'WSName = "Довітки"
'ColumnName = "NДодат"
'Set b = ЗаголовокСтолбца(WB, WSName, ColumnName)
'With Worksheets("Акты")
'.Activate
'If .AutoFilterMode = True And .FilterMode = True Then
'.Rows(1).AutoFilter
'.Rows(1).AutoFilter
'End If
'Set R = .Columns(b.Column)
'
'v = CellAutoFilterVisible(ipn)
'
'  End With
  Worksheets("Довітки").Activate
            ФИО = Trim$(.Cells(7))
            Filename = НоваяПапка & ФИО & РасширениеСоздаваемыхФайлов

            pi.StartNewAction p, p + A / 3, "Создание нового файла на основании шаблона", ФИО
            Set WD = WA.Documents.Add(ПутьШаблона): DoEvents

            pi.StartNewAction p + A / 3, p + A * 2 / 3, "Замена данных ...", ФИО
            For I = 1 To 21
                FindText = Cells(1, I): ReplaceText = Trim$(.Cells(I))

                ' так почему-то заменяет не всё (не затрагивает таблицу)
                'WA.Selection.Find.Execute FindText, , , , , , , wdFindContinue, False, ReplaceText, True

                pi.line3 = "Заменяется поле " & FindText
                With WD.Range.Find
                    .Text = FindText
                    .Replacement.Text = ReplaceText
                    .Forward = True
                    .Wrap = 1
                    .Format = False: .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    .Execute Replace:=2
                  
                    
                    
                End With
            
    
                
               
  
              
                 
                    
                    
              
                
                DoEvents
  
 
            Next I
'             Set rwend = WD.bookmarks("\EndOfDoc")
'             Set rw = WD.Range(rwend.End - 1, rwend.End - 1)
'
'
'               Set WDTable = WD.Tables.Add(Range:=rw, NumRows:=UBound(v), NumColumns:=6, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
'     wdAutoFitWindow)
'         With WDTable
'        If .Style <> "Изысканная таблица" Then
'            .Style = "Изысканная таблица"
'        End If
'        .ApplyStyleHeadingRows = True
'        .ApplyStyleLastRow = True
'        .ApplyStyleFirstColumn = True
'        .ApplyStyleLastColumn = True
'    End With
''               (Range:=rw, NumRows:=5, NumColumnxts:= _
''        5, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
''        wdAutoFitFixed)
'        k = 1
'        z = Array(50, 40, 40, 120, 150, 30)
'        For i = 1 To UBound(v)
'         For j = 1 To 6
'         Set vv = WDTable.Range.Cells(k).Range
'         vv.Font.Size = 6
'         vv.Font.Bold = True
'         vv.Text = v(i, j)
'         vv.Columns.Width = z(j - 1)
'         k = k + 1
'         Next j
'         Next i
            pi.StartNewAction p + A * 2 / 3, p + A, "Сохранение файла ...", ФИО, " "
            WD.SaveAs Filename: WD.Close False: DoEvents
            p = p + A
        End With
        
        
        
    Next Row

    pi.StartNewAction s2, , "Завершение работы приложения Microsoft Word", " ", " "
    WA.Quit False: pi.Hide
    Msg = "Сформировано " & rc + 2 - n - 1 & " договоров. Все они находятся в папке" & vbNewLine & НоваяПапка
    MsgBox Msg, vbInformation, "Готово"
End Sub
Sub ПечатьWordизExcel(Filename, doc, Optional PageNumber)
Dim wdApp As Object, wdDoc As Object: Set wdApp = CreateObject("Word.Application")



 
wdApp.Visible = True

Set wdDoc = wdApp.Documents.Open(Filename)
 
'Background:=False

'wdDoc.PrintOut Copies:=1
'wdDoc.Close
'wdApp.Quit


nd = 10
КоличествоКопий = fAkt.TextBox_Копий
' Перейти в конец Word-документа



With wdDoc.Content
'.Characters(1).Select
'Set arangenach = Selection
'vr = 6 'wdStory
'wdDoc.Selection.Start = .End - 1
'Set arangekon = wdDoc.Selection.EndKey(Unit:=6)
'Set arangekon = Selection
Select Case doc
Case 1

Set arangeend = wdDoc.Bookmarks("D" & doc + 1)
Set arange = wdDoc.Range( _
Start:=.Start, _
End:=arangeend.Range.Start)
Case nd

Set arangestart = wdDoc.Bookmarks("D" & doc)
'Set arangeend = arangekon


Set arange = wdDoc.Range( _
Start:=arangestart.Range.Start, _
End:=.End - 1)
Case Else
Set arangestart = wdDoc.Bookmarks("D" & doc)
Set arangeend = wdDoc.Bookmarks("D" & doc)
Set arange = wdDoc.Range( _
Start:=arangestart.Range.Start, _
End:=arangeend.Range.Start)
End Select


'arangestart.Range.Start
'arangeend.Range.Start
'arangestart.Select
'Selection.expand (arangeend)


arange.Select
КоличествоСтраницДляПечати = arange.ComputeStatistics(Statistic:=2) - 1


If КоличествоСтраницДляПечати > 1 Then
КонСтраницДляПечати = arange.Information(wdActiveEndPageNumber)
НачСтраницДляПечати = КонСтраницДляПечати - КоличествоСтраницДляПечати
If Application.WorksheetFunction.Even(КоличествоСтраницДляПечати) = КоличествоСтраницДляПечати Then
If КоличествоСтраницДляПечати \ 2 = 1 Then
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(КоличествоКопий)
                    .Range = wdPrintRangeOfPages
                    .Pages = НачСтраницДляПечати
                    .Execute
                End With
                MsgBox ""
                With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(КоличествоКопий)
                    .Range = wdPrintRangeOfPages
                    .Pages = КонСтраницДляПечати
                    .Execute
                End With
Else
'четніе
'For j = НачСтраницДляПечати To КоличествоСтраницДляПечати Step 2
For j = КоличествоСтраницДляПечатиНачСтраницДляПечати To НачСтраницДляПечати Step -2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(КоличествоКопий)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j
MsgBox "Переверните лист"
For j = НачСтраницДляПечати + 1 To КоличествоСтраницДляПечати Step 2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(КоличествоКопий)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j
End If

Else

'For j = НачСтраницДляПечати To КоличествоСтраницДляПечати Step 2
For j = КоличествоСтраницДляПечати To НачСтраницДляПечати Step -2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(КоличествоКопий)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j
MsgBox "Переверните лист"
For j = НачСтраницДляПечати + 1 To Fix(КоличествоСтраницДляПечати / 2) + 2 Step 2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(КоличествоКопий)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j

End If

Else
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(КоличествоКопий)
                    .Range = wdPrintRangeOfPages
                    .Pages = arange.Information(wdActiveEndPageNumber)
                    .Execute
                End With
End If

 
End With
wdDoc.Close
wdApp.Quit
Set wdApp = Nothing




End Sub

Sub ПечатьПоУмолчанию()
If fAkt.ComboBox_ВідШаблона = "РГК" Then
v = Array("1", "7", "4", "#", "5", "6", "11")
Else
v = Array("1", "7", "4", "#", "5", "6")
End If
For I = 0 To UBound(v)
If v(I) = "#" Then
MsgBox "Переверните лист"
Else
fAkt.ListBox_Печать.ListIndex = Val(v(I)) - 1
fAkt.CommandButton13.value = True
End If
Next I
End Sub
Sub Print_3()
Dim Y&
        Y = MsgBox(Prompt:="Какую сторону Акта печатать? Первую - Да, Вторую - Нет", _
                   Buttons:=vbYesNoCancel + vbQuestion, _
                   Title:="Печать акта")
        Select Case Y
            Case vbYes
                With Application.Dialogs(wdDialogFilePrint)
                    .Background = False
                    .NumCopies = CStr(3)
                    .Range = wdPrintRangeOfPages
                    .Pages = CStr(1)
                    .Execute
                End With
            Case vbNo
                With Application.Dialogs(wdDialogFilePrint)
                    .Background = False
                    .NumCopies = CStr(3)
                    .Range = wdPrintRangeOfPages
                    .Pages = CStr(2)
                    .Execute
                End With
            Case vbCancel
               End
        End Select
End Sub

Sub WordПоДодаткам()
'begin
'  try
'    WordApp := CreateOleObject('Word.Application');
'  except
'    // Error....
'    ShowMessage('Error!');
'    Exit;
'  end;
'  try
'    WordApp.Visible := False;
'    WordApp.Documents.Open(AWordDoc);
'    sleep(1000);//
'    intNumberOfPages := WordApp.ActiveDocument.BuiltInDocumentProperties[wdPropertyPages].Value;
'    Form1.ProgressBar1.Max:=intNumberOfPages*10;
'    For intPage:= 1 to intNumberOfPages do
'    begin
'      Form1.ProgressBar1.Position:=intPage*10;
'      If intPage = intNumberOfPages Then
'        begin
'          WordApp.Selection.EndKey(wdStory, wdExtend);
'          WordApp.Selection.Copy();
'        End
'      Else
'        begin
'          WordApp.Selection.GoTo(wdGoToPage, 2);
'          WordApp.Selection.MoveLeft(Unit:=wdCharacter, Count:=1);
'          WordApp.Selection.HomeKey(wdStory, wdExtend);
'          WordApp.Selection.Copy();
'        end;
'//создаем новый док-нт и вставляем в него clipboard
'      WordAppDest:= CreateOleObject('Word.Application');
'      WordAppDest.Visible := False;
'      WordAppDest.Documents.Add;
'      WordAppDest.Selection.Paste();
'      WordAppDest.ActiveDocument.SaveAs('c:\NewFile'+IntToStr(intPage)+'.doc');
'      WordAppDest.Quit(SaveChanges, EmptyParam, EmptyParam);
'      WordApp.Selection.GoTo(wdGoToPage, 2);
'      WordApp.Selection.HomeKey(wdStory, wdExtend);
'      WordApp.Selection.Delete();
'//      ShowMessage('Go on?');
'    end;
'
'  finally
'    SaveChanges := false;
'    WordApp.Quit(SaveChanges, EmptyParam, EmptyParam);
'    ShowMessage('Complete!');
'  end;
'end;
End Sub
Sub splitter()
Dim Counter As Long, Source As Document, Target As Document
Set Source = ActiveDocument
Selection.HomeKey Unit:=wdStory
Pages = Source.BuiltinDocumentProperties(wdPropertyPages)
Counter = 0
While Counter < Pages
Counter = Counter + 1
DocName = "Page" & Format(Counter)
Source.Bookmarks("\Page").Range.Cut
Set Target = Documents.Add
Target.Range.Paste
Target.SaveAs Filename:=DocName
Target.Close
Wend

End Sub
Sub each_sheet_into_new_document()
'помещает каждый лист в отдельный документ
Dim word_App As New Word.Application
Dim word_Doc As Word.Document
Dim word_Shape As Shape
Selection.HomeKey Unit:=wdStory
pages_count = ActiveDocument.Range.ComputeStatistics(wdStatisticPages)
For I = 1 To pages_count
    Set word_Doc = word_App.Documents.Add
    With word_Doc
    If I = pages_count Then
        Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Else
        Selection.ExtendMode = True
        Selection.GoToNext wdGoToPage
        Selection.MoveLeft wdCharacter, 1, False
    End If
        text_for_input = Selection
        Selection.ExtendMode = False
        Selection.MoveRight wdCharacter, 2, False
        .StoryRanges(wdMainTextStory) = text_for_input
        word_App.Visible = True
        If I < 10 Then
            p = "0" & I
        Else
            p = I
        End If
        .SaveAs "C:\Лист-" & p & ".doc"
    End With
Next I
'этот кусок кода переносит shapes в новый документ
        For Each word_Shape In ActiveDocument.Shapes
            'MsgBox ActiveDocument.Shapes.Count
            word_Shape.Select
            p = Selection.Information(wdActiveEndPageNumber)
            If p < 10 Then
                p = "0" & Selection.Information(wdActiveEndPageNumber)
            End If
            Selection.Copy
            word_App.Documents("Лист-" & p & ".doc").Activate
            word_App.Selection.Paste
        Next word_Shape
        For Each doc In word_App.Documents
            doc.Save
        Next doc
Set word_App = Nothing
End Sub



Function NewFolderName1() As String
    NewFolderName1 = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "Договоры, сформированные " & Get_Now)
    MkDir NewFolderName1
End Function

Function CellAutoFilterVisible1(k)
Dim sh As Worksheet

Set sh = Worksheets("Места работи")
Dim sh_cv As Worksheet
'On Error Resume Next
z = 0
With sh
.Rows(1).AutoFilter Field:=1, Criteria1:=k
     If .AutoFilterMode = True And .FilterMode = True Then
        With .AutoFilter.Range.Columns(1)
             Set iFilterRange = _
             .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlVisible)
            
            Set iCell = sh.Range(iFilterRange.Cells(1), iFilterRange.Cells(iFilterRange.Cells.Rows.Count).Offset(columnoffset:=5))
            
        End With
'        .ShowAllData 'Отобразить всё - необязательно
     End If
End With
CellAutoFilterVisible = iCell
End Function
Function Get_Date() As String: Get_Date = Replace(Replace(DateValue(Now), "/", "-"), ".", "-"): End Function
Function Get_Time() As String: Get_Time = Replace(TimeValue(Now), ":", "-"): End Function
'Function Get_Now() As String: Get_Now = Get_Date & " в " & Get_Time: End Function
Function Get_Now() As String: Get_Now = Get_Date: End Function
Function ЗаголовокСтолбца(wb As Workbook, sh As String, NameZag As String) As Range
Dim ABS_WB As Workbook
Dim Lst As Worksheet
Dim Nel_lst As Worksheet
Dim O_lst As Worksheet
Dim ABS_lst As Worksheet
Dim nr As Range
Dim f As Range

Set Nel_lst = wb.Worksheets(sh)
Set b = LastColumn(sh, 1)
Set zags = Nel_lst.Range(Nel_lst.Cells(1, 1), b)
Set ЗаголовокСтолбца = zags.Find(NameZag)

End Function
Function LastColumn(SheetName As String, r As Long) As Range

'Определение последней используемой ячейки в строке r на листе с именем SheetName
Dim sh As Worksheet
Dim EndCell As Range
Set sh = Worksheets(SheetName)
Set EndCell = sh.Cells(r, 256)
Set LastColumn = EndCell.End(xlToLeft)
End Function
Sub Сохраним_Форму()
Dim sh As Worksheet
Dim tsh As Worksheet
Dim wb As Workbook
Dim ls As ListObject
Dim lr As ListRow
Dim lc As ListColumn
Dim lcКод As ListColumn
Dim lcТип As ListColumn
Dim z As Range
Dim hpl As myHuperlink
Dim fZagolovok As Range
Dim fZnach As Range
Dim ИмяТаблициФорма As String
ИмяТаблициФорма = "Форма"
ИмяТаблициРеестр = "Таблица Форма"
ИмяСтолбцаКлючи = "Заголовки"
ИмяСтолбцаПроверки = "Заголовки"
ИмяСтолбцаЗначений = "Значения"


Set wb = ThisWorkbook
Set sh = wb.Worksheets(ИмяТаблициФорма)
Set tsh = wb.Worksheets(ИмяТаблициРеестр)
Set ls = tsh.ListObjects(1)
ИмяШаблона = sh.Range("fDot").value
Set ЗаголовокСтолбцаКлючи = FindAll(sh.Rows(1), ИмяСтолбцаКлючи)
Set ЗаголовокСтолбцаПроверки = FindAll(sh.Rows(1), ИмяСтолбцаПроверки)
Set ЗаголовокСтолбцаЗначения = FindAll(sh.Rows(1), ИмяСтолбцаЗначений)
n = 2
r = LastRow(ИмяТаблициФорма)
'Set lr = ls.ListRows.add
'ls.ListRows(ls.ListRows.Count).Delete
'ls.ListRows(ls.ListRows.Count).Range.Copy
'Set z = Union(lr.Range, ls.ListColumns("Вид DOT").Range)
'z.Select
'ls.ListColumns("Вид DOT").Range.Copy
'ls.ShowTotals = True
'ls.ListColumns("Вид DOT").Total.Copy
'ls.ShowTotals = False


For I = n To r
'Содержит ли гиперссілку
Set fZagolovok = sh.Cells(I, 1)
Set fZnach = sh.Cells(I, 2)

If fZnach.Hyperlinks.Count > 0 Then
hpl.myLink = fZnach.Hyperlinks(1).address
hpl.myCapton = fZnach.Hyperlinks(1).TextToDisplay

End If
If ColumnExistListObject(ls, fZagolovok.value) Then
Set lc = ls.ListColumns(fZagolovok.value)
Else
Set lc = ls.ListColumns.Add
lc.name = fZagolovok.value


End If



Set tZagolovok = tsh.ListObjects("Таблица_Форма").HeaderRowRange
If sh.Cells(I, 2).huperlinks.Count > 0 Then
End If
Next I
End Sub
 Function ColumnExistListObject(ls As ListObject, name As String) As Boolean
On Error Resume Next
Dim lc As ListColumn
Set lc = ls.ListColumns(name)
If Err = 0 Then ColumnExistListObject = True Else ColumnExistListObject = False

 End Function
