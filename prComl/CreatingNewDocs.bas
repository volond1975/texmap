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



Function ��������������������()
Dim wb As Workbook
Dim sh As Worksheet
Dim WSName As String
Dim b As Range
Dim ColumnName As String
Dim ��������������� ' "������.dot"
Dim �������������������������������� ' 21
Dim ��������������������������� ' ".doc"
Dim ��������������� As String
Dim ������������������ As String
Dim ������������������ As String
Dim ��������������������� As Range
Dim ������������������������ As Range
Dim ������������������������ As Range
Dim ���������� As String
On Error Resume Next
 'Range("cDog")

��������������������������� = ".doc"
���������� = "�����"
��������������� = "���������"
������������������ = "���������"
������������������ = "��������"


Set wb = ThisWorkbook
Set sh = wb.Worksheets(����������)
���������� = sh.Range("fDot").value
��������������� = ���������� & ".dot"
Set ��������������������� = FindAll(sh.Rows(1), ���������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)

    ����������� = Replace(wb.FullName, wb.name, ���������������)
    ���������� = NewFolderName & Application.PathSeparator
    Dim Row As Range, pi As New ProgressIndicator
    n = 2
r = LastRow(����������)
    rc = r - n - 2
'    r = Cells(Rows.Count, "A").End(xlUp).row: Rc = r - 2
'    If Rc < 1 Then MsgBox "����� ��� ��������� �� �������", vbCritical: Exit Sub

    pi.Show "������������ ���������": pi.ShowPercents = True: s1 = 10: s2 = 90: p = s1: A = (s2 - s1) / rc
    pi.StartNewAction , s1, "������ ���������� Microsoft Word"

    ' Dim WA As Word.Application, WD As Word.Document: Set WA = New Word.Application    ' c ������������ ���������� Word
    Dim WA As Object, WD As Object: Set WA = CreateObject("Word.Application")    ' ��� ����������� ���������� Word
WA.Visible = True
WA.TrackRevisions = False



'    For Each row In ActiveSheet.Rows(N & ":" & r)
'        With row
'         ipn = Trim$(.Cells(3))
'         ipn = Trim$(.Cells(8))
'WSName = "����"
'ColumnName = "N�����"
'Set b = ����������������(WB, WSName, ColumnName)
'With Worksheets("����")
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
''  v = Split(Range("f����").Value, "\")
'����� = sh.Range("f������ĳ�").Value
'����� = Replace(�����, ".", "")
'            �������� = ����� & "-" & sh.Range("fDot").Value & "-" & sh.Range("f���").Value ' Trim$(MonthTK(Range("cMonth")) & "_" & ����TK(Range("cLicVa")) & "_" & Range("�N"))
            �������� = NewFaleName
'            FileNameDoc = ���������� & �������� & ���������������������������
FileNameDoc = NewFileFullName(����������, ��������, ���������������������������)
'            pi.StartNewAction p, p + a / 3, "�������� ������ ����� �� ��������� �������", ���
            Set WD = WA.Documents.Add(�����������): DoEvents
WD.Revisions.AcceptAll
WD.DeleteAllComments
            pi.StartNewAction p + A / 3, p + A * 2 / 3, "������ ������ ...", ��������
            For I = n To r
            If sh.Cells(I, 1).value <> "" Then
          If HasLinkToObject(sh.Cells(I, 2).Text) Then
          
      Call InsertObjectIntoDOC(WD, sh.Cells(I, 2).Text, "{" & sh.Cells(I, 1) & "}")
          Else
                FindText = "{" & sh.Cells(I, 1) & "}": ReplaceText = Trim$(sh.Cells(I, 2))

                ' ��� ������-�� �������� �� �� (�� ����������� �������)
                'WA.Selection.Find.Execute FindText, , , , , , , wdFindContinue, False, ReplaceText, True

                pi.line3 = "���������� ���� " & FindText
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
           
            pi.StartNewAction p + A * 2 / 3, p + A, "���������� ����� ...", ��������, " "
            WD.SaveAs FileNameDoc: WD.Close False: DoEvents
'            p = p + a
'        End With
'
'
'
'    Next row

    pi.StartNewAction s2, , "���������� ������ ���������� Microsoft Word", " ", " "
    WA.Quit False: pi.Hide
    Msg = "������������ " & rc + 2 - n - 1 & " ���������. ��� ��� ��������� � �����" & vbNewLine & ����������
    �������������������� = ����������
'    MsgBox msg, vbInformation, "������"
End Function










Function NewFolderName() As String
Dim twb As Workbook
Dim fso As Scripting.FileSystemObject
Dim WShConst As Worksheet
Set twb = ThisWorkbook
With twb

Set fso = New Scripting.FileSystemObject
Set WShConst = .Worksheets("�����")
v = Split(WShConst.Range("f���").value, "\")
���� = v(0)
 'WShConst.Range("cN���������").Value
'�������� = WShConst.Range("f�������").Value
'˳������� = WShConst.Range("cLicVa").Value
sl = v(1) '
����� = WShConst.Range("f������ĳ�").value
����� = Replace(�����, ".", "")
p = "��������"
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If
p = "��������\" & ����
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If
p = "��������\" & ���� & "\" & sl
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If
 p = "��������\" & ���� & "\" & sl & "\" & �����
    NewFolderName = Replace(ThisWorkbook.FullName, ThisWorkbook.name, Trim(p))
    Debug.Print NewFolderName
    If Not PathExists(NewFolderName) Then
'    MkDir NewFolderName
    fso.CreateFolder (NewFolderName)
    End If

    
    End With
    
'  FilenName = ThisWorkbook.Path & "\" & "��������" & "\" & Range(z & "NTK") & ".doc"
    
End Function
Function NewFaleName()
Dim shp As Worksheet
Dim lo As ListObject
Dim q As Range
Set wb = ThisWorkbook
Set sh = wb.Worksheets("�����")
���������� = sh.Range("fDot").value
��������������� = ���������� & ".dot"
Set ��������������������� = FindAll(sh.Rows(1), ���������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)

  
    ���������� = NewFolderName & Application.PathSeparator

 'sh.Activate
'  v = Split(Range("f����").Value, "\")
����� = sh.Range("f������ĳ�").value
����� = Replace(�����, ".", "")
            NewFaleName = ����� & "-" & sh.Range("fDot").value & "-" & sh.Range("f���").value & "-" & sh.Range("f����������").value & "-" & sh.Range("f��������").value ' Trim$(MonthTK(Range("cMonth")) & "_" & ����TK(Range("cLicVa")) & "_" & Range("�N"))
End Function
Function NewFileFullName(����������, ��������, ���������������������������)
NewFileFullName = ���������� & �������� & ���������������������������
End Function
Sub ��������������������1()
Dim wb As Workbook
Dim WSName As String
Dim b As Range
Dim ColumnName As String
Set wb = ThisWorkbook
    ����������� = Replace(wb.FullName, wb.name, ���������������)
    ���������� = NewFolderName & Application.PathSeparator
    Dim Row As Range, pi As New ProgressIndicator
    r = Cells(Rows.Count, "A").End(xlUp).Row: rc = r - 2
    If rc < 1 Then MsgBox "����� ��� ��������� �� �������", vbCritical: Exit Sub

    pi.Show "������������ ���������": pi.ShowPercents = True: s1 = 10: s2 = 90: p = s1: A = (s2 - s1) / rc
    pi.StartNewAction , s1, "������ ���������� Microsoft Word"

    ' Dim WA As Word.Application, WD As Word.Document: Set WA = New Word.Application    ' c ������������ ���������� Word
    Dim WA As Object, WD As Object: Set WA = CreateObject("Word.Application")    ' ��� ����������� ���������� Word
n = ActiveCell.Row

    For Each Row In ActiveSheet.Rows(n & ":" & r)
        With Row
         ipn = Trim$(.Cells(3))
         ipn = Trim$(.Cells(8))
'WSName = "������"
'ColumnName = "N�����"
'Set b = ����������������(WB, WSName, ColumnName)
'With Worksheets("����")
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
  Worksheets("������").Activate
            ��� = Trim$(.Cells(7))
            Filename = ���������� & ��� & ���������������������������

            pi.StartNewAction p, p + A / 3, "�������� ������ ����� �� ��������� �������", ���
            Set WD = WA.Documents.Add(�����������): DoEvents

            pi.StartNewAction p + A / 3, p + A * 2 / 3, "������ ������ ...", ���
            For I = 1 To 21
                FindText = Cells(1, I): ReplaceText = Trim$(.Cells(I))

                ' ��� ������-�� �������� �� �� (�� ����������� �������)
                'WA.Selection.Find.Execute FindText, , , , , , , wdFindContinue, False, ReplaceText, True

                pi.line3 = "���������� ���� " & FindText
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
'        If .Style <> "���������� �������" Then
'            .Style = "���������� �������"
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
            pi.StartNewAction p + A * 2 / 3, p + A, "���������� ����� ...", ���, " "
            WD.SaveAs Filename: WD.Close False: DoEvents
            p = p + A
        End With
        
        
        
    Next Row

    pi.StartNewAction s2, , "���������� ������ ���������� Microsoft Word", " ", " "
    WA.Quit False: pi.Hide
    Msg = "������������ " & rc + 2 - n - 1 & " ���������. ��� ��� ��������� � �����" & vbNewLine & ����������
    MsgBox Msg, vbInformation, "������"
End Sub
Sub ������Word��Excel(Filename, doc, Optional PageNumber)
Dim wdApp As Object, wdDoc As Object: Set wdApp = CreateObject("Word.Application")



 
wdApp.Visible = True

Set wdDoc = wdApp.Documents.Open(Filename)
 
'Background:=False

'wdDoc.PrintOut Copies:=1
'wdDoc.Close
'wdApp.Quit


nd = 10
��������������� = fAkt.TextBox_�����
' ������� � ����� Word-���������



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
�������������������������� = arange.ComputeStatistics(Statistic:=2) - 1


If �������������������������� > 1 Then
������������������� = arange.Information(wdActiveEndPageNumber)
������������������� = ������������������� - ��������������������������
If Application.WorksheetFunction.Even(��������������������������) = �������������������������� Then
If �������������������������� \ 2 = 1 Then
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(���������������)
                    .Range = wdPrintRangeOfPages
                    .Pages = �������������������
                    .Execute
                End With
                MsgBox ""
                With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(���������������)
                    .Range = wdPrintRangeOfPages
                    .Pages = �������������������
                    .Execute
                End With
Else
'�����
'For j = ������������������� To �������������������������� Step 2
For j = ��������������������������������������������� To ������������������� Step -2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(���������������)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j
MsgBox "����������� ����"
For j = ������������������� + 1 To �������������������������� Step 2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(���������������)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j
End If

Else

'For j = ������������������� To �������������������������� Step 2
For j = �������������������������� To ������������������� Step -2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(���������������)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j
MsgBox "����������� ����"
For j = ������������������� + 1 To Fix(�������������������������� / 2) + 2 Step 2
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(���������������)
                    .Range = wdPrintRangeOfPages
                    .Pages = j
                    .Execute
                End With
Next j

End If

Else
With wdApp.Dialogs(88)
                    .Background = False
                    .NumCopies = CStr(���������������)
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

Sub �����������������()
If fAkt.ComboBox_³�������� = "���" Then
v = Array("1", "7", "4", "#", "5", "6", "11")
Else
v = Array("1", "7", "4", "#", "5", "6")
End If
For I = 0 To UBound(v)
If v(I) = "#" Then
MsgBox "����������� ����"
Else
fAkt.ListBox_������.ListIndex = Val(v(I)) - 1
fAkt.CommandButton13.value = True
End If
Next I
End Sub
Sub Print_3()
Dim Y&
        Y = MsgBox(Prompt:="����� ������� ���� ��������? ������ - ��, ������ - ���", _
                   Buttons:=vbYesNoCancel + vbQuestion, _
                   Title:="������ ����")
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

Sub Word����������()
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
'//������� ����� ���-�� � ��������� � ���� clipboard
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
'�������� ������ ���� � ��������� ��������
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
        .SaveAs "C:\����-" & p & ".doc"
    End With
Next I
'���� ����� ���� ��������� shapes � ����� ��������
        For Each word_Shape In ActiveDocument.Shapes
            'MsgBox ActiveDocument.Shapes.Count
            word_Shape.Select
            p = Selection.Information(wdActiveEndPageNumber)
            If p < 10 Then
                p = "0" & Selection.Information(wdActiveEndPageNumber)
            End If
            Selection.Copy
            word_App.Documents("����-" & p & ".doc").Activate
            word_App.Selection.Paste
        Next word_Shape
        For Each doc In word_App.Documents
            doc.Save
        Next doc
Set word_App = Nothing
End Sub



Function NewFolderName1() As String
    NewFolderName1 = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "��������, �������������� " & Get_Now)
    MkDir NewFolderName1
End Function

Function CellAutoFilterVisible1(k)
Dim sh As Worksheet

Set sh = Worksheets("����� ������")
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
'        .ShowAllData '���������� �� - �������������
     End If
End With
CellAutoFilterVisible = iCell
End Function
Function Get_Date() As String: Get_Date = Replace(Replace(DateValue(Now), "/", "-"), ".", "-"): End Function
Function Get_Time() As String: Get_Time = Replace(TimeValue(Now), ":", "-"): End Function
'Function Get_Now() As String: Get_Now = Get_Date & " � " & Get_Time: End Function
Function Get_Now() As String: Get_Now = Get_Date: End Function
Function ����������������(wb As Workbook, sh As String, NameZag As String) As Range
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
Set ���������������� = zags.Find(NameZag)

End Function
Function LastColumn(SheetName As String, r As Long) As Range

'����������� ��������� ������������ ������ � ������ r �� ����� � ������ SheetName
Dim sh As Worksheet
Dim EndCell As Range
Set sh = Worksheets(SheetName)
Set EndCell = sh.Cells(r, 256)
Set LastColumn = EndCell.End(xlToLeft)
End Function
Sub ��������_�����()
Dim sh As Worksheet
Dim tsh As Worksheet
Dim wb As Workbook
Dim ls As ListObject
Dim lr As ListRow
Dim lc As ListColumn
Dim lc��� As ListColumn
Dim lc��� As ListColumn
Dim z As Range
Dim hpl As myHuperlink
Dim fZagolovok As Range
Dim fZnach As Range
Dim ��������������� As String
��������������� = "�����"
���������������� = "������� �����"
��������������� = "���������"
������������������ = "���������"
������������������ = "��������"


Set wb = ThisWorkbook
Set sh = wb.Worksheets(���������������)
Set tsh = wb.Worksheets(����������������)
Set ls = tsh.ListObjects(1)
���������� = sh.Range("fDot").value
Set ��������������������� = FindAll(sh.Rows(1), ���������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)
Set ������������������������ = FindAll(sh.Rows(1), ������������������)
n = 2
r = LastRow(���������������)
'Set lr = ls.ListRows.add
'ls.ListRows(ls.ListRows.Count).Delete
'ls.ListRows(ls.ListRows.Count).Range.Copy
'Set z = Union(lr.Range, ls.ListColumns("��� DOT").Range)
'z.Select
'ls.ListColumns("��� DOT").Range.Copy
'ls.ShowTotals = True
'ls.ListColumns("��� DOT").Total.Copy
'ls.ShowTotals = False


For I = n To r
'�������� �� ����������
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



Set tZagolovok = tsh.ListObjects("�������_�����").HeaderRowRange
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
