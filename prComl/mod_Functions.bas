Attribute VB_Name = "mod_Functions"


 '---------------------------------------------------------------------------------------
 ' Module        : mod_Functions
 ' Автор     : EducatedFool  (Игорь)                    Дата: 26.03.2012
 ' Разработка макросов для Excel, Word, CorelDRAW. Быстро, профессионально, недорого.
 ' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
 ' Реквизиты для оплаты работы: http://ExcelVBA.ru/payments
 '---------------------------------------------------------------------------------------
 Option Compare Text
 Option Private Module

 #If Win64 Then
     #If VBA7 Then    ' Windows x64, Office 2010
         Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
                 (ByVal hWnd As LongLong, ByVal pszPath As String, ByVal psa As Any) As LongLong
     #Else    ' Windows x64,Office 2003-2007
         Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
                                              (ByVal hWnd As LongLong, ByVal pszPath As String, _
                                               ByVal psa As Any) As LongLong
     #End If
 #Else
     #If VBA7 Then    ' Windows x86, Office 2010
         Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
                 (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
     #Else    ' Windows x86, Office 2003-2007
         Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
                                              (ByVal hWnd As Long, ByVal pszPath As String, _
                                               ByVal psa As Any) As Long
     #End If
 #End If

 Sub CtrlShiftV()    ' PasteFormulasForSeparateLetters
     On Error Resume Next: Err.Clear
     With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
         .GetFromClipboard
         txt$ = .GetText
     End With
     CopyFormulas txt$
 End Sub

 Sub CopyFormulas(ByVal txt)
     On Error Resume Next: Err.Clear
     Dim ra As Range, n As Long, nn As Long, cell As Range

     Application.ScreenUpdating = False
     Set ra = Selection
     For Each cell In ra.Cells
         n = n + 1
         If n = 1 Then
             addr = cell.address(1, 1, xlA1)
             cell.value = txt
             cell.Font.Color = vbWhite
             cell.Font.Size = 1
         Else
             If cell.address = cell.MergeArea.Cells(1).address Then
                 k = k + 1: cell.NumberFormat = "General"
                 cell.Formula = "=MID(" & addr & "," & k & ",1)"
             End If
         End If
     Next cell
     If Err = 0 Then Shell "Cmd.exe /c echo " & Chr(7), vbHide
     Application.ScreenUpdating = True
 End Sub

 Sub Enable_HotKeys()
     ' назначает комбинации клавиш, если соответствующие опция включены в настройках программы
     On Error Resume Next
     With Application
         If SettingsBoolean("CheckBox_PasteFormulasForSeparateLetters") Then .OnKey "^+v", "CtrlShiftV" Else .OnKey "^+v"
         If SettingsBoolean("CheckBox_InsertTableLinks") Then .OnKey "^+t", "CtrlShiftT" Else .OnKey "^+t"
         If SettingsBoolean("CheckBox_InsertImageLinks") Then .OnKey "^+i", "CtrlShiftI" Else .OnKey "^+i"
     End With
 End Sub

 Sub Disable_HotKeys()
     On Error Resume Next: Err.Clear
     Application.OnKey "^+v"
 End Sub

 Function SpecialCells_TypeConstants(ByRef ra As Range) As Range
     ' возвращает диапазон, содержащий все заполненные ячейки диапазона ra
     On Error Resume Next: en& = Err.Number
     If ra.Worksheet.ProtectContents Then    ' если лист защищён
         Dim cell As Range
         ' перебираем все ячейки в диапазоне
         For Each cell In Intersect(ra, ra.Worksheet.UsedRange).Cells
             If Trim(cell.value) <> "" Then    ' если ячейка непустая
                 ' то добавляем её в результат
                 If SpecialCells_TypeConstants Is Nothing Then
                     Set SpecialCells_TypeConstants = cell
                 Else
                     Set SpecialCells_TypeConstants = Union(SpecialCells_TypeConstants, cell)
                 End If
             End If
         Next cell

     Else    ' если защита листа не установлена - используем штатные средства Excel
         Set SpecialCells_TypeConstants = ra.SpecialCells(xlCellTypeConstants)
     End If
     If en& = 0 Then Err.Clear
 End Function

 Function SpecialCells_VisibleRows(ByRef ra As Range) As Range
     On Error Resume Next: en& = Err.Number
     If ra.Worksheet.ProtectContents Then
         Dim ro As Range
         For Each ro In Intersect(ra, ra.Worksheet.UsedRange.EntireRow).Rows
             If ro.EntireRow.Hidden = False Then
                 If SpecialCells_VisibleRows Is Nothing Then
                     Set SpecialCells_VisibleRows = ro
                 Else
                     Set SpecialCells_VisibleRows = Union(SpecialCells_VisibleRows, ro)
                 End If
             End If
         Next ro
     Else
         Set SpecialCells_VisibleRows = ra.SpecialCells(xlCellTypeVisible)
     End If
     If en& = 0 Then Err.Clear
 End Function

 Function RenderString(ByVal txt$, ByRef options As Dictionary) As String
     On Error Resume Next: Arr = options.Keys
     For I = LBound(Arr) To UBound(Arr)
         txt$ = Replace(txt$, Arr(I), options(Arr(I)))
     Next I
     RenderString = txt$
 End Function

 Function CreatePathForFile(ByVal OldFilename$, ByRef options As Dictionary) As String
     On Error Resume Next: Err.Clear
     mask$ = OUTPUT_MASK$    ' f.e., {str} - {filename}.{ext}
     If Len(TMP_OUTPUT_MASK$) Then mask$ = TMP_OUTPUT_MASK$

     ShortOldFilename$ = Replace(OldFilename$, TEMPLATES_FOLDER$, "")

     subfolder$ = Left(ShortOldFilename$, InStrRev(ShortOldFilename$, "\") - 1)
     If Len(subfolder$) Then subfolder$ = subfolder$ & "\"

     Filename$ = Dir(OldFilename$)
     Filename$ = Left(Filename$, InStrRev(Filename$, ".") - 1)

     If Filename$ Like "*{print=#*}*" Then
         pcc& = Val(Split(Filename$, "{print=")(1))
         Filename$ = Replace(Filename$, "{print=" & pcc& & "}", "")
         options("{%pcc%}") = pcc&
     End If

     options("{%filename%}") = RenderString(Filename$, options)
     options("{%ext%}") = GetExtensionForNewFile(ShortOldFilename$)

     NewFilename$ = OUTPUT_FOLDER$ & subfolder$ & Replace_symbols2(RenderString(mask$, options))

     ' создание папки для файла
     NewFolderPath$ = Left(NewFilename$, InStrRev(NewFilename$, "\"))
     If Len(Dir(NewFolderPath$, vbDirectory)) = 0 Then    ' если папка отсутствует
         SHCreateDirectoryEx Application.hWnd, NewFolderPath$, ByVal 0&     ' создаём путь
     End If

     If Val(Application.Version) > 11 And PRINT_TO_PDF Then    ' вывод в ПДФ
         If TemplateType(OldFilename$) <> "TXT" Then
             NewFilename$ = Left(NewFilename$, InStrRev(NewFilename$, ".") - 1) & ".pdf"
         End If
     End If

     CreatePathForFile = NewFilename$
 End Function

 Function GetExtensionForNewFile(ByVal Filename$)
     On Error Resume Next: Err.Clear
     Select Case Extension(Filename$)
         Case "XLT": GetExtensionForNewFile = "XLS"
         Case "XLTM": GetExtensionForNewFile = "XLSM"
         Case "XLTX": GetExtensionForNewFile = "XLSX"
         Case "DOT": GetExtensionForNewFile = "DOC"
         Case "DOTM": GetExtensionForNewFile = "DOCM"
         Case "DOTX": GetExtensionForNewFile = "DOCX"

         Case Else: GetExtensionForNewFile = Extension(Filename$)
     End Select
 End Function

 Function GetFileFormatForNewFile(ByVal Filename$)
     On Error Resume Next: Err.Clear
     Select Case Extension(Filename$)
         Case "CSV": GetFileFormatForNewFile = xlCSV
         Case "XLS": GetFileFormatForNewFile = xlWorkbookNormal
         Case "XLSM": GetFileFormatForNewFile = xlOpenXMLWorkbookMacroEnabled
         Case "XLSX": GetFileFormatForNewFile = xlOpenXMLWorkbook
         Case "DOC": GetFileFormatForNewFile = 0    ' wdFormatDocument
         Case "DOCM": GetFileFormatForNewFile = 13    ' wdFormatXMLDocumentMacroEnabled
         Case "DOCX": GetFileFormatForNewFile = 12    ' wdFormatXMLDocument
     End Select
 End Function

 Function ReadOptions(ByRef ro As Range) As Dictionary
     ' возвращает коллекцию значений для подстановки
     Set ReadOptions = New Dictionary: Dim cell As Range, KeysRange As Range, Key$, txt$
     On Error Resume Next
     Set KeysRange = SpecialCells_TypeConstants(ro.Worksheet.Rows(HEADER_ROW))
     For Each cell In KeysRange.Cells
         Key$ = Trim(cell)
         If Not Key$ Like "*}" Then Key$ = Key$ & "}"
         If Not Key$ Like "{*" Then Key$ = "{" & Key$
         txt$ = Intersect(ro.EntireRow, cell.EntireColumn).Text
         If cell.EntireColumn.Hidden Then txt$ = Intersect(ro.EntireRow, cell.EntireColumn).value

         If Len(txt$) < 152 Then    ' ограничение длины заменяемой строки - 255 символов
             ReadOptions.Add Key$, txt$
         Else
             LenStep& = 250 - Len(Key$) - 12    ' немного индийского кода:) но работает!
             baseKey$ = Key$
             For I = 1 To Len(txt$) Step LenStep&
                 txt_part$ = Mid(txt$, I, LenStep&)
                 newkey$ = baseKey$ & "{l=" & I & "}"
                 If I + LenStep& - 1 < Len(txt$) Then txt_part$ = txt_part$ & newkey$
                 ReadOptions.Add Key$, txt_part$
                 Key$ = newkey$
             Next I
         End If
     Next cell

     AddNamedRangesIntoDictionary ReadOptions, ro.Worksheet.Parent

     ReadOptions.Add "{%str%}", ro.Row
     ReadOptions.Add "{%date%}", Format(Now, "YYYY-MM-DD")
     ReadOptions.Add "{%shortdate%}", Format(Now, "YYMMDD")

     ReadOptions.Add "{%longdate%}", Format(Now, "DD MMMM YYYY")
     ReadOptions.Add "{%time%}", Format(Now, "HH-NN-SS")
     ReadOptions.Add "{%shorttime%}", Format(Now, "HHNNSS")

     ReadOptions.Add "{%datetime%}", Format(Now, "YYYY-MM-DD HH-NN-SS")
     ReadOptions.Add "{%shortdatetime%}", Format(Now, "YYMMDD-HHNNSS")
     ReadOptions.Add "{%longdatetime%}", Format(Now, "DD MMMM YYYY HH-NN-SS")

     ReadOptions.Add "{%sheet_name%}", ro.Worksheet.name
     ReadOptions.Add "{%sheet_index%}", ro.Worksheet.Index
     wbname$ = ro.Worksheet.Parent.name: If wbname$ Like "*.*" Then wbname$ = Left(wbname$, InStrRev(wbname$, ".") - 1)
     ReadOptions.Add "{%workbook_name%}", wbname$
 End Function

 Sub AddNamedRangesIntoDictionary(ByRef dict As Dictionary, ByRef wb As Workbook)
     On Error Resume Next

     ' ==================================
     PrintCopies_FieldName$ = Trim(Settings("TextBox_PrintCopies_FieldName", ""))
     If Not PrintCopies_FieldName$ Like String(Len(PrintCopies_FieldName$), "#") Then
         If Not PrintCopies_FieldName$ Like "*}" Then PrintCopies_FieldName$ = PrintCopies_FieldName$ & "}"
         If Not PrintCopies_FieldName$ Like "{*" Then PrintCopies_FieldName$ = "{" & PrintCopies_FieldName$
     End If

     If Len(PrintCopies_FieldName$) > 2 Or Val(PrintCopies_FieldName$) > 0 Then
         dict.Add "{%PrintCopiesCount%}", PrintCopies_FieldName$
     End If
     ' ==================================

     Dim n As name, cell As Range
     For Each n In wb.Names
         Set cell = Nothing: Set cell = n.RefersToRange.Cells(1)
         If Not cell Is Nothing Then
             Key$ = "{=" & n.name & "}"
             txt$ = cell.Text

             If Len(txt$) < 152 Then    ' ограничение длины заменяемой строки - 255 символов
                 dict.Add Key$, txt$
             Else
                 LenStep& = 250 - Len(Key$) - 12    ' немного индийского кода:) но работает!
                 baseKey$ = Key$
                 For I = 1 To Len(txt$) Step LenStep&
                     txt_part$ = Mid(txt$, I, LenStep&)
                     newkey$ = baseKey$ & "{l=" & I & "}"
                     If I + LenStep& - 1 < Len(txt$) Then txt_part$ = txt_part$ & newkey$
                     dict.Add Key$, txt_part$
                     Key$ = newkey$
                 Next I
             End If
             ' --------------
             Key$ = "{=" & cell.address(0, 0) & "}"
             txt$ = wb.ActiveSheet.Range(cell.address).Text

             If Len(txt$) < 152 Then    ' ограничение длины заменяемой строки - 255 символов
                 dict.Add Key$, txt$
             Else
                 LenStep& = 250 - Len(Key$) - 12    ' немного индийского кода:) но работает!
                 baseKey$ = Key$
                 For I = 1 To Len(txt$) Step LenStep&
                     txt_part$ = Mid(txt$, I, LenStep&)
                     newkey$ = baseKey$ & "{l=" & I & "}"
                     If I + LenStep& - 1 < Len(txt$) Then txt_part$ = txt_part$ & newkey$
                     dict.Add Key$, txt_part$
                     Key$ = newkey$
                 Next I
             End If
         End If
     Next

     ' заменяем коды символов на сами символы
     Arr = dict.Keys
     For I = LBound(Arr) To UBound(Arr)
         Key$ = Arr(I)
         txt$ = "": txt$ = dict(Key$)

         If txt$ Like "*{chr#*}*" Then
             txt = Replace(txt, "{chr10}", Chr(10))
             txt = Replace(txt, "{chr11}", Chr(11))
             txt = Replace(txt, "{chr13}", Chr(13))
             txt = Replace(txt, "{chr1310}", vbNewLine)
             dict(Key$) = txt$
         End If
     Next I

 End Sub

 Function ReadMultirowOptions(ByRef ra As Range) As Dictionary
     ' возвращает коллекцию значений для подстановки
     Set ReadMultirowOptions = New Dictionary: Dim ro As Range, cell As Range, KeysRange As Range, Key$, txt$, rn&
     On Error Resume Next
     Set KeysRange = SpecialCells_TypeConstants(ra.Worksheet.Rows(HEADER_ROW))
     For Each ro In ra.Rows    ' перебираем все выделенные строки
         rn = rn + 1
         For Each cell In KeysRange.Cells
             Key$ = Trim(cell)
             If Not Key$ Like "*}" Then Key$ = Key$ & "}"
             If Not Key$ Like "{*" Then Key$ = "{" & Key$

             key2$ = Left(Key$, Len(Key$) - 1) & "#" & rn & "}"
             txt$ = Intersect(ro.EntireRow, cell.EntireColumn).Text
             If cell.EntireColumn.Hidden Then txt$ = Intersect(ro.EntireRow, cell.EntireColumn).value

             If Len(txt$) < 152 Then    ' ограничение длины заменяемой строки - 255 символов
                 If rn = 1 Then ReadMultirowOptions.Add Key$, txt$
                 ReadMultirowOptions.Add key2$, txt$
             Else
                 LenStep& = 250 - Len(Key$) - 12    ' немного индийского кода:) но работает!
                 baseKey$ = Key$
                 For I = 1 To Len(txt$) Step LenStep&
                     txt_part$ = Mid(txt$, I, LenStep&)
                     newkey$ = baseKey$ & "{l=" & I & "}"
                     If I + LenStep& - 1 < Len(txt$) Then txt_part$ = txt_part$ & newkey$
                     If rn = 1 Then ReadMultirowOptions.Add Key$, txt_part$
                     ReadMultirowOptions.Add key2$, txt_part$
                     Key$ = newkey$
                 Next I
             End If
         Next cell
         str_txt$ = str_txt$ & "," & ro.Row
     Next ro

     AddNamedRangesIntoDictionary ReadMultirowOptions, ra.Worksheet.Parent

     ReadMultirowOptions.Add "{%str%}", Mid(str_txt$, 2)
     ReadMultirowOptions.Add "{%rc%}", rn

     ReadMultirowOptions.Add "{%date%}", Format(Now, "YYYY-MM-DD")
     ReadMultirowOptions.Add "{%shortdate%}", Format(Now, "YYMMDD")

     ReadMultirowOptions.Add "{%longdate%}", Format(Now, "DD MMMM YYYY")
     ReadMultirowOptions.Add "{%time%}", Format(Now, "HH-NN-SS")
     ReadMultirowOptions.Add "{%shorttime%}", Format(Now, "HHNNSS")

     ReadMultirowOptions.Add "{%datetime%}", Format(Now, "YYYY-MM-DD HH-NN-SS")
     ReadMultirowOptions.Add "{%shortdatetime%}", Format(Now, "YYMMDD-HHNNSS")
     ReadMultirowOptions.Add "{%longdatetime%}", Format(Now, "DD MMMM YYYY HH-NN-SS")

     ReadMultirowOptions.Add "{%sheet_name%}", ro.Worksheet.name
     ReadMultirowOptions.Add "{%sheet_index%}", ro.Worksheet.Index
     wbname$ = ro.Worksheet.Parent.name: If wbname$ Like "*.*" Then wbname$ = Left(wbname$, InStrRev(wbname$, ".") - 1)
     ReadMultirowOptions.Add "{%workbook_name%}", wbname$
 End Function

 Function CollectionOfRowsBlocks(ByRef ra As Range) As Collection
     ' получает диапазон строк ra, ищет в столбце ComboBox_Multirow_GroupColumn уникальные значения,
     ' разбивает диапазон на блоки строк, по каждому из уникальных значений
     On Error Resume Next: Err.Clear
     Set CollectionOfRowsBlocks = New Collection
     Dim cell As Range, coll As New Collection, txt$, block As Range

     If SettingsBoolean("CheckBox_Multirow_GroupRows") Then
         col& = Val(Settings("ComboBox_Multirow_GroupColumn"))
         If col& = 0 Then
             Msg$ = "В настройках программы включен режим «MiltiRow» с опцией" & vbNewLine & _
                    "«Группировать строки по заданному столбцу»" & vbNewLine & vbNewLine & _
                    "А номер столбца, по которому надо группировать строки, — не указан." & vbNewLine & vbNewLine & _
                    "Измените настройки программы, и снова запустите формирование документов."
             MsgBox Msg, vbExclamation, "Не задан столбец, по которому группировать строки"
             ShowSettingsPage
             F_Settings.MultiPage_Options.value = 4
             F_Settings.ComboBox_Multirow_GroupColumn.SetFocus
             F_Settings.ComboBox_Multirow_GroupColumn.BackColor = vbRed
             Exit Function
         End If

         For Each cell In Intersect(ra.EntireRow, ra.Worksheet.Columns(col&)).Cells
             txt$ = Trim(cell): If Len(txt$) Then coll.Add txt$, txt$
         Next cell

         If coll.Count = 0 Then
             Msg$ = "В настройках программы включен режим «MiltiRow» с опцией" & vbNewLine & _
                    "«Группировать строки по заданному столбцу»" & vbNewLine & vbNewLine & _
                    "Указан номер столбца, по которому надо группировать строки: «" & col& & "»" & vbNewLine & vbNewLine & _
                    "В этом столбце, в выбранных строках, программа не нашла ни одной заполненной ячейки." & vbNewLine & _
                    "Измените настройки программы, и снова запустите формирование документов."
             MsgBox Msg, vbExclamation, "Не задан столбец, по которому группировать строки"
             ShowSettingsPage
             F_Settings.MultiPage_Options.value = 4
             F_Settings.ComboBox_Multirow_GroupColumn.SetFocus
             F_Settings.ComboBox_Multirow_GroupColumn.BackColor = vbRed
             Exit Function
         End If

         For Each v In coll
             Set block = Nothing
             For Each cell In Intersect(ra.EntireRow, ra.Worksheet.Columns(col&)).Cells
                 If Trim(cell) = v Then
                     If block Is Nothing Then Set block = cell Else Set block = Union(block, cell)
                 End If
             Next cell
             If block Is Nothing Then
                 MsgBox "Ошибка группировки строк в режиме Multirow", vbCritical, "Обратитесь к разработчику программы"
                 Exit Function
             Else
                 CollectionOfRowsBlocks.Add block.EntireRow
             End If
         Next v

     Else    ' возвращаем один блок - со всеми строками
         CollectionOfRowsBlocks.Add ra
     End If
 End Function

 Private Sub test_CollectionOfRowsBlocks()
     On Error Resume Next: Err.Clear
     For Each Item In CollectionOfRowsBlocks(ActiveSheet.UsedRange.Offset(1))
         Debug.Print Item.address
     Next
 End Sub

 Function UniqueValuesFromColumn(ByVal ra As Range) As Collection
     ' перебирает все значения в диапазоне ra в поисках уникальных значений.
     ' Возвращает двумерный массив, содержащий уникальные значения из диапазона ra
     Set UniqueValuesFromColumn = coll
 End Function

 Function CreateAndFill_XLS(ByVal TemplateFilename$, ByVal NewFilename$, _
                            ByRef options As Dictionary, Optional ByRef pi As ProgressIndicator) As Boolean

     On Error Resume Next: Err.Clear
     Dim wb As Workbook, sh As Worksheet, nam As name, ra As Range
     pi.line3 = "Файл: " & Dir(TemplateFilename$, vbNormal)

     calc = Application.Calculation
     Application.Calculation = xlCalculationManual
     Application.DisplayAlerts = False

     If TemplateType(TemplateFilename$) Like "*template*" Then
         pi.line2 = "Создание документа Excel по шаблону ..."
         Set wb = Application.Workbooks.Add(TemplateFilename$)
     Else
         pi.line2 = "Открытие исходного документа Excel ..."
         Set wb = Application.Workbooks.Open(TemplateFilename$, False, True)
     End If

     '  Main_PI.Log "Документ создан? " & Not (WB Is Nothing)

     If MULTIROW_MODE Then    ' размножение специальных строк в шаблоне
         Dim rc&: rc = Val(options("{%rc%}"))
         If rc& = 0 Then Main_PI.Log vbTab & "Ошибка при подготовке документа Excel: rc& = 0": Exit Function
         pi.line2 = "Добавление строк (" & rc& & " шт.) - режим MULTIROW ..."

         For Each nam In wb.Names
             If nam.name Like "MultiRow*" Then
                 Set ra = Nothing: Set ra = nam.RefersToRange.EntireRow
                 If Not ra Is Nothing Then
                     For I = 1 To rc&
                         ra.Offset(I).Insert Shift:=xlDown
                         ra.Copy ra.Offset(I)
                         ra.Offset(I).Replace "#}", "#" & I & "}", xlPart
                         ra.Offset(I).Replace "{%index%}", I, xlPart
                     Next I
                     ra.EntireRow.Delete
                 End If
             End If
         Next
     End If

     txt_Line2$ = "Подстановка значений в созданный по шаблону документ ..."
     pi.line2 = txt_Line2$
     Arr = options.Keys
     Dim RIC As Boolean: RIC = REPLACE_IN_COLON

     For I = LBound(Arr) To UBound(Arr)
         Key$ = Arr(I)
         txt$ = options(Arr(I))

         For Each sh In wb.Worksheets
             sh.UsedRange.Replace Key$, txt$, xlPart, , False

             If RIC Then
                 With sh.PageSetup
                     .LeftFooter = Replace(.LeftFooter, Key$, txt$, , , vbTextCompare)
                     .LeftHeader = Replace(.LeftHeader, Key$, txt$, , , vbTextCompare)
                     .CenterFooter = Replace(.CenterFooter, Key$, txt$, , , vbTextCompare)
                     .CenterHeader = Replace(.CenterHeader, Key$, txt$, , , vbTextCompare)
                     .RightFooter = Replace(.RightFooter, Key$, txt$, , , vbTextCompare)
                     .RightHeader = Replace(.RightHeader, Key$, txt$, , , vbTextCompare)
                 End With
             End If
         Next sh

         If I Mod IIf(RIC, 5, 30) = 0 Then
             pi.line2 = txt_Line2$ & "  (выполнено " & Format(I / UBound(Arr), "0%") & ")"
         End If
         DoEvents
     Next I

     pi.line2 = "Вычисление (пересчёт) формул ..."
     For Each sh In wb.Worksheets
         sh.Calculate
         If SettingsBoolean("CheckBox_FormulasToValues") Then sh.UsedRange.value = sh.UsedRange.value
     Next sh


     pi.line2 = "Сохранение заполненного документа ..."
     pi.line3 = "Новое имя файла: " & Split(NewFilename$, "\")(UBound(Split(NewFilename$, "\")))
     Main_PI.Log vbTab & "Сохранение созданного файла: " & Replace(NewFilename$, OUTPUT_FOLDER$, "...\")
     pi.FP.Repaint

     If Val(Application.Version) > 11 And PRINT_TO_PDF Then    ' вывод в ПДФ
         wb.ExportAsFixedFormat xlTypePDF, NewFilename$
     Else    ' обычное сохранение файла Excel
         File_Format = GetFileFormatForNewFile(NewFilename$)
         If Len(File_Format) Then
             wb.SaveAs NewFilename$, Val(File_Format)
         Else
             wb.SaveAs NewFilename$
         End If
     End If
     If IMMEDIATE_PRINTOUT Then wb.PrintOut , , PrintCopiesCount(options)
     wb.Close False

     CreateAndFill_XLS = Err = 0

     Application.Calculation = calc
     Application.DisplayAlerts = True
 End Function

 Function PrintCopiesCount(ByRef options As Dictionary) As Long
     On Error Resume Next: en& = Err.Number
     PrintCopiesCount = 1

     PrintCopiesField$ = options("{%PrintCopiesCount%}")
     If PrintCopiesField$ Like "{*?}" Then
         CopiesCount& = Fix(Val(options(PrintCopiesField$)))
     Else
         CopiesCount& = Fix(Val(PrintCopiesField$))
     End If
     If CopiesCount& > 0 Then PrintCopiesCount = CopiesCount&

     pcc = options("{%pcc%}")
     If pcc <> "" Then PrintCopiesCount = Val(pcc)

     If en& = 0 Then Err.Clear    ' Debug.Print "PrintCopiesCount = " & PrintCopiesCount
 End Function

 Function CreateAndFill_DOC(ByVal TemplateFilename$, ByVal NewFilename$, _
                            ByRef options As Dictionary, Optional ByRef pi As ProgressIndicator) As Boolean
     On Error Resume Next: Err.Clear
     Dim doc As Object, ecount As Long, bm As Object, myStoryRange As Object
     pi.line3 = "Файл: " & Dir(TemplateFilename$, vbNormal)

     If TemplateType(TemplateFilename$) Like "*template*" Then
         pi.line2 = "Создание документа Word по шаблону ..."
         Set doc = WA.Documents.Add(TemplateFilename$)
     Else
         pi.line2 = "Открытие исходного документа Word ..."
         Set doc = WA.Documents.Open(TemplateFilename$, , True, False)
     End If

     '  Main_PI.Log "Документ создан? " & Not (doc Is Nothing)
     doc.ActiveWindow.View.ShowFieldCodes = True    ' отображаем поля

     If MULTIROW_MODE Then    ' размножение специальных строк в шаблоне
         Dim rc&: rc = Val(options("{%rc%}"))
         If rc& = 0 Then Main_PI.Log vbTab & "Ошибка при подготовке документа Word: rc& = 0": Exit Function

         '  Dim bm As Bookmark, ra As word.Range, oFirstCellRange As word.Range
         For Each bm In doc.Bookmarks
             If bm.name Like "MultiRow*" Then
                 If bm.Range.Information(12) Then    'Закладка в таблице
                     For I = 1 To rc&
                         With bm.Range
                             Set oFirstCellRange = .Cells(1).Range
                             oFirstCellRange.Collapse 1    'wdCollapseStart
                             .Copy
                             'Вставка строки из закладки над закладкой
                             oFirstCellRange.PasteAndFormat 16    'wdFormatOriginalFormatting
                             WordReplacements .Tables(1).Rows(.Rows(1).Index).Range, "#}", "#" & I & "}"
                             WordReplacements .Tables(1).Rows(.Rows(1).Index).Range, "{%index%}", I
                         End With
                     Next
                     bm.Range.Rows(1).Delete
                 Else
                     bmText$ = bm.Range.Text
                     For I = rc& To 1 Step -1
                         With bm.Range
                             .InsertParagraphAfter
                             With .Paragraphs.First.Next
                                 .Range.InsertCrossReference ReferenceType:=2, ReferenceKind:=-1, _
                                                             ReferenceItem:=bm.name, InsertAsHyperlink:=False, _
                                                             IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
                                 .Range.Fields.Unlink
                             End With
                             WordReplacements .Paragraphs.First.Next.Range, "#}", "#" & I & "}"
                             WordReplacements .Paragraphs.First.Next.Range, "{%index%}", I
                         End With
                     Next
                     bm.Range.Delete
                 End If
                 DoEvents
             End If
         Next
     End If

     pi.line2 = "Подстановка значений в созданный по шаблону документ ..."

     Arr = options.Keys
     Replace_LF_with$ = Replace(LINEFEED_CHAR, "del", "")
     Dim FullReplace As Boolean: FullReplace = REPLACE_IN_COLON

     For I = LBound(Arr) To UBound(Arr)
         Key$ = Arr(I)
         txt$ = options(Arr(I))
         txt = Replace(txt, Chr(10), Replace_LF_with$)  ' переносы строк

         Err.Clear


         If HasLinkToObject(txt$, Key$) Then
             InsertObjectIntoDOC doc, txt$, Key$
             Err.Clear
         Else

             If FullReplace Then

                 ' новая версия замены
                 For Each myStoryRange In doc.StoryRanges
                     DoEvents
                     myStoryRange.Find.Execute Key$, False, , False, , , , , , txt$, 2
                     While Not (myStoryRange.NextStoryRange Is Nothing)
                         Set myStoryRange = myStoryRange.NextStoryRange
                         myStoryRange.Find.Execute Key$, False, , False, , , , , , txt$, 2
                     Wend
                 Next myStoryRange

                 '            For Each oSec In doc.Sections
                 '                For Each oHF In oSec.Headers
                 '                    oHF.Range.Find.Execute key$, False, , False, , , , , , txt$, 2
                 '                Next
                 '                For Each oHF In oSec.Footers
                 '                    oHF.Range.Find.Execute key$, False, , False, , , , , , txt$, 2
                 '                Next
                 '            Next

             Else
                 ' обычная быстрая замена
                 doc.Range.Find.Execute Key$, False, , False, , , , , , txt$, 2
             End If


         End If

         If Err Then
             ecount = ecount + 1
             pi.Parent.Log "ОШИБКА " & Err.Number & " при подстановке данных в поле " & Key$ & ": " & Err.Description
         End If

     Next I
     doc.ActiveWindow.View.ShowFieldCodes = False    ' скрываем поля


     pi.line2 = "Сохранение заполненного документа ..."
     pi.line3 = "Новое имя файла: " & Split(NewFilename$, "\")(UBound(Split(NewFilename$, "\")))
     Main_PI.Log vbTab & "Сохранение созданного файла: " & Replace(NewFilename$, OUTPUT_FOLDER$, "...\")
     pi.FP.Repaint

     If Val(WA.Version) > 11 And PRINT_TO_PDF Then    ' вывод в ПДФ
         doc.ExportAsFixedFormat NewFilename$, 17
     Else    ' обычное сохранение файла Word
         File_Format = GetFileFormatForNewFile(NewFilename$)
         If Len(File_Format) Then
             doc.SaveAs NewFilename$, Val(File_Format)
         Else
             doc.SaveAs NewFilename$
         End If
     End If
     If IMMEDIATE_PRINTOUT Then doc.PrintOut Copies:=PrintCopiesCount(options)
     doc.Close False
     CreateAndFill_DOC = (Err = 0 And ecount = 0)
 End Function

 Sub WordReplacements(Rng As Object, ByVal FindText As String, ByVal ReplaceText As String)
     Rng.Find.Execute FindText:=FindText, ReplaceWith:=ReplaceText, Replace:=2
 End Sub


 Function CreateAndFill_TXT(ByVal TemplateFilename$, ByVal NewFilename$, _
                            ByRef options As Dictionary, Optional ByRef pi As ProgressIndicator) As Boolean

     On Error Resume Next: Err.Clear
     pi.line3 = "Файл: " & Dir(TemplateFilename$, vbNormal)

     pi.line2 = "Чтение текстового документа ..."
     TextFile$ = ReadTXTfile(TemplateFilename$)

     pi.line2 = "Подстановка значений в текстовый файл ..."
     Arr = options.Keys
     For I = LBound(Arr) To UBound(Arr)
         Key$ = Arr(I)
         txt$ = options(Arr(I))
         TextFile$ = Replace(TextFile$, Key$, txt$, , , vbTextCompare)
     Next I

     pi.line2 = "Сохранение заполненного документа ..."
     pi.line3 = "Новое имя файла: " & Split(NewFilename$, "\")(UBound(Split(NewFilename$, "\")))
     Main_PI.Log vbTab & "Сохранение созданного файла: " & Replace(NewFilename$, OUTPUT_FOLDER$, "...\")
     pi.FP.Repaint
     SaveTXTfile NewFilename$, TextFile$
     CreateAndFill_TXT = Err = 0
 End Function

 Function TemplateType(ByVal Filename$) As String
     Select Case Extension(Filename$)
         Case "XLS", "XLSM", "XLSX", "XLSB", "CSV": TemplateType = "XLS"
         Case "XLT", "XLTM", "XLTX": TemplateType = "XLS-template"
         Case "DOC", "DOCM", "DOCX", "DOCB": TemplateType = "DOC"
         Case "DOT", "DOTM", "DOTX": TemplateType = "DOC-template"
         Case "TXT", "DAT", "XML": TemplateType = "TXT"
     End Select
 End Function

 Function TemplateTypeForListbox(ByVal Filename$) As String
     Select Case Extension(Filename$)
         Case "XLS", "XLSM", "XLSX", "XLSB", "CSV": TemplateTypeForListbox = "Excel"
         Case "XLT", "XLTM", "XLTX": TemplateTypeForListbox = "Excel"
         Case "DOC", "DOCM", "DOCX", "DOCB": TemplateTypeForListbox = "Word"
         Case "DOT", "DOTM", "DOTX": TemplateTypeForListbox = "Word"
         Case "TXT", "DAT", "XML": TemplateTypeForListbox = "Text"
         Case Else: TemplateTypeForListbox = "?"
     End Select
 End Function

 Function CheckTemplateFiles(ByRef coll As Collection) As Boolean
     On Error Resume Next

     Dim Msg$
     If coll.Count = 0 And Not SEND_MAIL_MODE Then
         Msg$ = "В папке с шаблонами документов не найдено ни одного файла.    " & _
                vbNewLine & "Убедитесь, что вы верно задали папку, содержащую шаблоны документов." & vbNewLine & vbNewLine & _
                "Путь к папке с шаблонами (задан в настройках программы):" & vbNewLine & TEMPLATES_FOLDER$
         MsgBox Msg, vbCritical, "Не найдены файлы шаблонов"
         Debug.Print "Шаблоны не найдены"
         ShowSettingsPage
         Exit Function
     End If

     If PL_(Msg) Then
         If PIBL_ Then Exit Function
         MsgBox Msg, vbCritical, ChrW(1044) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1085) & ChrW(1077) & _
                                 ChrW(1081) & ChrW(1096) & ChrW(1077) & ChrW(1077) & ChrW(32) & ChrW(1080) & ChrW(1089) & ChrW(1087) & _
                                 ChrW(1086) & ChrW(1083) & ChrW(1100) & ChrW(1079) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & _
                                 ChrW(1080) & ChrW(1077) & ChrW(32) & ChrW(1087) & ChrW(1088) & ChrW(1086) & ChrW(1075) & ChrW(1088) & _
                                 ChrW(1072) & ChrW(1084) & ChrW(1084) & ChrW(1099) & ChrW(32) & ChrW(171) & PROJECT_NAME$ & ChrW(187) & _
                                 ChrW(32) & ChrW(1085) & ChrW(1077) & ChrW(1074) & ChrW(1086) & ChrW(1079) & ChrW(1084) & ChrW(1086) & ChrW(1078) & ChrW(1085) & ChrW(1086) & ChrW(33)
         F_About.Show
         F_About.MultiPage1.value = 1
         StopMacro = True
         Exit Function
     End If

     For I = coll.Count To 1 Step -1
         Filename = coll(I)
         ttype$ = TemplateType(Filename)
         If ttype$ = "" Then
             n& = n& + 1
             Select Case n
                 Case Is < 4: Msg$ = Msg$ & Replace(Filename, TEMPLATES_FOLDER$, "") & vbNewLine
                 Case 4: Msg$ = Msg$ & "и т.д." & vbNewLine
                 Case Else
             End Select

             coll.ove I
         End If
     Next I

     If coll.Count > 200 Then
         Msg$ = "В папке с шаблонами документов обнаружено слишком много файлов (" & coll.Count & " шт.)    " & _
                vbNewLine & "Убедитесь, что вы верно задали папку, содержащую шаблоны документов." & vbNewLine & vbNewLine & _
                "Путь к папке с шаблонами (задан в настройках программы):" & vbNewLine & TEMPLATES_FOLDER$ & _
                vbNewLine & vbNewLine & "Начать заполнение документов?"
         ttl$ = "Слишком много файлов шаблонов - так и должно быть?"

         If MsgBox(Msg, vbExclamation + vbDefaultButton2 + vbOKCancel, ttl$) = vbCancel Then
             Debug.Print "Найдено дофига шаблонов: " & coll.Count
             ShowSettingsPage
             Exit Function
         End If

         If coll.Count > 500 Then
             Msg$ = "В папке с шаблонами документов обнаружено слишком много файлов (" & coll.Count & " шт.)    " & _
                    vbNewLine & "Убедитесь, что вы верно задали папку, содержащую шаблоны документов." & vbNewLine & vbNewLine & _
                    "Путь к папке с шаблонами (задан в настройках программы):" & vbNewLine & TEMPLATES_FOLDER$ & _
                    vbNewLine & vbNewLine & "Операция формирования документов отменена."
             ttl$ = "Ограничение на максимальное количество файлов шаблонов - 500 шт."

             MsgBox Msg, vbExclamation, ttl$
             Exit Function
         End If
     End If

     If Len(Msg) Then
         Msg$ = "В папке с шаблонами документов обнаружены неподдерживаемые файлы (" & n & " шт.):    " & _
                vbNewLine & vbNewLine & Msg$ & vbNewLine & vbNewLine & _
                "Обработать только подходящие файлы?" & vbNewLine & vbNewLine & _
                "Путь к папке с шаблонами (задан в настройках программы):" & vbNewLine & TEMPLATES_FOLDER$
         ttl$ = "Некоторые файлы шаблонов не будут обработаны"
         If MsgBox(Msg, vbExclamation + vbYesNo + vbDefaultButton2, ttl$) = vbYes Then
             CheckTemplateFiles = True
         End If
     Else
         CheckTemplateFiles = True
     End If
 End Function

 Function OUTPUT_MASK$()
     On Error Resume Next
     outputMask$ = "{%str%} - {%filename%}.{%ext%}"
     If Settings("TextBox_OutputMask") = "" Then
         SaveSetting PROJECT_NAME$, "Settings", "TextBox_OutputMask", outputMask$
     End If
     OUTPUT_MASK$ = Settings("TextBox_OutputMask")
 End Function

 Function TABLES_FOLDER$()
     On Error Resume Next
     folder$ = ThisWorkbook.Path & "\Таблицы\"
     If Settings("TextBox_TablesFolder") = "" Then
         SaveSetting PROJECT_NAME$, "Settings", "TextBox_TablesFolder", folder$
     End If
     If Dir(folder$, vbDirectory) = "" Then MkDir folder$

     TABLES_FOLDER$ = Settings("TextBox_TablesFolder")
 End Function

 Function OUTPUT_FOLDER$(Optional ByVal ForTextbox As Boolean = False)
     On Error Resume Next
     If USE_CURRENT_FOLDER Then
         If ActiveWorkbook Is Nothing Then Exit Function
         OUTPUT_FOLDER$ = Replace(ActiveWorkbook.FullName, ActiveWorkbook.name, "Документы\")
         If ForTextbox Then OUTPUT_FOLDER$ = "<папка, в которой сохранена таблица Excel>\Документы\"
         Err.Clear: Exit Function
     End If
     outputFolder$ = ThisWorkbook.Path & "\Документы\"
     If Settings("TextBox_OutputFolder") = "" Then
         SaveSetting PROJECT_NAME$, "Settings", "TextBox_OutputFolder", outputFolder$
     End If
     If Dir(outputFolder$, vbDirectory) = "" Then MkDir outputFolder$

     OUTPUT_FOLDER$ = Settings("TextBox_OutputFolder")
 End Function

 Function TEMPLATES_FOLDER$(Optional ByVal ForTextbox As Boolean = False)
     On Error Resume Next
     If USE_CURRENT_FOLDER Then
         If ActiveWorkbook Is Nothing Then Exit Function
         TEMPLATES_FOLDER$ = Replace(ActiveWorkbook.FullName, ActiveWorkbook.name, "Шаблоны\")
         If ForTextbox Then TEMPLATES_FOLDER$ = "<папка, в которой сохранена таблица Excel>\Шаблоны\"
         Err.Clear: Exit Function
     End If

     templatesFolder$ = ThisWorkbook.Path & "\Шаблоны\"
     If Settings("TextBox_TemplatesFolder") = "" Then
         SaveSetting PROJECT_NAME$, "Settings", "TextBox_TemplatesFolder", templatesFolder$
     End If
     If Dir(templatesFolder$, vbDirectory) = "" Then MkDir templatesFolder$

     TEMPLATES_FOLDER$ = Settings("TextBox_TemplatesFolder")
 End Function

 Function ADD_HYPERLINKS() As Boolean
     On Error Resume Next: en& = Err.Number
     ADD_HYPERLINKS = SettingsBoolean("CheckBox_AddHyperlinks")
     If en& = 0 Then Err.Clear
 End Function

 Function USE_CURRENT_FOLDER() As Boolean
     On Error Resume Next: en& = Err.Number
     USE_CURRENT_FOLDER = CBool(Settings("CheckBox_UseCurrentFolder"))
     If en& = 0 Then Err.Clear
 End Function

 Function PRINT_TO_PDF() As Boolean
     On Error Resume Next: en& = Err.Number
     PRINT_TO_PDF = CBool(Settings("CheckBox_PDF"))
     If en& = 0 Then Err.Clear
 End Function

 Function IMMEDIATE_PRINTOUT() As Boolean
     On Error Resume Next: en& = Err.Number
     IMMEDIATE_PRINTOUT = CBool(Settings("CheckBox_ImmediatePrintOut", False))
     If en& = 0 Then Err.Clear
 End Function

 Function REPLACE_IN_COLON() As Boolean
     On Error Resume Next: en& = Err.Number
     REPLACE_IN_COLON = CBool(Settings("CheckBox_ReplaceInColon"))
     If en& = 0 Then Err.Clear
 End Function

 Function ShowFolderWhenDone() As Boolean
     On Error Resume Next: en& = Err.Number
     ShowFolderWhenDone = CBool(Settings("CheckBox_ShowFolderWhenDone"))
     If en& = 0 Then Err.Clear
 End Function

 Function HEADER_ROW() As Long
     On Error Resume Next
     HEADER_ROW = Val(Settings("ComboBox_FirstRow"))
     If HEADER_ROW = 0 Then HEADER_ROW = 1
 End Function
 Function HEADER_Column() As Long
     On Error Resume Next
     Set r = FindAll(Worksheets("Форма").Cells, "Заголовки")
     HEADER_Column = Val(r.Column)
     If HEADER_Column = 0 Then HEADER_Column = 1
 End Function

 Function FullDate(ByVal D As Date) As String
     Application.Volatile True
     FullDate = Format(D, "«DD» mmmm yyyy года")
 End Function

 Function LINEFEED_CHAR() As String
     On Error Resume Next
     LINEFEED_CHAR = Settings("ComboBox_LineFeed")
     If LINEFEED_CHAR = "" Then LINEFEED_CHAR = Chr(11)
 End Function

 Function LineFeedOptions()
     On Error Resume Next
     ReDim Arr(1 To 5, 1 To 2)
     Arr(1, 1) = " ": Arr(1, 2) = "пробел"
     Arr(2, 1) = Chr(13): Arr(2, 2) = "перевод абзаца"
     Arr(3, 1) = Chr(11): Arr(3, 2) = "перевод строки"
     Arr(4, 1) = Chr(31): Arr(4, 2) = "мягкий перенос"
     Arr(5, 1) = "del": Arr(5, 2) = "<удалять переносы>"
     LineFeedOptions = Arr
 End Function

 Function USE_TEMPLATES_WITH_NAMES_LIKE_WORKSHEET_NAME() As Boolean
     On Error Resume Next: en& = Err.Number
     USE_TEMPLATES_WITH_NAMES_LIKE_WORKSHEET_NAME = CBool(Settings("CheckBox_USE_TEMPLATES_WITH_NAMES_LIKE_WORKSHEET_NAME"))
     If en& = 0 Then Err.Clear
 End Function

 Function MULTIROW_MODE() As Boolean
     On Error Resume Next: en& = Err.Number
     MULTIROW_MODE = CBool(Settings("CheckBox_MultiRow"))
     If en& = 0 Then Err.Clear
 End Function

 Function CombineXLSsheets(ByRef coll As Collection) As String
     On Error Resume Next: Err.Clear
     ' объединяет листы из всех файлов в один сводный файл
     ' возвращает путь к сводному файлу, если все прошло успешно
     NewFilename$ = Replace_symbols(Settings("TextBox_CombineXLS_filename"))
     If Not NewFilename$ Like "*.xls*" Then NewFilename$ = NewFilename$ & IIf(Val(Application.Version) >= 12, ".xlsx", ".xls")
     NewFilename$ = OUTPUT_FOLDER$ & NewFilename$

     If coll.Count = 1 Then    ' просто переименовываем файл
         Name coll(1) As NewFilename$
         Exit Function
     End If

     Dim wb As Workbook, commonWB As Workbook, sh As Worksheet
     Set commonWB = Application.Workbooks.Add(xlWBATWorksheet)
     For Each Filename In coll
         fname$ = Dir(Filename, vbNormal)
         If Len(fname$) > 0 And (fname$ Like "*?.xls*") Then
             ShName$ = Left(fname$, InStrRev(fname$, ".") - 1)
             Set wb = Workbooks.Open(Filename, False, True)
             If Not wb Is Nothing Then
                 '  For Each sh In WB.Worksheets : : Next sh
                 ' копируем только первый лист
                 Err.Clear: wb.Worksheets(1).Copy , commonWB.Worksheets(commonWB.Worksheets.Count)
                 If Err = 0 Then    ' лист успешно скопирован
                     Set sh = commonWB.Worksheets(commonWB.Worksheets.Count)
                     sh.name = Left(fname$, 31)
                 End If
                 wb.Close False
             End If
         End If
     Next
     If commonWB.Worksheets.Count > 1 Then
         Application.DisplayAlerts = False
         commonWB.Worksheets(1).Delete

         File_Format = GetFileFormatForNewFile(NewFilename$)
         If Len(File_Format) Then
             commonWB.SaveAs NewFilename$, Val(File_Format)
         Else
             commonWB.SaveAs NewFilename$
         End If
         commonWB.Close False

         If SettingsBoolean("CheckBox_CombineXLS_DeleteSourceFiles") Then
             For Each Filename In coll
                 Kill Filename
             Next
         End If

         Application.DisplayAlerts = True
         CombineXLSsheets = NewFilename$
     Else
         commonWB.Close False
     End If
 End Function
