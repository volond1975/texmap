Attribute VB_Name = "mod_Main"

 '---------------------------------------------------------------------------------------
 ' Module        : mod_Main
 ' Автор     : EducatedFool  (Игорь)                    Дата: 26.03.2012
 ' Разработка макросов для Excel, Word, CorelDRAW. Быстро, профессионально, недорого.
 ' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
 ' Реквизиты для оплаты работы: http://ExcelVBA.ru/payments
 '---------------------------------------------------------------------------------------
 Option Compare Text
 Option Private Module
 Public StopMacro As Boolean, WA As Object, Main_PI As ProgressIndicator
 Public Const HYPERLINK_START$ = "Гиперссылка - "
 Public SelectedTemplates As Collection, SelectedRowsCount&    ' для выбора шаблонов на форме
 Public CombineXLScollection As Collection

 Sub CreateProgramCommandBar()
     On Error Resume Next:

     ' настройки по-умолчанию
     If Settings("TextBox_CombineXLS_filename", "") = "" Then SaveSetting PROJECT_NAME$, "Settings", "TextBox_CombineXLS_filename", "Сводный файл.xls"

     Application.ScreenUpdating = False
     ' получаем ссылку на пользовательскую панель инструментов
     Set AddinMenu = GetCommandBar(PROJECT_NAME, True)

     ' добавление новых элементов управления на панель
     'Add_Control AddinMenu, ct_BUTTON, 271, "BackupThisFile", "Создать резервную копию программы", , True
     Add_Control AddinMenu, ct_BUTTON, 593, "CreateAllDocuments", "Сформировать документы", msoButtonIconAndCaption, True    ' 248

     '    Add_Control AddinMenu, ct_BUTTON, 1088, "SetIsAddinAsTrue", "Скрыть листы файла программы", , True
     '    Add_Control AddinMenu, ct_BUTTON, 1087, "SetIsAddinAsFalse", "Отобразить листы файла программы", , False

     If SettingsBoolean("CheckBox_ShowAdditionalMenu") Then
         Set ExtendedMenu = Add_Control(AddinMenu, ct_POPUP, 0, "", "  Дополнительно")
         Add_Control ExtendedMenu, ct_BUTTON, 385, "UpdateUDFs", "Восстановить формулы", msoButtonIconAndCaption, True    ' 202
         Add_Control ExtendedMenu, ct_BUTTON, 142, "CtrlShiftT", "Вставить ссылку на таблицу... (Ctrl + Shift + T)", msoButtonIconAndCaption, True
         Add_Control ExtendedMenu, ct_BUTTON, 218, "CtrlShiftI", "Вставить ссылку на изображение... (Ctrl + Shift + I)", msoButtonIconAndCaption, False    ' 508
     End If

     Add_Control AddinMenu, ct_BUTTON, 548, "ShowSettingsPage", "Настройки", msoButtonIconAndCaption, True
     Add_Control AddinMenu, ct_BUTTON, 487, "ShowMainForm", "О программе ...", msoButtonIconAndCaption, True

     If Len(Trim(UpdatesInfo_$)) Then
         Arr = Split(UpdatesInfo_$, "&&")
         Set subMenu = Add_Control(AddinMenu, ct_POPUP, 0, "", " Обновления ", , True)
         Add_Control subMenu, ct_BUTTON, 1759, "ManualInstallUpdate", "Установить последнюю версию", msoButtonIconAndCaption, True    ' 1759

         For I = LBound(Arr) To UBound(Arr)
             Caption$ = Split(Arr(I), "==")(0)
             descr$ = Split(Arr(I), "==", 2)(1)
             Set subMenu_Updates = Add_Control(subMenu, ct_POPUP, 4356, "", Caption$, , I = LBound(Arr))
             For Each v In Split(descr$, vbLf)
                 bf& = 0    '534
                 If Trim(v) Like "+*" Then bf& = 535: v = Split(Trim(v), , 2)(1)
                 If Trim(v) Like "-*" Then bf& = 536: v = Split(Trim(v), , 2)(1)
                 If Len(Trim(v)) Then Add_Control subMenu_Updates, ct_BUTTON, bf&, "", v, msoButtonIconAndCaption, False, 1     ' 231
             Next v
         Next I
     End If

     Add_Control AddinMenu, ct_BUTTON, IIf(Val(Application.Version) <= 11, 4356, 923), "ExitProgram", "Закрыть программу", msoButtonIcon, True
     Set ThisWorkbook.app = Application
 End Sub

 Sub ExitProgram()
     On Error Resume Next
     Msg$ = "Вы уверены, что хотите закрыть надстройку для заполнения документов?   "
     If MsgBox(Msg, vbQuestion + vbDefaultButton2 + vbOKCancel, "Завершение работы программы") = vbCancel Then Exit Sub
     DeleteProgramCommandBar
     ThisWorkbook.Close False
 End Sub

 Function GetFile_MainPicture() As String
     ' создаёт во временной папке файл, возвращает путь к созданному файлу
     On Error Resume Next:
     Dim F_TXT$, buf$, tmp_file$: Const BufLen& = 5000
     F_TXT$ = F_TXT$ & "FFD8FFE000104A46494600010101012C012C0000FFDB0043000302020302020303030304030304050805050404050A070706080C0A0C0C0B0A0B0B0D0E12100D0E110E0B0B1016101113141515150C0F171816141812141514FFDB00430103040405040509050509140D0B0D1414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414141414FFC00011080030003003012200021101031101FFC400190000030101010000000000000000000000060708040905FFC40034100001030303020305060700000000000001020304050611000712213108132223324151611415717281911617245382A1A2FFC4001801000301010000000000000000000000000405060201FFC4002811000102060104000700000000000000000102110003040521314112132361323342517191E1FFDA000C03010002110311003F00EA9E82EEFDEAB06C0756CDC579512912539CC69339B4BDD3BFB3CF23FB6A6F455EA7E212BB5F37DEE1B9615AB02E155B4C59D4575516548925452D2653E472579B8C84A7D2A00E0A483A6BDA9B0BB55B774D4AE15914286FB5E8765CE6448712B1DCF9AF152867BF7F96B8F968D360AA31CAF1CDB2AC3CB618BC155092919F260D2E63C55F810D71FF007A67"
     F_TXT$ = F_TXT$ & "EDB6E7DB5BB96CA2BF6AD49352A71714CA95C14DADA713EF2168500A4A8641C11D8823A10740957A641AA4528A30A83C8EBC5119A5AE3FE8A5613FB13A9AAAF7F567C20EE755EE216FCC5D06EB88E34F53494F13554A14A8EEA42547A38414AB1D7AA8F5200D6886C98C020EA191E2777B6E4977FA6C0B1EB8EDBED52228A85C75A88DA5C758E63D8464641C2943D671D4829EA0056883C226F45D5B833EEFB5AE99AC57DFB70455355E8EC0655212F254783A84FA4389E1DD27AE7E9932B57E6CFDBFB3DF4CF52EAB79572519D5420725CBA83E7D0C803B8493C703A745918CEAE4F0CFB3436576BE152E514BD70CE51A856657425D94E60A867E2940C207E5CFC4EA36C5799D7A9F513909029D27A5079536CFE391AC11CC3CAEA34514996857CC564FA1C085D6F5C4776537C287B8D01EFBBEDDBB7CBB7AE57529050C48C7F473540F4CA4FA0A8F40918EEAD1CEE2DFD44DA8A741A95420546E59B2E4A6031163321D7E44929529011C8A50080859383D00FA68DF75F6F20EEC6DCDC1695470235522A990E1192D39DDB707D52B0950FCBA866BD7EDC53A2EDCDD1715424BD26DD151B72753D653C6257633652952940057379B0559513D720633AAA9EB32E52963603C2BA79626CE4CB27E22DFB865DCBE263712B0CADC8B1ADFB069DFDC"
     F_TXT$ = F_TXT$ & "94B352983F5CB6D27FEF4978551B937277426D46E5B8AA574C1B4E2B5528D1A5A1086DC92EE436F36CB684278200CE403EA03AE0EA8B776EE914C91C9EA0B956ABCF67CD764A5030CAD5D17C14E2B8B585F2F4A3A8C763AC7797856A7F888AD53EF1B96AEFDBD528713EEC9EDD2DFE42506CABD4B2A09E190A27033E950F97584956DBB5C513A55C6B1BAD253D2848012FB2FB38C64F382F98ACAB996FA444B5D1C87C8254A2E4B70DA1FCD44BD47A1DE9B9F7E1AFDB11A6B9FC3CFA9EA64D654CA5B5CF694DA94A3E6AD21C42393692067A2BEBAE8CECCEE635BB160C2AE18C69F514AD70EA54F5FBD0E63478BCD1FC14323E69524FC74ACA2DA966DA56052A9F46A84A71BA3151FB05216A506D414A121398C9E473ED082B249F4924EBCCDA0AAC7B0FC4655683104D6A877953CCF613350F026A11B01C29F33AFAD85249CF7F246AAAD76C976A9029E4A89480031660DB231F51C9C9CEA272B2B155ABEE2D201CE72E7ECF9E34303DC53DA86BC4BD9A2DFDDFB9E8EDE59837E52D35CA6908E41AACC0C1250803A9520214AC64AB2411AB9749FF00125B2351DE3A1505DA05523516E7A054533E04B96D95B441494B8DAC0EA12A041E80FBA07C73A710BE046C3BBFF9A9B7146AD427DDA63D2584BEF06D2952DA7127CB90D7AC10087123B8"
     F_TXT$ = F_TXT$ & "CE4A8FC74516952E814BABCA3582C496DD693210BAABE1D2975278A94942C9EA52A4754A46388F98D26697B45BB3B26E552834482EEE1FDFAA1514D618718A7330A5385425A1495ACF14A92968A4A41C92A38CF4D1150BC33EE55CB2FEDB73DE54FB590B6D2DAA25BF1CCB91C3912417DE012951CE094B67B0C76D0DD9F2F70183454F83B043FB86556378AD5B6555A7532D4F450E216BCA030CB4AF2D208E6B23A10127DDF89EF9D25769ABCEEFCEEFDA72EDB8A98F6ED96E2E649ACF985E52F932E30D454B984A54A5A5654BC27A0477C919725B5E11B6D6852D33AA3497AEFAA0214675CF2153D448EC7CB57B24FF008A0761F21A704488C408CDC78ACB71A3B69E2869A404A103E400E8068980A3FFD9"
     For I = 1 To Len(F_TXT$) / 2
         buf$ = buf$ & Chr(Val("&H" & Mid(F_TXT$, 2 * I - 1, 2)))
         If Len(buf$) > BufLen& Then res$ = res$ & buf$: buf$ = "": DoEvents
     Next: res$ = res$ & buf$
     tmp_file$ = Environ("tmp") & "\file_MainPicture_" & PROJECT_NAME$: Kill tmp_file$
     ff& = FreeFile: Open tmp_file$ For Binary Access Write As #ff
     Put #ff, , res$
     Close #ff
     If FileLen(tmp_file$) = Len(F_TXT$) / 2 Then GetFile_MainPicture = tmp_file$ Else Debug.Print FileLen(tmp_file$), Len(F_TXT$) / 2
 End Function

 Sub CreateAllDocuments()
     On Error Resume Next
     If ActiveWorkbook Is Nothing Then
         Msg$ = "Не открыта таблица Excel, содержащая исходные данные для создаваемых документов.    " & _
                vbNewLine & "Откройте таблицу, и заново запустите формирование документов"
         MsgBox Msg, vbCritical, "Не найдены исходные данные для документов"
         Exit Sub
     End If

     If USE_CURRENT_FOLDER Then
         If ActiveWorkbook.Path = "" Then
             Msg$ = "В настройках программы указано, что папку «ШАБЛОНЫ» следует искать в той же папке," & _
                    vbNewLine & "где сохранена текущая книга Excel." & vbNewLine & vbNewLine & _
                    "Но открытый в данный момент в Excel файл «" & ActiveWorkbook.name & "» ещё не был сохранён" & vbNewLine & vbNewLine & _
                    "Измените настройки программы, или предварительно сохраните на диск таблицу Excel с исходными данными"
             MsgBox Msg, vbCritical, "Невозможно установить местонахождение папки с шаблонами"
             ShowSettingsPage
             Exit Sub
         End If
     End If

     StopMacro = False
     If Dir(TEMPLATES_FOLDER$, vbDirectory) = "" Then
         Msg$ = "Не найдена папка, содержащая шаблоны формируемых документов.    " & _
                vbNewLine & "Убедитесь, что вы верно задали папку, содержащую шаблоны документов." & vbNewLine & vbNewLine & _
                "Путь к папке с шаблонами (задан в настройках программы):" & vbNewLine & TEMPLATES_FOLDER$
         MsgBox Msg, vbCritical, "Не найдены папка с шаблонами"
         Debug.Print "Папка для шаблонов не найдена"
         ShowSettingsPage
         Exit Sub
     End If

     If USE_CURRENT_FOLDER Then
         MkDir OUTPUT_FOLDER$
         Err.Clear
         If Dir(OUTPUT_FOLDER$, vbDirectory) = "" Then
             Msg$ = "Не удалось создать папку для формируемых документов.    " & _
                    vbNewLine & "Убедитесь, что вы верно задали настройки путей к папкам." & vbNewLine & vbNewLine & _
                    "Путь к папке для документов (задан в настройках программы):" & vbNewLine & OUTPUT_FOLDER$
             MsgBox Msg, vbCritical, "Ошибка создания папки"
             ShowSettingsPage
             Exit Sub
         End If
     End If


     Dim TemplatesFilenames As Collection: Set TemplatesFilenames = FilenamesCollection(TEMPLATES_FOLDER$)

     If USE_TEMPLATES_WITH_NAMES_LIKE_WORKSHEET_NAME Then
         Dim NewColl As New Collection, ShName$
         ShName$ = ActiveSheet.name

         For Each Item In TemplatesFilenames
             Filename = "": Filename = Dir(Item)
             Filename = Left(Filename, InStrRev(Filename, ".") - 1)
             If ShName$ Like "*" & Filename & "*" Then NewColl.Add Item
         Next

         If NewColl.Count = 0 Then
             Msg = "В настройка программы включена опция «Использовать только шаблоны, имена файлов которых " & vbNewLine & _
                   "содержатся в имени обрабатываемого листа Excel»" & vbNewLine & vbNewLine & _
                   "Среди " & TemplatesFilenames.Count & " шаблонов в папке «" & TEMPLATES_FOLDER$ & "»" & vbNewLine & _
                   "не было найдено ни одного файла-шаблона, имя файла которого " & vbNewLine & _
                   "присутствовало бы в имени листа «" & ShName$ & "»"

             MsgBox Msg, vbExclamation, "Не найдены подходящие шаблоны - продолжение невозможно"

             ShowSettingsPage
             Exit Sub
         End If
         Set TemplatesFilenames = NewColl

     End If
     If Not CheckTemplateFiles(TemplatesFilenames) Then Exit Sub    ' если с файлами шаблонов что-то не так...

     BaseCol& = Fix(Val(Settings("ComboBox_BaseColumn", 2)))
     If BaseCol& <= 0 Or BaseCol& >= 256 Then BaseCol& = 2

     Dim ra As Range, ro As Range, cell As Range, newRa As Range
     Set ra = Range(Cells(HEADER_ROW + 1, BaseCol&), Cells(Rows.Count, BaseCol&).End(xlUp))
     If ra.Row < HEADER_ROW + 1 Then
         Msg$ = "На текущем листе Excel программа не нашла заполненные строки." & vbNewLine & vbNewLine & _
                "Пустая строка, или нет, - программа определяет по столбцу номер " & BaseCol& & vbNewLine & _
                "(это задано в настройках программы)" & vbNewLine & vbNewLine & _
                "В открытом файле, согласно этим настройкам, в указанном столбце все ячейки пустые," & vbNewLine & _
                "в связи с чем программа не может сформировать документы." & vbNewLine & vbNewLine & _
                "Измените настройки программы, и снова запустите формирование документов."
         MsgBox Msg, vbExclamation, "На листе не найдёны заполненные строки"
         ShowSettingsPage
         F_Settings.MultiPage_Options.value = 2
         F_Settings.Label_BaseColumn.Font.Bold = True
         F_Settings.Label_BaseColumn.ForeColor = vbRed
         F_Settings.ComboBox_BaseColumn.SetFocus
         Exit Sub
     End If

     If Settings("CheckBox_UseAllRows", False) Then
         For Each cell In ra.Cells
             If Trim(cell) <> "" Then
                 If newRa Is Nothing Then Set newRa = cell Else Set newRa = Union(newRa, cell)
             End If
         Next
         If newRa Is Nothing Then
             Msg$ = "В настройках программы включена опция" & vbNewLine & _
                    "«Формировать документы ПО ВСЕМ строкам таблицы (а не только по выделенным)»" & vbNewLine & vbNewLine & _
                    "Пустая строка, или нет, - программа определяет по столбцу номер " & BaseCol& & vbNewLine & _
                    "(это тоже задано в настройках)" & vbNewLine & vbNewLine & _
                    "В открытом файле, согласно этим настройкам, НЕ НАЙДЕНЫ ЗАПОЛНЕННЫЕ СТРОКИ," & vbNewLine & _
                    "в связи с чем программа не может сформировать документы." & vbNewLine & vbNewLine & _
                    "Измените настройки программы, и снова запустите формирование документов."
             MsgBox Msg, vbExclamation, "Нет заполненных строк, по которым будут формироваться документы."
             ShowSettingsPage
             F_Settings.MultiPage_Options.value = 2
             F_Settings.ComboBox_BaseColumn.SetFocus
             Exit Sub
         Else
             Set ra = newRa.EntireRow
         End If
     Else
         Set ra = Intersect(Selection.EntireRow, Selection.EntireRow, ra.EntireRow)
     End If

     Err.Clear
     Set ra = SpecialCells_VisibleRows(ra).EntireRow
     If ra Is Nothing Then
         Msg$ = "Не выделено ни одной ячейки в видимых строках, ниже строки заголовка таблицы." & vbNewLine & vbNewLine & _
                "В настройках программы указано:" & vbNewLine & _
                " -   строкой заголовка таблицы считается строка номер " & HEADER_ROW & "," & vbNewLine & _
                " -   чтобы строка считалась заполненной, ячейка в столбце " & ColunmNameByColumnNumber(BaseCol&) & _
                " не должна быть пустой." & vbNewLine & vbNewLine & _
                "Кроме того, программа обрабатывает только видимые строки" & vbNewLine & _
                "(строки, скрытые вручную или при помощи автофильтра, не участвуют в обработке)" & vbNewLine & vbNewLine & _
                "Выделите заполненные ВИДИМЫЕ строки, и снова запустите создание документов"
         MsgBox Msg, vbExclamation, "Ошибка - не определены строки с исходными данными"
         Exit Sub
     End If
     'If ra Is Nothing Then MsgBox "Не выделено ни одной заполненной строки!", vbExclamation, "Ошибка": End

     Set ra = Intersect(ra.EntireRow, ra.EntireRow)
     rc& = Intersect(Columns(1), ra).Cells.Count

     If SettingsBoolean("CheckBox_SelectTemplates") Then
         ' диалоговое окно выбора шаблонов для заполнения
         Set SelectedTemplates = TemplatesFilenames
         SelectedRowsCount& = rc&
         F_Templates.Show
         If SelectedTemplates.Count = 0 Then Exit Sub
         Set TemplatesFilenames = SelectedTemplates
     End If

     Set ExcelTablesToBeClosed = New Collection
     Application.ScreenUpdating = False
     Set CombineXLScollection = New Collection
     Dim pi2 As ProgressIndicator, res As Boolean, template$
     Set Main_PI = New ProgressIndicator
     Main_PI.Show "Формирование документов по шаблону", , 3

     Main_PI.StartNewAction 2, 4, "Запуск приложения Microsoft Word ..."    ' , , , rc&

     Dim NeedWord As Boolean, WordAlreadyOpen As Boolean
     For Each Filename In TemplatesFilenames
         NeedWord = NeedWord Or (TemplateType(Filename) Like "DOC*")
     Next

     If NeedWord Then
         Set WA = GetObject(, "Word.Application")
         If WA Is Nothing Then Set WA = CreateObject("Word.Application") Else WordAlreadyOpen = True
         If WA Is Nothing Then
             Msg$ = "Для формирования документов требуется приложение Microsoft Word," & vbNewLine & _
                    "которое не удалось запустить по невыясненным причинам" & vbNewLine & _
                    "(возможно, Microsoft Word не установлен, или неправильно настроен)"
             MsgBox Msg$, vbCritical, "Невозможно создать документы - проблема с использованием Microsoft Word"
             Exit Sub
         End If
         WA.Visible = False
     End If

     Main_PI.StartNewAction 4, 6, "Проверка и загрузка шаблонов ..."    ', , , rc&
     If PIBL_ Then Exit Sub


     Dim options As Dictionary, FilesCreated As Long, KeysRange As Range, HLcell As Range, Blocks As Collection
     Set KeysRange = SpecialCells_TypeConstants(ActiveSheet.Rows(HEADER_ROW))

     If KeysRange Is Nothing Then
         Main_PI.Hide
         Application.ScreenUpdating = True
         MsgBox "Не найдено ни одной заполненной ячейки в строке заголовка таблицы!" & vbNewLine & _
                "(строка номер " & HEADER_ROW & ")" & vbNewLine & vbNewLine & _
                "Проверьте настройки программы, и снова запустите создание документов", _
                vbExclamation, "Ошибка - отсутствует заголовок таблицы"
         ShowSettingsPage
         F_Settings.MultiPage_Options.value = 2
         F_Settings.Label_HeaderRow.Font.Bold = True: F_Settings.Label_HeaderRow.ForeColor = vbRed
         Exit Sub
     End If

     If MULTIROW_MODE Then    ' из всех выделенных строк - один файл на каждый шаблон
         Set Blocks = CollectionOfRowsBlocks(ra)
         If Blocks.Count = 0 Then GoTo EndFor
         rc& = Blocks.Count
     End If

     Main_PI.StartNewAction 6, , "Формирование документов по шаблонам ...", , , rc&
     Application.ScreenUpdating = False


     HL_text$ = Settings("TextBox_HyperlinkText"): If HL_text$ = "" Then HL_text$ = "ссылка"

     Dim ls As New Letters, Lett As Letter

     If MULTIROW_MODE Then    ' из всех выделенных строк - один файл на каждый шаблон

         For Each ra In Blocks

             Main_PI.SubAction , "Обрабатываются все выделенные строки одновременно"
             Set options = ReadMultirowOptions(ra.EntireRow)
             Set Lett = ls.CreateNewLetter: Lett.Render options

             If TemplatesFilenames.Count Then
                 Main_PI.Log "": Main_PI.Log "Создание документов для диапазона строк " & ra.address & " таблицы Excel"
             Else
                 If SEND_MAIL_MODE Then Main_PI.Log vbTab & "Создано письмо для: " & Lett.email
             End If

             Set pi2 = Main_PI.AddChildIndicator("Заполнение шаблонов для строк " & ra.address & " таблицы Excel")
             pi2.StartNewAction , , , , , TemplatesFilenames.Count + sss

             For I = 1 To TemplatesFilenames.Count

                 template$ = TemplatesFilenames(I): Err.Clear
                 pi2.SubAction "Формируется документ $index из $count"
                 Main_PI.Log vbTab & String(40, "=")

                 Main_PI.Log vbTab & "Создание документа " & I & " из " & TemplatesFilenames.Count
                 Main_PI.Log vbTab & "Шаблон: " & Replace(template$, TEMPLATES_FOLDER$, "...\")

                 res = CreateAndFillDocument(template$, options, pi2)


                 Main_PI.Log vbTab & "Результат: " & IIf(res, "УСПЕШНО", "ОШИБКА")
                 If Err <> 0 Then Main_PI.Log "Ошибка " & Err.Number & vbTab & Err.Description
                 If SEND_MAIL_MODE Then Main_PI.Log vbTab & "Создано письмо для: " & Lett.email

                 NewFilename$ = "": NewFilename$ = CreatePathForFile(template$, options)
                 If SettingsBoolean("CheckBox_Mail_AttachCreatedFiles") Then Lett.AddAttachment NewFilename$, RenderString(Settings("TextBox_AttachCreatedFilesMask"), options)

                 If res And ADD_HYPERLINKS Then    ' проставляем гиперссылки
                     Set HLcell = Nothing
                     Set HLcell = Intersect(Get_HLink_Column(template$, KeysRange), ro)
                     HLcell.Hyperlinks.Add HLcell, NewFilename$, , "Открыть созданный файл" & vbLf & Dir(NewFilename$), HL_text$
                 End If
                 Err.Clear

                 FilesCreated = FilesCreated - res
                 If PL_(Msg) Then
                     Msg$ = "Создание документов прервано - ограничение на количество " & _
                            "формируемых документов в тестовой версии программы"
                     pi2.Hide: Main_PI.Log Msg: GoTo EndFor
                 End If
                 If StopMacro Then
                     Msg$ = vbNewLine & "Создание документов прервано по команде пользователя"
                     pi2.Hide: Main_PI.Log Msg: GoTo EndFor
                 End If
             Next I
             pi2.Hide
             Main_PI.Log vbTab & String(40, "=")
         Next ra


     Else    ' для каждой строки - отдельный набор файлов (обычный режим)

         For Each ro In ra.Rows
             Main_PI.SubAction , "Обрабатывается строка $index из $count", "$time"
             Set options = ReadOptions(ro)
             Set Lett = ls.CreateNewLetter: Lett.Render options

             If TemplatesFilenames.Count Then
                 Main_PI.Log "": Main_PI.Log "Создание документов для строки " & ro.Row & " таблицы Excel"
             Else
                 If SEND_MAIL_MODE Then Main_PI.Log vbTab & "Создано письмо для: " & Lett.email
             End If

             Set pi2 = Main_PI.AddChildIndicator("Заполнение шаблонов для строки " & ro.Row & " таблицы Excel")
             pi2.StartNewAction , , , , , TemplatesFilenames.Count + sss


             For I = 1 To TemplatesFilenames.Count

                 template$ = TemplatesFilenames(I): Err.Clear
                 pi2.SubAction "Формируется документ $index из $count"
                 Main_PI.Log vbTab & String(40, "=")

                 Main_PI.Log vbTab & "Создание документа " & I & " из " & TemplatesFilenames.Count
                 Main_PI.Log vbTab & "Шаблон: " & Replace(template$, TEMPLATES_FOLDER$, "...\")

                 res = CreateAndFillDocument(template$, options, pi2)


                 Main_PI.Log vbTab & "Результат: " & IIf(res, "УСПЕШНО", "ОШИБКА")
                 If Err <> 0 Then Main_PI.Log "Ошибка " & Err.Number & vbTab & Err.Description
                 If SEND_MAIL_MODE Then Main_PI.Log vbTab & "Создано письмо для: " & Lett.email

                 NewFilename$ = "": NewFilename$ = CreatePathForFile(template$, options)
                 If SettingsBoolean("CheckBox_Mail_AttachCreatedFiles") Then Lett.AddAttachment NewFilename$, RenderString(Settings("TextBox_AttachCreatedFilesMask"), options)

                 If res And ADD_HYPERLINKS Then    ' проставляем гиперссылки
                     Set HLcell = Nothing
                     Set HLcell = Intersect(Get_HLink_Column(template$, KeysRange), ro)
                     HLcell.Hyperlinks.Add HLcell, NewFilename$, , "Открыть созданный файл" & vbLf & Dir(NewFilename$), HL_text$
                 End If
                 Err.Clear

                 FilesCreated = FilesCreated - res
                 If PL_(Msg) Then
                     Msg$ = "Создание документов прервано - ограничение на количество " & _
                            "формируемых документов в тестовой версии программы"
                     pi2.Hide: Main_PI.Log Msg: GoTo EndFor
                 End If
                 If StopMacro Then
                     Msg$ = vbNewLine & "Создание документов прервано по команде пользователя"
                     pi2.Hide: Main_PI.Log Msg: GoTo EndFor
                 End If
             Next I
             pi2.Hide
             If TemplatesFilenames.Count Then Main_PI.Log vbTab & String(40, "=")
         Next ro
     End If

     If SettingsBoolean("CheckBox_CombineXLSsheets") Then
         If Not PRINT_TO_PDF Then
             If CombineXLScollection.Count Then
                 Main_PI.StartNewAction 98, 100, "ФОбъединение созданных файлов Excel в один ..."
                 commonFilename$ = CombineXLSsheets(CombineXLScollection)
                 If Len(commonFilename$) Then
                     Main_PI.Log "Листы из файлов Excel объединены в файл" & vbNewLine & commonFilename$
                 End If
             End If
         End If
     End If

EndFor:
     Application.DisplayAlerts = False
     For Each File In ExcelTablesToBeClosed
         Workbooks(CStr(File)).Close False
     Next
     Application.DisplayAlerts = True
     Application.ScreenUpdating = True

     If SEND_MAIL_MODE Then
         res = ls.SendAll
         Main_PI.Log vbTab: Main_PI.Log vbTab & "Отправка писем: " & IIf(res, "УСПЕШНО", "ОШИБКА")
     End If

     If Not WordAlreadyOpen Then
         WA.Quit    ' закрываем Word, если мы его запускали
     Else
         WA.Visible = True
     End If
     Set WA = Nothing

     AppActivate Application.name
     Main_PI.StartNewAction 100, 100, "Формирование документов завершено", _
                            "Создано файлов: " & FilesCreated & " из " & TemplatesFilenames.Count * rc&, _
                            "Обработано " & IIf(SettingsBoolean("CheckBox_Multirow_GroupRows"), "блоков ", "") & "строк: " & rc& & _
                            ", использовано файлов шаблонов: " & TemplatesFilenames.Count
     Main_PI.FP.SpinButton_log.Visible = False
     Main_PI.CancelButton.Caption = "ЗАКРЫТЬ"

     If TemplatesFilenames.Count = 0 And SEND_MAIL_MODE = True Then
         Main_PI.line1 = "Создание писем завершено"
         Main_PI.line2 = "Сформировано писем: " & ls.Items.Count
     Else
         Main_PI.Log "": Main_PI.Log "Результаты работы программы:"
         Main_PI.Log vbTab & "Использовано шаблонов: " & TemplatesFilenames.Count
         Main_PI.Log vbTab & "Обработано записей в таблице Excel: " & rc&
         Main_PI.Log vbTab & "Создано файлов: " & FilesCreated & IIf(MULTIROW_MODE, " (включена опция «MULTIROW»)", "")
         ErrFilesCount& = TemplatesFilenames.Count * rc& - FilesCreated
         If ErrFilesCount& Then Main_PI.Log vbTab & "Не удалось создать файлов: " & ErrFilesCount&
     End If


     StopMacro = True
     Main_PI.Log vbNewLine    '& vbNewLine

     'pi.Hide
     Set ThisWorkbook.app = Application

     Main_PI.CancelButton.Width = 0
     If FilesCreated > 0 Then
         If ShowFolderWhenDone Then
             OpenFolder OUTPUT_FOLDER$
         Else
             Main_PI.AddButton "Открыть папку с созданными файлами", "OpenDestinationFolder"
         End If
     End If

     If SettingsBoolean("CheckBox_CloseProgressBar") Then Main_PI.Hide

     'Debug.Print TemplatesInfo(TemplatesFilenames)
     info$ = "Rows=" & rc& & ", " & "Templates=" & TemplatesInfo(TemplatesFilenames) & ", " & _
             "Files=" & FilesCreated & "/" & TemplatesFilenames.Count * rc& & vbNewLine & ", Counters: " & CountersCurrentValues
     'ND "run macro", info$
 End Sub


 Function Get_HLink_Column(ByVal Filename$, ByRef KeysRange As Range) As Range
     On Error Resume Next: Err.Clear
     celltext$ = HYPERLINK_START$ & Dir(Filename$, vbNormal)
     If celltext$ Like "*.*" Then celltext$ = Left(celltext$, InStrRev(celltext$, ".") - 1)

     Set Get_HLink_Column = KeysRange.Find(celltext$, , xlValues, xlPart).EntireColumn
 End Function


 Sub OpenDestinationFolder()
     On Error Resume Next: OpenFolder OUTPUT_FOLDER$
 End Sub

 Function CreateAndFillDocument(ByVal TemplateFilename$, ByRef options As Dictionary, _
                                Optional ByRef pi As ProgressIndicator) As Boolean
     On Error Resume Next: options("{%pcc%}") = ""
     NewFilename$ = CreatePathForFile(TemplateFilename$, options)

     ttype$ = TemplateType(TemplateFilename$)
     Select Case True
         Case ttype$ Like "XLS*"
             CreateAndFillDocument = CreateAndFill_XLS(TemplateFilename$, NewFilename$, options, pi)
             If CreateAndFillDocument Then CombineXLScollection.Add NewFilename$
         Case ttype$ Like "DOC*"
             CreateAndFillDocument = CreateAndFill_DOC(TemplateFilename$, NewFilename$, options, pi)
         Case ttype$ Like "TXT"
             CreateAndFillDocument = CreateAndFill_TXT(TemplateFilename$, NewFilename$, options, pi)
         Case Else:
             CreateAndFillDocument = False
     End Select
 End Function




 ' =================== тестирование =======================================


 Private Sub TestFilenamesMask()
     On Error Resume Next

     Dim TemplatesFilenames As Collection: Set TemplatesFilenames = FilenamesCollection(TEMPLATES_FOLDER$)
     BaseCol& = 2

     Worksheets(2).Activate
     Dim ra As Range, ro As Range, cell As Range, newRa As Range
     Set ra = Range(Cells(2, BaseCol&), Cells(3, BaseCol&)).EntireRow
     Set ra = Intersect(ra.EntireRow, ra.EntireRow)
     rc& = Intersect(Columns(1), ra).Cells.Count


     Dim options As Dictionary, FilesCreated As Long, KeysRange As Range, HLcell As Range
     Set KeysRange = SpecialCells_TypeConstants(ActiveSheet.Rows(HEADER_ROW))

     Worksheets(3).Activate
     For Each ro In ra.Rows
         Range("a555").End(xlUp).Offset(1).value = "<b>Строка №" & ro.Row & "  (" & ro.Cells(3) & ")</b>"

         Set options = ReadOptions(ro)
         For I = 1 To TemplatesFilenames.Count
             template$ = TemplatesFilenames(I)
             res1 = Replace(template$, TEMPLATES_FOLDER$, "...\Шаблоны\" & "<b>") & "</b>"

             ' res = CreateAndFillDocument(template$, options, pi2)
             NewFilename$ = "": NewFilename$ = CreatePathForFile(template$, options)
             res2 = Replace(NewFilename$, OUTPUT_FOLDER$, "...\Документы\" & "<b>") & "</b>"

             Range("a555").End(xlUp).Offset(1).Resize(, 3).value = _
             Array(OUTPUT_MASK$, res1, res2)
         Next I
     Next ro
 End Sub

End Sub
