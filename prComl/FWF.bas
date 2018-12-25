Attribute VB_Name = "FWF"


 '---------------------------------------------------------------------------------------
 ' Module        : mod_CommonFunctions
 ' Автор     : EducatedFool  (Игорь)                    Дата: 26.07.2012
 ' Разработка макросов для Excel, Word, CorelDRAW. Быстро, профессионально, недорого.
 ' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
 ' Реквизиты для оплаты работы: http://ExcelVBA.ru/payments
 '---------------------------------------------------------------------------------------

 Option Private Module
 Const FWF_VERSION = 2

 #If Win64 Then
     #If VBA7 Then    ' Windows x64, Office 2010
         Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                 (ByVal pCaller As LongLong, ByVal szURL As String, ByVal szFileName As String, _
                  ByVal dwReserved As LongLong, ByVal lpfnCB As LongLong) As LongLong
     #Else    ' Windows x64,Office 2003-2007
         Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                                            (ByVal pCaller As LongLong, ByVal szURL As String, ByVal szFileName As String, _
                                             ByVal dwReserved As LongLong, ByVal lpfnCB As LongLong) As LongLong
     #End If
 #Else
     #If VBA7 Then    ' Windows x86, Office 2010
         Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                 (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
                  ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
     #Else    ' Windows x86, Office 2003-2007
         Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                                            (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
                                             ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
     #End If
 #End If

 Function DownLoadFileFromURL(ByVal URL$, ByVal LocalPath$) As Boolean
     On Error Resume Next: Kill LocalPath$

     shortFilename$ = Mid(LocalPath$, InStrRev(LocalPath$, "\") + 1)
     If shortFilename$ <> Replace_symbols(shortFilename$) Then
         Debug.Print "Wrong symbols in filename: " & shortFilename$
         Exit Function
     End If

     Randomize ' чтобы избежать кеширования
     URL$ = URL$ & "?HID=" & HID & "&rnd=" & Left(Rnd(Now) * 1E+15, 10)

     DownLoadFileFromURL = URLDownloadToFile(0, URL$, LocalPath$, 0, 0) = 0
 End Function

 Function GetURLstatus(ByVal URL$, Optional ByVal timeout& = 2) As Long
     ' функция проверяет наличие доступа к ресурсу URL$ (файлу или каталогу)
     ' возвращает код ответа сервера (число), либо 0, если ссылка ошибочная
     ' (200 - ресурс доступен, 404 - не найден, 403 - нет доступа, и т.д.)
     On Error Resume Next: URL$ = Replace(URL$, "\", "/")
     Dim xmlhttp As New WinHttpRequest
     xmlhttp.Open "GET", URL, True
     xmlhttp.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
     xmlhttp.send
     If xmlhttp.WaitForResponse(timeout) Then
         GetURLstatus = Val(xmlhttp.Status)
     Else
         GetURLstatus = 408 ' Request Timeout (истекло время ожидания)
     End If
 End Function

 Function Extension(ByVal Filename$) As String
     On Error Resume Next
     Extension = Split(Filename$, ".")(UBound(Split(Filename$, ".")))
 End Function


 Function GetFolderPath(Optional ByVal Title As String = "Выберите папку", _
                        Optional ByVal InitialPath As String = "c:\") As String
     ' функция выводит диалоговое окно выбора папки с заголовком Title,
     ' начиная обзор диска с папки InitialPath
     ' возвращает полный путь к выбранной папке, или пустую строку в случае отказа от выбора
     Dim PS As String: PS = Application.PathSeparator
     With Application.FileDialog(msoFileDialogFolderPicker)
         If Not Right$(InitialPath, 1) = PS Then InitialPath = InitialPath & PS
         .ButtonName = "Выбрать": .Title = Title: .InitialFileName = InitialPath
         If .Show <> -1 Then Exit Function
         GetFolderPath = .SelectedItems(1)
         If Not Right$(GetFolderPath, 1) = PS Then GetFolderPath = GetFolderPath & PS
     End With
 End Function

 Function GetFilePath(Optional ByVal Title As String = "Выберите файл для обработки", _
                      Optional ByVal InitialPath As String = "c:\", _
                      Optional ByVal FilterDescription As String = "Книги Excel", _
                      Optional ByVal FilterExtension As String = "*.xls*") As String
     ' функция выводит диалоговое окно выбора файла с заголовком Title,
     ' начиная обзор диска с папки InitialPath
     ' возвращает полный путь к выбранному файлу, или пустую строку в случае отказа от выбора
     ' для фильтра можно указать описание и расширение выбираемых файлов
     On Error Resume Next
     With Application.FileDialog(msoFileDialogOpen)
         .ButtonName = "Выбрать": .Title = Title: .InitialFileName = InitialPath
         .Filters.Clear: .Filters.Add FilterDescription, FilterExtension
         If .Show <> -1 Then Exit Function
         GetFilePath = .SelectedItems(1): PS = Application.PathSeparator
     End With
 End Function

 Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal mask As String = "", _
                              Optional ByVal SearchDeep As Long = 999) As Collection
     ' Получает в качестве параметра путь к папке FolderPath,
     ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
     ' и глубину поиска SearchDeep в подпапках (если SearchDeep=1, то подпапки не просматриваются).
     ' Возвращает коллекцию, содержащую полные пути найденных файлов
     ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)

     Set FilenamesCollection = New Collection    ' создаём пустую коллекцию
     Set fso = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
     GetAllFileNamesUsingFSO FolderPath, mask, fso, FilenamesCollection, SearchDeep    ' поиск
     Set fso = Nothing: Application.StatusBar = False    ' очистка строки состояния Excel
 End Function

 Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal mask As String, ByRef fso, _
                                  ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
     ' перебирает все файлы и подпапки в папке FolderPath, используя объект FSO
     ' перебор папок осуществляется в том случае, если SearchDeep > 1
     ' добавляет пути найденных файлов в коллекцию FileNamesColl
     On Error Resume Next: Set curfold = fso.GetFolder(FolderPath)
     If Not curfold Is Nothing Then    ' если удалось получить доступ к папке

         ' раскомментируйте эту строку для вывода пути к просматриваемой
         ' в текущий момент папке в строку состояния Excel
         ' Application.StatusBar = "Поиск в папке: " & FolderPath

         For Each fil In curfold.files    ' перебираем все файлы в папке FolderPath
             If fil.name Like "*" & mask Then FileNamesColl.Add fil.Path
         Next
         SearchDeep = SearchDeep - 1    ' уменьшаем глубину поиска в подпапках
         If SearchDeep Then    ' если надо искать глубже
             For Each sfol In curfold.SubFolders    ' перебираем все подпапки в папке FolderPath
                 GetAllFileNamesUsingFSO sfol.Path, mask, fso, FileNamesColl, SearchDeep
             Next
         End If
         Set fil = Nothing: Set curfold = Nothing    ' очищаем переменные
     End If
 End Function

 Function ReadTXTfile(ByVal Filename As String) As String
     Set fso = CreateObject("scripting.filesystemobject")
     Set ts = fso.OpenTextFile(Filename, 1, True): ReadTXTfile = ts.ReadAll: ts.Close
     Set ts = Nothing: Set fso = Nothing
 End Function

 Function SaveTXTfile(ByVal Filename As String, ByVal txt As String) As Boolean
     On Error Resume Next: Err.Clear
     Set fso = CreateObject("scripting.filesystemobject")
     Set ts = fso.CreateTextFile(Filename, True)
     ts.Write txt: ts.Close
     SaveTXTfile = Err = 0
     Set ts = Nothing: Set fso = Nothing
 End Function

 Function AddIntoTXTfile(ByVal Filename As String, ByVal txt As String) As Boolean
     On Error Resume Next: Err.Clear
     Set fso = CreateObject("scripting.filesystemobject")
     Set ts = fso.OpenTextFile(Filename, 8, True): ts.Write txt: ts.Close
     Set ts = Nothing: Set fso = Nothing
     AddIntoTXTfile = Err = 0
 End Function

 Function SubFoldersCollection(ByVal FolderPath$, Optional ByVal mask$ = "*") As Collection
     Set SubFoldersCollection = New Collection    ' создаём пустую коллекцию
     Set fso = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
     If Right(FolderPath$, 1) <> "\" Then FolderPath$ = FolderPath$ & "\"
     On Error Resume Next: Set curfold = fso.GetFolder(FolderPath$)
     For Each folder In curfold.SubFolders    ' перебираем все подпапки в папке FolderPath
         If folder.Path Like FolderPath$ & mask$ Then SubFoldersCollection.Add folder.Path & "\"
     Next folder
     Set fso = Nothing
 End Function

 Function GetFilenamesCollection(Optional ByVal Title As String = "Выберите файлы для обработки", _
                                 Optional ByVal InitialPath As String = "c:\") As FileDialogSelectedItems
     ' функция выводит диалоговое окно выбора нескольких файлов с заголовком Title,
     ' начиная обзор диска с папки InitialPath
     ' возвращает массив путей к выбранным файлам, или пустую строку в случае отказа от выбора
     With Application.FileDialog(3)    ' msoFileDialogFilePicker
         .ButtonName = "Выбрать": .Title = Title: .InitialFileName = InitialPath
         If .Show <> -1 Then Exit Function
         Set GetFilenamesCollection = .SelectedItems
     End With
 End Function

 Function Replace_symbols(ByVal txt As String) As String
     st$ = "/\:?*|""<>"    ' а эти символы - разрешены: ~!@#$%^=`
     For I% = 1 To Len(st$)
         txt = Replace(txt, Mid(st$, I, 1), "_")
     Next
     Replace_symbols = txt
 End Function

 Function Replace_symbols2(ByVal txt As String) As String
     st$ = "/:?*|""<>"    ' а эти символы - разрешены: ~!@#$%^=`
     For I% = 1 To Len(st$)
         txt = Replace(txt, Mid(st$, I, 1), "_")
     Next
     Replace_symbols2 = txt
 End Function

 Sub OpenFolder(ByVal FolderPath$)
     ' открывает папку FolderPath$ в Проводнике Windows
     On Error Resume Next
     'CreateObject("wscript.shell").Run "explorer.exe /e,/root, """ & FolderPath$ & """"
     CreateObject("wscript.shell").Run "explorer.exe /e, """ & FolderPath$ & """"
 End Sub

 Sub ShowFile(ByVal FilePath$)
     ' открывает файл FilePath$ в Проводнике Windows
     On Error Resume Next
     CreateObject("wscript.shell").Run "explorer.exe /e,/select,""" & FilePath$ & """"
 End Sub

 Sub ShowText(ByVal txt As String, Optional ByVal Index As Long)
     ' макрос сохраняет текст из переменной txt в текстовый файл
     ' (файл создаётся в папке для временных файлов, получает имя типа text####.txt,
     ' где #### - число, заданное через параметр index, или случайное 10-значное)
     ' После создания текстового файла он открывается в программе по-умолчанию (например, в Блокноте)

     On Error Resume Next: Err.Clear
     ' формируем имя для временного файла
     Filename$ = Environ("TEMP") & "\text" & IIf(Index, Index, Left(Rnd() * 1E+15, 10)) & ".txt"
     ' сохраняем текст в файл
     With CreateObject("scripting.filesystemobject").CreateTextFile(Filename, True)
         .Write txt: .Close
     End With
     ' открываем созданный файл
     CreateObject("wscript.shell").Run """" & Filename$ & """"
 End Sub

 Function ChangeFileCharset(ByVal Filename$, ByVal DestCharset$, _
                            Optional ByVal SourceCharset$) As Boolean
     ' функция перекодировки (смены кодировки) текстового файла
     ' В качестве параметров функция получает путь filename$ к текстовому файлу,
     ' и название кодировки DestCharset$ (в которую будет переведён файл)
     ' Функция возвращает TRUE, если перекодировка прошла успешно
     On Error Resume Next: Err.Clear
     With CreateObject("ADODB.Stream")
         .Type = 2
         If Len(SourceCharset$) Then .Charset = SourceCharset$    ' указываем исходную кодировку
         .Open
         .LoadFromFile Filename$    ' загружаем данные из файла
         FileContent$ = .ReadText   ' считываем текст файла в переменную FileContent$
         .Close
         .Charset = DestCharset$    ' назначаем новую кодировку
         .Open
         .WriteText FileContent$
         .SaveToFile Filename$, 2   ' сохраняем файл уже в новой кодировке
         .Close
     End With
     ChangeFileCharset = Err = 0
 End Function

 Function temp_folder$()
     On Error Resume Next
     temp_folder$ = Environ("TEMP") & "\ExcelTemporaryFiles\"
     If Dir(temp_folder$, vbDirectory) = "" Then MkDir temp_folder$
 End Function

 Function temp_filename$()
     On Error Resume Next: Dim iter&
get_rnd:      iter& = iter& + 1: txt$ = Left(Rnd(Now) * 1E+15, 10)
     temp_filename$ = temp_folder$ & "temp_file_" & Format(Now, "YYYY-MM-DD--HH-NN-SS") & "__" & txt$
     If Dir(temp_filename$, vbNormal) <> "" Then If iter& < 5 Then GoTo get_rnd
 End Function
