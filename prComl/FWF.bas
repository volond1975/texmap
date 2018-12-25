Attribute VB_Name = "FWF"


 '---------------------------------------------------------------------------------------
 ' Module        : mod_CommonFunctions
 ' �����     : EducatedFool  (�����)                    ����: 26.07.2012
 ' ���������� �������� ��� Excel, Word, CorelDRAW. ������, ���������������, ��������.
 ' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
 ' ��������� ��� ������ ������: http://ExcelVBA.ru/payments
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

     Randomize ' ����� �������� �����������
     URL$ = URL$ & "?HID=" & HID & "&rnd=" & Left(Rnd(Now) * 1E+15, 10)

     DownLoadFileFromURL = URLDownloadToFile(0, URL$, LocalPath$, 0, 0) = 0
 End Function

 Function GetURLstatus(ByVal URL$, Optional ByVal timeout& = 2) As Long
     ' ������� ��������� ������� ������� � ������� URL$ (����� ��� ��������)
     ' ���������� ��� ������ ������� (�����), ���� 0, ���� ������ ���������
     ' (200 - ������ ��������, 404 - �� ������, 403 - ��� �������, � �.�.)
     On Error Resume Next: URL$ = Replace(URL$, "\", "/")
     Dim xmlhttp As New WinHttpRequest
     xmlhttp.Open "GET", URL, True
     xmlhttp.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
     xmlhttp.send
     If xmlhttp.WaitForResponse(timeout) Then
         GetURLstatus = Val(xmlhttp.Status)
     Else
         GetURLstatus = 408 ' Request Timeout (������� ����� ��������)
     End If
 End Function

 Function Extension(ByVal Filename$) As String
     On Error Resume Next
     Extension = Split(Filename$, ".")(UBound(Split(Filename$, ".")))
 End Function


 Function GetFolderPath(Optional ByVal Title As String = "�������� �����", _
                        Optional ByVal InitialPath As String = "c:\") As String
     ' ������� ������� ���������� ���� ������ ����� � ���������� Title,
     ' ������� ����� ����� � ����� InitialPath
     ' ���������� ������ ���� � ��������� �����, ��� ������ ������ � ������ ������ �� ������
     Dim PS As String: PS = Application.PathSeparator
     With Application.FileDialog(msoFileDialogFolderPicker)
         If Not Right$(InitialPath, 1) = PS Then InitialPath = InitialPath & PS
         .ButtonName = "�������": .Title = Title: .InitialFileName = InitialPath
         If .Show <> -1 Then Exit Function
         GetFolderPath = .SelectedItems(1)
         If Not Right$(GetFolderPath, 1) = PS Then GetFolderPath = GetFolderPath & PS
     End With
 End Function

 Function GetFilePath(Optional ByVal Title As String = "�������� ���� ��� ���������", _
                      Optional ByVal InitialPath As String = "c:\", _
                      Optional ByVal FilterDescription As String = "����� Excel", _
                      Optional ByVal FilterExtension As String = "*.xls*") As String
     ' ������� ������� ���������� ���� ������ ����� � ���������� Title,
     ' ������� ����� ����� � ����� InitialPath
     ' ���������� ������ ���� � ���������� �����, ��� ������ ������ � ������ ������ �� ������
     ' ��� ������� ����� ������� �������� � ���������� ���������� ������
     On Error Resume Next
     With Application.FileDialog(msoFileDialogOpen)
         .ButtonName = "�������": .Title = Title: .InitialFileName = InitialPath
         .Filters.Clear: .Filters.Add FilterDescription, FilterExtension
         If .Show <> -1 Then Exit Function
         GetFilePath = .SelectedItems(1): PS = Application.PathSeparator
     End With
 End Function

 Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal mask As String = "", _
                              Optional ByVal SearchDeep As Long = 999) As Collection
     ' �������� � �������� ��������� ���� � ����� FolderPath,
     ' ����� ����� ������� ������ Mask (����� �������� ������ ����� � ����� ������/�����������)
     ' � ������� ������ SearchDeep � ��������� (���� SearchDeep=1, �� �������� �� ���������������).
     ' ���������� ���������, ���������� ������ ���� ��������� ������
     ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)

     Set FilenamesCollection = New Collection    ' ������ ������ ���������
     Set fso = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
     GetAllFileNamesUsingFSO FolderPath, mask, fso, FilenamesCollection, SearchDeep    ' �����
     Set fso = Nothing: Application.StatusBar = False    ' ������� ������ ��������� Excel
 End Function

 Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal mask As String, ByRef fso, _
                                  ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
     ' ���������� ��� ����� � �������� � ����� FolderPath, ��������� ������ FSO
     ' ������� ����� �������������� � ��� ������, ���� SearchDeep > 1
     ' ��������� ���� ��������� ������ � ��������� FileNamesColl
     On Error Resume Next: Set curfold = fso.GetFolder(FolderPath)
     If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����

         ' ���������������� ��� ������ ��� ������ ���� � ���������������
         ' � ������� ������ ����� � ������ ��������� Excel
         ' Application.StatusBar = "����� � �����: " & FolderPath

         For Each fil In curfold.files    ' ���������� ��� ����� � ����� FolderPath
             If fil.name Like "*" & mask Then FileNamesColl.Add fil.Path
         Next
         SearchDeep = SearchDeep - 1    ' ��������� ������� ������ � ���������
         If SearchDeep Then    ' ���� ���� ������ ������
             For Each sfol In curfold.SubFolders    ' ���������� ��� �������� � ����� FolderPath
                 GetAllFileNamesUsingFSO sfol.Path, mask, fso, FileNamesColl, SearchDeep
             Next
         End If
         Set fil = Nothing: Set curfold = Nothing    ' ������� ����������
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
     Set SubFoldersCollection = New Collection    ' ������ ������ ���������
     Set fso = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
     If Right(FolderPath$, 1) <> "\" Then FolderPath$ = FolderPath$ & "\"
     On Error Resume Next: Set curfold = fso.GetFolder(FolderPath$)
     For Each folder In curfold.SubFolders    ' ���������� ��� �������� � ����� FolderPath
         If folder.Path Like FolderPath$ & mask$ Then SubFoldersCollection.Add folder.Path & "\"
     Next folder
     Set fso = Nothing
 End Function

 Function GetFilenamesCollection(Optional ByVal Title As String = "�������� ����� ��� ���������", _
                                 Optional ByVal InitialPath As String = "c:\") As FileDialogSelectedItems
     ' ������� ������� ���������� ���� ������ ���������� ������ � ���������� Title,
     ' ������� ����� ����� � ����� InitialPath
     ' ���������� ������ ����� � ��������� ������, ��� ������ ������ � ������ ������ �� ������
     With Application.FileDialog(3)    ' msoFileDialogFilePicker
         .ButtonName = "�������": .Title = Title: .InitialFileName = InitialPath
         If .Show <> -1 Then Exit Function
         Set GetFilenamesCollection = .SelectedItems
     End With
 End Function

 Function Replace_symbols(ByVal txt As String) As String
     st$ = "/\:?*|""<>"    ' � ��� ������� - ���������: ~!@#$%^=`
     For I% = 1 To Len(st$)
         txt = Replace(txt, Mid(st$, I, 1), "_")
     Next
     Replace_symbols = txt
 End Function

 Function Replace_symbols2(ByVal txt As String) As String
     st$ = "/:?*|""<>"    ' � ��� ������� - ���������: ~!@#$%^=`
     For I% = 1 To Len(st$)
         txt = Replace(txt, Mid(st$, I, 1), "_")
     Next
     Replace_symbols2 = txt
 End Function

 Sub OpenFolder(ByVal FolderPath$)
     ' ��������� ����� FolderPath$ � ���������� Windows
     On Error Resume Next
     'CreateObject("wscript.shell").Run "explorer.exe /e,/root, """ & FolderPath$ & """"
     CreateObject("wscript.shell").Run "explorer.exe /e, """ & FolderPath$ & """"
 End Sub

 Sub ShowFile(ByVal FilePath$)
     ' ��������� ���� FilePath$ � ���������� Windows
     On Error Resume Next
     CreateObject("wscript.shell").Run "explorer.exe /e,/select,""" & FilePath$ & """"
 End Sub

 Sub ShowText(ByVal txt As String, Optional ByVal Index As Long)
     ' ������ ��������� ����� �� ���������� txt � ��������� ����
     ' (���� �������� � ����� ��� ��������� ������, �������� ��� ���� text####.txt,
     ' ��� #### - �����, �������� ����� �������� index, ��� ��������� 10-�������)
     ' ����� �������� ���������� ����� �� ����������� � ��������� ��-��������� (��������, � ��������)

     On Error Resume Next: Err.Clear
     ' ��������� ��� ��� ���������� �����
     Filename$ = Environ("TEMP") & "\text" & IIf(Index, Index, Left(Rnd() * 1E+15, 10)) & ".txt"
     ' ��������� ����� � ����
     With CreateObject("scripting.filesystemobject").CreateTextFile(Filename, True)
         .Write txt: .Close
     End With
     ' ��������� ��������� ����
     CreateObject("wscript.shell").Run """" & Filename$ & """"
 End Sub

 Function ChangeFileCharset(ByVal Filename$, ByVal DestCharset$, _
                            Optional ByVal SourceCharset$) As Boolean
     ' ������� ������������� (����� ���������) ���������� �����
     ' � �������� ���������� ������� �������� ���� filename$ � ���������� �����,
     ' � �������� ��������� DestCharset$ (� ������� ����� �������� ����)
     ' ������� ���������� TRUE, ���� ������������� ������ �������
     On Error Resume Next: Err.Clear
     With CreateObject("ADODB.Stream")
         .Type = 2
         If Len(SourceCharset$) Then .Charset = SourceCharset$    ' ��������� �������� ���������
         .Open
         .LoadFromFile Filename$    ' ��������� ������ �� �����
         FileContent$ = .ReadText   ' ��������� ����� ����� � ���������� FileContent$
         .Close
         .Charset = DestCharset$    ' ��������� ����� ���������
         .Open
         .WriteText FileContent$
         .SaveToFile Filename$, 2   ' ��������� ���� ��� � ����� ���������
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
