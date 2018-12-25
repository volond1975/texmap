Attribute VB_Name = "modFSO"
'http://www.script-coding.com/WSH/FileSystemObject.html
'http://citforum.ck.ua/programming/digest/fsovb6.shtml
'http://www.4guysfromrolla.com/webtech/faq/FileSystemObject/faq1.shtml

'http://excelvba.ru/code/FilenamesCollection

Type PathSplitString
sDrive As String
sDir As String
sFileName As String
oGetFile As Object
sGetBaseName As String
sExtension As String
sGetParentFolderName As String
bFileExists As Boolean
bFolderExists As Boolean
bDriveExists As Boolean
End Type



Function txtReadAll(strFileName)
'Чтение текстового файла в переменную
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim objTextStream


Const fsoForReading = 1

If objFSO.fileExists(strFileName) Then
    'The file exists, so open it and output its contents
    Set objTextStream = objFSO.OpenTextFile(strFileName, fsoForReading)
    txtReadAll = objTextStream.ReadAll
    objTextStream.Close
    Set objTextStream = Nothing
Else
    'The file did not exist
    Debug.Print strFileName & " was not found."
End If

'Clean up
Set objFSO = Nothing

End Function
Function txtReadLines(strFileName, Optional Lines)
'Считывает из текстового файла в переменную указаное количество строк с по
'По умолчанию если не задано то считывает весь файл
'Строки задаются так "2:10"
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.GetFile("C:\boot.ini")
Set TextStream = file.OpenAsTextStream(1)
str = vbNullString
While Not TextStream.AtEndOfStream
    str = str & TextStream.ReadLine() & vbCrLf
Wend
txtReadLines = str
TextStream.Close
End Function
Sub testIsFileExist()
Dim ps As PathSplitString
'"C:\Program Files\1cv82\common\1cestart.exe"
ps = PathSplit("C:\Program Files\1cv82\common\1cestart.exe")
MsgBox ps.bFileExists
End Sub

Function IsFileExist(FullPath As String)
Dim ps As PathSplitString
ps = PathSplit(FullPath)
IsFileExist = ps.bFileExists
End Function
Function PathSplit(FullPath As String) As PathSplitString
Dim ps As PathSplitString
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
With fso
ps.bDriveExists = .DriveExists(FullPath)
ps.bFileExists = .fileExists(FullPath)
ps.bFolderExists = .FolderExists(FullPath)
ps.sExtension = .GetExtensionName(FullPath)
ps.sGetBaseName = .GetBaseName(FullPath)
ps.sDrive = .GetDriveName(FullPath)
ps.sGetParentFolderName = .GetParentFolderName(FullPath)
ps.sFileName = .GetFileName(FullPath)
End With
PathSplit = ps
End Function
Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                             Optional ByVal SearchDeep As Long = 999) As Collection
    ' Получает в качестве параметра путь к папке FolderPath,
    ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
    ' и глубину поиска SearchDeep в подпапках (если SearchDeep=1, то подпапки не просматриваются).
    ' Возвращает коллекцию, содержащую полные пути найденных файлов
    ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)

    Set FilenamesCollection = New Collection    ' создаём пустую коллекцию
    Set fso = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, fso, FilenamesCollection, SearchDeep ' поиск
    Set fso = Nothing: Application.StatusBar = False    ' очистка строки состояния Excel
End Function
 
Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef fso, _
                                 ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    ' перебирает все файлы и подпапки в папке FolderPath, используя объект FSO
    ' перебор папок осуществляется в том случае, если SearchDeep > 1
    ' добавляет пути найденных файлов в коллекцию FileNamesColl
    On Error Resume Next: Set curfold = fso.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' если удалось получить доступ к папке

        ' раскомментируйте эту строку для вывода пути к просматриваемой
        ' в текущий момент папке в строку состояния Excel
        ' Application.StatusBar = "Поиск в папке: " & FolderPath

        For Each fil In curfold.Files    ' перебираем все файлы в папке FolderPath
            If fil.name Like "*" & Mask Then FileNamesColl.Add fil.Path
        Next
        SearchDeep = SearchDeep - 1    ' уменьшаем глубину поиска в подпапках
        If SearchDeep Then    ' если надо искать глубже
            For Each sfol In curfold.SubFolders    ' перебираем все подпапки в папке FolderPath
                GetAllFileNamesUsingFSO sfol.Path, Mask, fso, FileNamesColl, SearchDeep
            Next
        End If
        Set fil = Nothing: Set curfold = Nothing    ' очищаем переменные
    End If
End Function

Sub Пример_FilenamesCollection()
    On Error Resume Next
    Dim folder$, coll As Collection
 
    folder$ = ThisWorkbook.Path & "\Платежи\"
    If Dir(folder$, vbDirectory) = "" Then
        MsgBox "Не найдена папка «" & folder$ & "»", vbCritical, "Нет папки ПЛАТЕЖИ"
        Exit Sub        ' выход, если папка не найдена
    End If
 
    Set coll = FilenamesCollection(folder$, "*.xls")        ' получаем список файлов XLS из папки
    If coll.Count = 0 Then
        MsgBox "В папке «" & Split(folder$, "\")(UBound(Split(folder$, "\")) - 1) & "» нет ни одного подходящего файла!", _
               vbCritical, "Файлы для обработки не найдены"
        Exit Sub        ' выход, если нет файлов
    End If
 
    ' перебираем все найденные файлы
    For Each file In coll
        Debug.Print file        ' выводим имя файла в окно Immediate
    Next
End Sub


 Function CreateFolder(FullPath As String)
 Dim ps As PathSplitString
' myPath & "\" & z.Value
With CreateObject("Scripting.FileSystemObject")
'Set lis = СписокЗначенийСтолбцаУмнойТаблицы("Лісництво", "Лісництво", "Лісництво")
'For Each z In lis.Cells
If Not .FolderExists(FullPath) Then
Set folder = .CreateFolder(FullPath)
Else
ps = PathSplit(FullPath)
Set folder = .GetFolder(FullPath)
End If
End With
Set CreateFolder = folder
 End Function
Function DataCreatedFile(FilePath As String, Параметр)
        Dim objFSO As Scripting.FileSystemObject
        Dim fsoFile, DateCreate
'        Dim FilePath As String
            Set objFSO = New Scripting.FileSystemObject
          
            Set fsoFile = objFSO.GetFile(FilePath)
            Select Case Параметр
            Case "Дата создания"
            DataCreatedFile = (fsoFile.DateCreated)
            Case "Последний доступ"
            DataCreatedFile = (fsoFile.DateLastAccessed)
            Case "Последнее изменение"
             DataCreatedFile = (fsoFile.DateLastModified)
            End Select
        
          
End Function
