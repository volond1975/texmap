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
'������ ���������� ����� � ����������
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
'��������� �� ���������� ����� � ���������� �������� ���������� ����� � ��
'�� ��������� ���� �� ������ �� ��������� ���� ����
'������ �������� ��� "2:10"
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
    ' �������� � �������� ��������� ���� � ����� FolderPath,
    ' ����� ����� ������� ������ Mask (����� �������� ������ ����� � ����� ������/�����������)
    ' � ������� ������ SearchDeep � ��������� (���� SearchDeep=1, �� �������� �� ���������������).
    ' ���������� ���������, ���������� ������ ���� ��������� ������
    ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)

    Set FilenamesCollection = New Collection    ' ������ ������ ���������
    Set fso = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, fso, FilenamesCollection, SearchDeep ' �����
    Set fso = Nothing: Application.StatusBar = False    ' ������� ������ ��������� Excel
End Function
 
Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef fso, _
                                 ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    ' ���������� ��� ����� � �������� � ����� FolderPath, ��������� ������ FSO
    ' ������� ����� �������������� � ��� ������, ���� SearchDeep > 1
    ' ��������� ���� ��������� ������ � ��������� FileNamesColl
    On Error Resume Next: Set curfold = fso.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����

        ' ���������������� ��� ������ ��� ������ ���� � ���������������
        ' � ������� ������ ����� � ������ ��������� Excel
        ' Application.StatusBar = "����� � �����: " & FolderPath

        For Each fil In curfold.Files    ' ���������� ��� ����� � ����� FolderPath
            If fil.name Like "*" & Mask Then FileNamesColl.Add fil.Path
        Next
        SearchDeep = SearchDeep - 1    ' ��������� ������� ������ � ���������
        If SearchDeep Then    ' ���� ���� ������ ������
            For Each sfol In curfold.SubFolders    ' ���������� ��� �������� � ����� FolderPath
                GetAllFileNamesUsingFSO sfol.Path, Mask, fso, FileNamesColl, SearchDeep
            Next
        End If
        Set fil = Nothing: Set curfold = Nothing    ' ������� ����������
    End If
End Function

Sub ������_FilenamesCollection()
    On Error Resume Next
    Dim folder$, coll As Collection
 
    folder$ = ThisWorkbook.Path & "\�������\"
    If Dir(folder$, vbDirectory) = "" Then
        MsgBox "�� ������� ����� �" & folder$ & "�", vbCritical, "��� ����� �������"
        Exit Sub        ' �����, ���� ����� �� �������
    End If
 
    Set coll = FilenamesCollection(folder$, "*.xls")        ' �������� ������ ������ XLS �� �����
    If coll.Count = 0 Then
        MsgBox "� ����� �" & Split(folder$, "\")(UBound(Split(folder$, "\")) - 1) & "� ��� �� ������ ����������� �����!", _
               vbCritical, "����� ��� ��������� �� �������"
        Exit Sub        ' �����, ���� ��� ������
    End If
 
    ' ���������� ��� ��������� �����
    For Each file In coll
        Debug.Print file        ' ������� ��� ����� � ���� Immediate
    Next
End Sub


 Function CreateFolder(FullPath As String)
 Dim ps As PathSplitString
' myPath & "\" & z.Value
With CreateObject("Scripting.FileSystemObject")
'Set lis = ���������������������������������("˳�������", "˳�������", "˳�������")
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
Function DataCreatedFile(FilePath As String, ��������)
        Dim objFSO As Scripting.FileSystemObject
        Dim fsoFile, DateCreate
'        Dim FilePath As String
            Set objFSO = New Scripting.FileSystemObject
          
            Set fsoFile = objFSO.GetFile(FilePath)
            Select Case ��������
            Case "���� ��������"
            DataCreatedFile = (fsoFile.DateCreated)
            Case "��������� ������"
            DataCreatedFile = (fsoFile.DateLastAccessed)
            Case "��������� ���������"
             DataCreatedFile = (fsoFile.DateLastModified)
            End Select
        
          
End Function
