Attribute VB_Name = "mTelegram"
Dim URL
Dim BOT_ID
Dim TOKEN
Dim Offset

'https://habrahabr.ru/post/306222/
'http://imperiya.by/video/nPJIa_ra4Op/Telegram-Bot-Tutorial-How-to-connect-your-Telegram-Bot-to-a-Google-Spreadsheet-Apps-Script.html






Function Send_to_Telegram_Bot_example(chat_id, txt, Optional bResponse = True)
Dim oHttp As Object
Dim sURI As String
'Создаёшь нового бота со смартфона через @FatherBot
'От присылает токен типа "000000000:AaAaAaAaAaAaAaAaAaAaAaAaAaAaAaAaAaAa"
'Посылаешь в браузере команду вида https://api.telegram.org/bot000000000:AaAaAaAaAaAaAaAaAaAaAaAaAaAaAaAaAaAa/getupdates и изнаёшь свой chat_id
'Посылаешь ...../sendmessage?chat_id=88888888&text=hello - и получаешь первое сообщение. Ура!
'Очень удобно! Сотрудник обновил расчеты на листе, а руководителю (мне) прилетели итоги на смартфон в telegram

TOKEN = "215801129:AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"
'chat_id = "192818801"
BotID = "215801129"
BotToken = "AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"

'ChatIDInna = "372782755"

'500866893: AAH550ElMohowqZhZecp52BIeNbsokjdjzw
'txt = "hello"
 s = InitBot(BotID, BotToken)
 Debug.Print URL
'sURI = "https://api.telegram.org/bot000000000:Aa....AaAa/getme"
'sURI = "https://api.telegram.org/bot000000000:Aa....AaAa/getupdates"
 
'sendproto не работает!!! Блин, не знаю, как скормить строке имя файла С:\temp\pic.png?
'запрос с файлом должен быть в формате multipart/form-data
'sURI = "https://api.telegram.org/bot000000000:Aa.....AaAa/sendproto?chat_id=88888888@photo=C:\temp\pic.png"
' https://api.telegram.org/bot215801129:AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I/getupdates
sURI = "https://api.telegram.org/bot215801129:AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I/sendmessage?chat_id=" & chat_id & "&text=" & txt
 
'MsgBox sURI, vbInformation, "запрос"
On Error Resume Next
Set oHttp = CreateObject("MSXML2.XMLHTTP")
If Err.Number <> 0 Then
Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
End If
On Error GoTo 0
If oHttp Is Nothing Then Exit Function
oHttp.Open "GET", sURI, False
oHttp.Send
ResponseText = oHttp.ResponseText

Send_to_Telegram_Bot_example = FNResponseText(ResponseText)
Set oHttp = Nothing
End Function
Function RussianStringToURLEncode_New(ByVal txt As String) As String
    For i = 1 To Len(txt)
        l = Mid(txt, i, 1)
        Select Case AscW(l)
            Case Is > 4095: t = "%" & Hex(AscW(l) \ 64 \ 64 + 224) & "%" & Hex(AscW(l) \ 64) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
            Case Is > 127: t = "%" & Hex(AscW(l) \ 64 + 192) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
            Case 32: t = "%20"
            Case Else: t = l
        End Select
        RussianStringToURLEncode_New = RussianStringToURLEncode_New & t
    Next
End Function
 
Sub TextUTF8Send()
ChatID = "192818801"
MsgBox "MSGID" & Send_to_Telegram_Bot_example(ChatID, RussianStringToURLEncode_New(ActiveCell.value))
End Sub

Function FNResponseText(ResponseText, Optional bResponse = False)
Dim OBJResponseText As Object
Dim STRDATE As Long
Set OBJResponseText = ParseJson(ResponseText)
If OBJResponseText("ok") Then
If bResponse Then MsgBox ResponseText, vbInformation, "ответ"
Set OBJRESULT = OBJResponseText("result")
MSGID = OBJResponseText("result")("message_id")
Text = OBJResponseText("result")("text")
STRDATE = OBJResponseText("result")("date")
MDate = Unix2Date(STRDATE)
Set from = OBJResponseText("result")("from")
Set CHAT = OBJResponseText("result")("chat")
Debug.Print ResponseText
FNResponseText = MSGID
End If

End Function

' #cs ===============================================================================
'   Function Name..:     _InitBot()
'   Description....:     Initialize your Bot with BotID and Token
'   Parameter(s)...:     $BotID - Your Bot ID (12345..)
'                        $BotToken - Your Bot Token (AbCdEf...)
'   Return Value(s):     Return True
'#ce ===============================================================================
 
 Function InitBot(BotID, BotToken)
 URL = "https://api.telegram.org/bot"
 z = GetURLstatus("https://GOOGLE.COM")
 Offset = 0
 BOT_ID = BotID
 TOKEN = BotToken
 URL = URL & BOT_ID & ":" & TOKEN
 InitBot = True
 End Function
' #cs ===============================================================================
'   Function Name..:     _Polling()
'   Description....:     Wait for incoming messages from user
'   Parameter(s)...:     None
'   Return Value(s):     Return an array with information about messages:
'                           $msgData[0] = Offset of the current update (used to 'switch' to next update)
'                           $msgData[1] = Username of the user
'                           $msgData[2] = ChatID used to interact with the user
'                           $msgData[3] = Text of the message
'#ce ===============================================================================
'Func _Polling()
'   While 1
'      Sleep(1000) ;Prevent CPU Overloading
'      $newUpdates = _GetUpdates()
'      If Not StringInStr($newUpdates,'update_id') Then ContinueLoop
'      $msgData = _JSONDecode($newUpdates)
'      $offset = $msgData[0] + 1
'      Return $msgData
'   Wend
'EndFunc
 
 
' ===============================================================================
'   Function Name..:     _GetUpdates()
'   Description....:     Used by _Polling() to get new messages
'   Parameter(s)...:     None
'   Return Value(s):     Return string with information encoded in JSON format
'#ce ===============================================================================
Sub testsetWebhook()
TOKEN = "215801129:AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"
chat_id = "19281880"
BotID = "215801129"
BotToken = "AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"
Call InitBot(BotID, BotToken)
Debug.Print URL
s = setWebhook()
MsgBox s, vbInformation, "ответ"
End Sub
Function testGetUpdates(Optional TOKEN)
TOKEN = "215801129:AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"
m = Split(TOKEN, ":")
chat_id = m(0) '"19281880"
BotID = m(1) '"215801129"
BotToken = "AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"






Call InitBot(BotID, BotToken)
Debug.Print URL
Debug.Print st
st = GetUpdates()
Dim JSON As Object
Set JSON = JsonConverter.ParseJson(st)
 If JSON("ok") Then
 MsgBox st, vbInformation, "ответ"
 JSON("result").Count
 Else
 End If
'MsgBox s, vbInformation, "ответ"
End Function




'https://api.telegram.org/bot215801129:AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I/setWebhook
Function setWebhook()
  setWebhook = HttpGet(URL & "/setWebhook")
End Function '==> _setWebhook


Function GetUpdates()
  GetUpdates = HttpGet(URL & "/getUpdates?offset=" & Offset)
End Function '==> _GetUpdates

'  cs ===============================================================================
'   Function Name..:     _GetMe()
'   Description....:     Get information about the bot (like name, @botname...)
'   Parameter(s)...:     None
'   Return Value(s):     Return string with information encoded in JSON format
'#ce ===============================================================================
Sub testgetme()
TOKEN = "215801129:AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"
chat_id = "19281880"
BotID = "215801129"
BotToken = "AAFkqUwJEV_N6dFLDYPL19bk7OHl_K67A4I"
Call InitBot(BotID, BotToken)
Debug.Print URL
s = GetMe()

MsgBox s, vbInformation, "ответ"

End Sub


Function GetMe()

   GetMe = HttpGet(URL & "/getMe")

End Function '==>_GetMe
'#cs ===============================================================================
'   Function Name..:     _SendDocument()
'   Description....:     Send a document
'   Parameter(s)...:     $ChatID: Unique identifier for the target chat
'                        $Path: Path to local file
'                        $Caption: Caption to send with document (optional)
'   Return Value(s):     Return File ID of the video as string
'#ce ===============================================================================
'Func _SendDocument($ChatID, $Path, $Caption = "")
'   Local $Query = $URL & '/sendDocument'
'   Local $hOpen = _WinHttpOpen()
'   Local $Form = '<form action="' & $Query & '" method="post" enctype="multipart/form-data">' & _
'                  ' <input type="text" name="chat_id"/>'  & _
'                  ' <input type="file" name="document"/>' & _
'                  ' <input type="text" name="caption"/>'  & _
'                 '</form>'
'   Local $Response = _WinHttpSimpleFormFill($Form, $hOpen, Default, _
'                        "name:chat_id",  $ChatID, _
'                        "name:document", $Path, _
'                        "name:caption",  $Caption)
'   _WinHttpCloseHandle($hOpen)
'   Return _GetFileID($Response)
'EndFunc ;==> _SendDocument

'#cs ===============================================================================
'   Function Name..:     _GetChat()
'   Description....:     Get basic information about chat, like username of the user, id of the user
'   Parameter(s)...:     $ChatID: Unique identifier for the target chat
'   Return Value(s):     Return string with information encoded in JSON format
'#ce ===============================================================================
'Func _GetChat($ChatID)
'   Local $Query = $URL & "/getChat?chat_id=" & $ChatID
'   Local $Response = HttpGet($Query)
'   Return $Response
'EndFunc
'#EndRegion
'
'#Region "@BACKGROUND FUNCTION"
'#cs ===============================================================================
'   Function Name..:     _GetFilePath()
'   Description....:     Get path of a specific file (specified by FileID) on Telegram Server
'   Parameter(s)...:     $FileID: Unique identifie for the file
'   Return Value(s):     Return FilePath as String
'#ce ===============================================================================
'Func _GetFilePath($FileID)
'   Local $Query = $URL & "/getFile?file_id=" & $FileID
'   Local $Response = HttpPost($Query)
'   Local $firstSplit = StringSplit($Response,':')
'   Local $FilePath = StringTrimLeft($firstSplit[6],1)
'   $FilePath = StringTrimRight($FilePath,3)
'   Return $FilePath
'EndFunc
'
'#cs ===============================================================================
'   Function Name..:     _GetFileID()
'   Description....:     Get file ID of the last uploaded file
'   Parameter(s)...:     $Output: Response from HTTP Request
'   Return Value(s):     Return FileID as String
'#ce ===============================================================================
'Func _GetFileID($Output)
'   If StringInStr($Output,"photo",1) and StringInStr($Output,"width",1) Then
'      Local $firstSplit  = StringSplit($Output,'[')
'      Local $secondSplit = StringSplit($firstSplit[2],',')
'      Local $thirdSplit  = StringSplit($secondSplit[9],':')
'      Local $FileID = StringTrimLeft($thirdSplit[2],1)
'      $FileID = StringTrimRight($FileID,1)
'      Return $FileID
'
'   ElseIf StringInStr($Output,'audio":',1) And StringInStr($Output,'mime_type":"audio',1) Then
'      Local $firstSplit = StringSplit($Output,',')
'      For $i=1 to $firstSplit[0]
'         If StringInStr($firstSplit[$i],"file_id",1) Then Local $secondSplit = StringSplit($firstSplit[$i],':')
'      Next
'      Local $FileID = StringTrimLeft($secondSplit[2],1)
'      $FileID = StringTrimRight($FileID,1)
'      Return $FileID
'
'   ElseIf StringInStr($Output,'video":',1) and StringInStr($Output,"width",1) Then
'      Local $firstSplit = StringSplit($Output,',')
'      For $i=1 to $firstSplit[0]
'         If StringInStr($firstSplit[$i],"file_id",1) and not StringInStr($firstSplit[$i],"thumb",1) Then Local $secondSplit = StringSplit($firstSplit[$i],':')
'      Next
'      Local $FileID = StringTrimLeft($secondSplit[2],1)
'      $FileID = StringTrimRight($FileID,1)
'      Return $FileID
'
'   ElseIf StringInStr($Output,'document":',1) and StringInStr($Output,"text/plain",1) Then
'      Local $firstSplit = StringSplit($Output,',')
'      For $i=1 to $firstSplit[0]
'         If StringInStr($firstSplit[$i],"file_id",1) Then Local $secondSplit = StringSplit($firstSplit[$i],':')
'      Next
'      Local $FileID = StringTrimLeft($secondSplit[2],1)
'      $FileID = StringTrimRight($FileID,1)
'      Return $FileID
'
'   ElseIf StringInStr($Output,'voice":',1) and StringInStr($Output,"audio/ogg",1) Then
'      Local $firstSplit = StringSplit($Output,',')
'      For $i=1 to $firstSplit[0]
'         If StringInStr($firstSplit[$i],"file_id",1) Then Local $secondSplit = StringSplit($firstSplit[$i],':')
'      Next
'      Local $FileID = StringTrimLeft($secondSplit[2],1)
'      $FileID = StringTrimRight($FileID,1)
'      Return $FileID
'
'   ElseIf StringInStr($Output,'sticker":',1) and StringInStr($Output,"width",1) Then
'      Local $firstSplit = StringSplit($Output,',')
'      For $i=1 to $firstSplit[0]
'         If StringInStr($firstSplit[$i],"file_id",1) Then Local $secondSplit = StringSplit($firstSplit[$i],':')
'      Next
'      Local $FileID = StringTrimLeft($secondSplit[2],1)
'      $FileID = StringTrimRight($FileID,1)
'      Return $FileID
'   End If
'EndFunc
'
'#cs ===============================================================================
'   Function Name..:     _DownloadFile()
'   Description....:     Download and save locally a file from the Telegram Server by FilePath
'   Parameter(s)...:     $FilePath: Path of the file on Telegram Server
'   Return Value(s):     Return True
'#ce ===============================================================================
'Func _DownloadFile($FilePath)
'   Local $firstSplit = StringSplit($FilePath,'/')
'   Local $fileName = $firstSplit[2]
'   Local $Query = "https://api.telegram.org/file/bot" & $BOT_ID & ":" & $TOKEN & "/" & $FilePath
'   InetGet($Query,$fileName)
'   Return True
'EndFunc


Sub SendNaryad(fileName, ChatID, PathSourse, ReportName, InfoBotFather)
'
Dim WshShell As Object
On Error Resume Next
With CreateObject("Scripting.FileSystemObject")
'Set wb = ThisWorkbook
'ActiveCell.Copy
sChatID = ShellCMDLineParametr("ChatID", ChatID)
sPathSourse = ShellCMDLineParametr("PathSourse", PathSourse)
sReportName = ShellCMDLineParametr("ReportName", ReportName)
sInfoBotFather = ShellCMDLineParametr("InfoBotFather", InfoBotFather)

file$ = ShellCMDLine(ThisWorkbook.Path & "\" & fileName, sChatID, sPathSourse, sReportName, sInfoBotFather)
ErrorCode = Shell(file$, 0)
myDBQ = wb.Path & "\End.ini"

Do
DoEvents
     If .fileExists(myDBQ) = True Then Exit Do
Loop
End With
End Sub
Function URLDecode(ByVal strIn)
    ' ????? ?????: zhaojunpeng.com/posts/2016/10/28/excel-urldecode
    ' ? ???????? EducatedFool
    On Error Resume Next
    Dim sl&, tl&, key$, kl&
    sl = 1:    tl = 1: key = "%": kl = Len(key)
    sl = InStr(sl, strIn, key, 1)
    Do While sl > 0
        If (tl = 1 And sl <> 1) Or tl < sl Then
            URLDecode = URLDecode & Mid(strIn, tl, sl - tl)
        End If
        Dim hh$, hi$, hl$, a$
        Select Case UCase(Mid(strIn, sl + kl, 1))
            Case "U"    'Unicode URLEncode
                a = Mid(strIn, sl + kl + 1, 4)
                URLDecode = URLDecode & ChrW("&H" & a)
                sl = sl + 6
            Case "E"    'UTF-8 URLEncode
                hh = Mid(strIn, sl + kl, 2)
                a = Int("&H" & hh)    'ascii?
                If Abs(a) < 128 Then
                    sl = sl + 3
                    URLDecode = URLDecode & Chr(a)
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2)
                    hl = Mid(strIn, sl + 6 + kl, 2)
                    a = ("&H" & hh And &HF) * 2 ^ 12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)
                    If a < 0 Then a = a + 65536
                    URLDecode = URLDecode & ChrW(a)
                    sl = sl + 9
                End If
            Case Else    'Asc URLEncode
                hh = Mid(strIn, sl + kl, 2)    '??
                a = Int("&H" & hh)    'ascii?

                If Abs(a) < 128 Then
                    sl = sl + 3
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2)    '??
                    'a = Int("&H" & hh & hi) '?ascii?
                    a = (Int("&H" & hh) - 194) * 64 + Int("&H" & hi)
                    sl = sl + 6
                End If
                URLDecode = URLDecode & ChrW(a)
        End Select
        tl = sl
        sl = InStr(sl, strIn, key, 1)
    Loop
    URLDecode = URLDecode & Mid(strIn, tl)
End Function
Function GetURLstatus(ByVal URL$) As Long
    ' функция проверяет наличие доступа к ресурсу URL$ (файлу или каталогу)
    ' возвращает код ответа сервера (число), либо 0, если ссылка ошибочная
    ' (200 - ресурес доступен, 404 - не найден, 403 - нет доступа, и т.д.)
    On Error Resume Next: URL$ = Replace(URL$, "\", "/")
    Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
    xmlhttp.Open "GET", URL, "False"
    xmlhttp.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
    xmlhttp.Send
    GetURLstatus = Val(xmlhttp.Status)
    Set xmlhttp = Nothing
End Function

Function URLEncode(ByVal txt As String) As String
    For i = 1 To Len(txt)
        l = Mid(txt, i, 1)
        Select Case AscW(l)
            Case Is > 4095: t = "%" & Hex(AscW(l) \ 64 \ 64 + 224) & "%" & Hex(AscW(l) \ 64) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
            Case Is > 127: t = "%" & Hex(AscW(l) \ 64 + 192) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
            Case 32: t = "+"
            Case Else: t = l
        End Select
        URLEncode = URLEncode & t
    Next
End Function
Function UTF8_Decode(ByVal sStr As String)
    Dim l As Long, sUTF8 As String, iChar As Integer, iChar2 As Integer
    For l = 1 To Len(sStr)
        iChar = Asc(Mid(sStr, l, 1))
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
            l = l + 1
        Else
            Dim iChar3 As Integer
            iChar2 = Asc(Mid(sStr, l + 1, 1))
            iChar3 = Asc(Mid(sStr, l + 2, 1))
            sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
            l = l + 2
        End If
            Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next l
    UTF8_Decode = sUTF8
End Function
