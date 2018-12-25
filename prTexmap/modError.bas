Attribute VB_Name = "modError"
Public Const MARKER As String = "NOT_TOPMOST"


Sub LogInfo(LogMessage As String)


    'set path and name of the log file where you want to save
    'the log file
    namepatn = ThisWorkbook.Path
    Const LogFileName As String = "c:\temp\LogFile.LOG"
    Dim FileNum As Integer
    
    FileNum = FreeFile ' next file number
    Open LogFileName For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file


End Sub
Option Explicit


' Reraises an error and adds line number and current procedure name
Sub RaiseError(ByVal errorno As Long, ByVal src As String _
                , ByVal proc As String, ByVal desc As String, ByVal lineno As Long)
    
    Dim sLineNo As Long, sSource As String
    
    ' If no marker then this is the first time RaiseError was called
    If Left(src, Len(MARKER)) <> MARKER Then

        ' Add error line number if present
        If lineno <> 0 Then
            sSource = vbCrLf & "Line no: " & lineno & " "
        End If
   
        ' Add marker and procedure to source
        sSource = MARKER & sSource & vbCrLf & proc
        
    Else
        ' If error has already been raised then just add on procedure name
        sSource = src & vbCrLf & proc
    End If
    
    ' If the code stops here, make sure DisplayError is placed in the top most Sub
    Err.Raise errorno, sSource, desc
    
End Sub

' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub
Sub DisplayError(ByVal src As String, ByVal desc As String _
                    , ByVal sProcname As String)

    ' Remove the marker
    src = Replace(src, MARKER, "")
    
    Dim sMsg As String
    sMsg = "The following error occurred: " & vbCrLf & Err.Description _
                    & vbCrLf & vbCrLf & "Error Location is: "
    
    sMsg = sMsg + src & vbCrLf & sProcname
    
    ' Display message
    MsgBox sMsg, Title:="Error"
    
End Sub

'An Example of using this strategy
''Here is a simple coding that use these subs.
'In this strategy, we don’t place any code in the topmost sub. We only call subs from it.
'https://excelmacromastery.com/vba-error-handling/
Sub Topmost()

    On Error GoTo EH
    
    Level1

Done:
    Exit Sub
EH:
    DisplayError Err.Source, Err.Description, "Module1.Topmost"
End Sub

Sub Level1()

    On Error GoTo EH
    
    Level2

Done:
    Exit Sub
EH:
   RaiseError Err.Number, Err.Source, "Module1.Level1", Err.Description, Erl
End Sub

Sub Level2()

    On Error GoTo EH
    
    ' Error here
    Dim a As Long
    a = "7 / 0"

Done:
    Exit Sub
EH:
    RaiseError Err.Number, Err.Source, "Module1.Level2", Err.Description, Erl
End Sub
