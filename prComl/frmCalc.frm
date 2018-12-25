VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalc 
   OleObjectBlob   =   "frmCalc.frx":0000
   Caption         =   "Калькулятор для Excel"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   133
End
Attribute VB_Name = "frmCalc"
Attribute VB_Base = "0{28BE4053-ECAB-4B55-B506-FAAE281D9F39}{CFB55B0A-B436-4CD5-9164-B771708DD50B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Private Const GWL_STYLE As Long = (-16)
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000

Private Const vbKeyEqual = &HBB
Private Const vbKeyMinus = &HBD
Private Const vbKeyPeriod = &HBE
Private Const vbKeySlash = &HBF

Private NewCalc As Boolean
Private Editing As Boolean
Private sFormula As String
Private InMemory As Single
Dim iLastInputValue As Single
Dim iLastInputSign As String
Dim DecimalAllow As Boolean

Private Sub cmd0_Click()
    HandleNumberClick "0"
End Sub

Private Sub cmd1_Click()
    HandleNumberClick "1"
End Sub

Private Sub cmd2_Click()
    HandleNumberClick "2"
End Sub

Private Sub cmd3_Click()
    HandleNumberClick "3"
End Sub

Private Sub cmd4_Click()
    HandleNumberClick "4"
End Sub

Private Sub cmd5_Click()
    HandleNumberClick "5"
End Sub

Private Sub cmd6_Click()
    HandleNumberClick "6"
End Sub

Private Sub cmd7_Click()
    HandleNumberClick "7"
End Sub

Private Sub cmd8_Click()
    HandleNumberClick "8"
End Sub

Private Sub cmd9_Click()
    HandleNumberClick "9"
End Sub

Private Sub cmdBack_Click()
    HandleBackspace
End Sub

Private Sub cmdClear_Click()
    HandleClearClick
End Sub

Private Sub cmdDecimal_Click()
    HandleDecimalClick
End Sub

Private Sub cmdAdd_Click()
    If Me.TxtBoxReadOut.Text <> "0" Then
        HandleOperatorClick "+"
    End If
End Sub

Private Sub cmdDivide_Click()
    If Me.TxtBoxReadOut.Text <> "0" Then
        HandleOperatorClick "/"
    End If
End Sub

Private Sub cmdEquals_Click()
    HandleOperatorClick "="
End Sub

Private Sub cmdMC_Click()
    InMemory = 0
    Me.lblShow = ""
    Me.lblShow.ControlTipText = ""
End Sub

Private Sub cmdMplus_Click()
    If InMemory = 0 Then
        InMemory = CSng(Me.TxtBoxReadOut.Text)
        Me.lblShow = "M"
        Me.lblShow.ControlTipText = InMemory
    Else
        If Me.TxtBoxReadOut.Text <> "0" Then
            If OperatorsPresent(Me.TxtBoxReadOut.Text) Then
                sFormula = Replace(Me.TxtBoxReadOut.Text, ",", ".")
                On Error Resume Next
                Me.TxtBoxReadOut.Text = Evaluate(sFormula)
                If Err <> 0 Then
                    MsgBox "Неверная формула!", 48, "Ошибка"
                    Err.Clear
                    Exit Sub
                End If
            End If
            InMemory = InMemory + CSng(Me.TxtBoxReadOut.Text)
            Me.lblShow.ControlTipText = InMemory
        End If
    End If
End Sub

Private Sub cmdMR_Click()
    If InMemory <> 0 Then Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & InMemory
    Editing = True
End Sub

Private Sub cmdMS_Click()
    If Me.TxtBoxReadOut.Text = "0" Then Exit Sub
    If OperatorsPresent(Me.TxtBoxReadOut.Text) Then
        sFormula = Replace(Me.TxtBoxReadOut.Text, ",", ".")
        On Error Resume Next
        Me.TxtBoxReadOut.Text = Evaluate(sFormula)
        If Err <> 0 Then
            MsgBox "Неверная формула!", 48, "Ошибка"
            Err.Clear
            Exit Sub
        End If
    End If
    InMemory = CSng(Me.TxtBoxReadOut.Text)
    Me.lblShow = "M"
    Me.lblShow.ControlTipText = InMemory
    NewCalc = True
End Sub

Private Sub cmdMultiply_Click()
    If Me.TxtBoxReadOut.Text <> "0" Then
        HandleOperatorClick "*"
    End If
End Sub

Private Sub cmdPercent_Click()
    HandleOperatorClick "%"
End Sub

Private Sub cmd1x_Click()
    If Me.TxtBoxReadOut.Text = "Деление на нуль запрещено" Then Me.TxtBoxReadOut.Text = "0"
    If Me.TxtBoxReadOut.Text = "0" Then
        Me.TxtBoxReadOut.Text = "Деление на нуль запрещено"
        Exit Sub
    End If
    If OperatorsPresent(Me.TxtBoxReadOut.Text) Then
        MsgBox "Неверная формула!", 48, "Ошибка"
        Exit Sub
    End If
    If Me.ListBoxHistory.ListCount = 5 Then
        Me.ListBoxHistory.RemoveItem (Me.ListBoxHistory.ListCount - 1)
    End If
    On Error Resume Next
    Me.ListBoxHistory.AddItem "1/" & Me.TxtBoxReadOut.Text & "=" & Evaluate(1 / CDbl(Me.TxtBoxReadOut.Text)), 0
    Me.TxtBoxReadOut.Text = Evaluate(1 / CDbl(Me.TxtBoxReadOut.Text))
    NewCalc = True
End Sub

Private Sub HandleDecimalClick()
    If InStr(1, Me.TxtBoxReadOut.Text, ",") = 0 Then DecimalAllow = True
    If Not NewCalc Then
        If Not DecimalAllow Then Exit Sub
        If Me.TxtBoxReadOut.Text <> "0" Then
            If Right(Me.TxtBoxReadOut.Text, 1) = "," Then Exit Sub
            If Not OperatorsPresent(Right(Me.TxtBoxReadOut.Text, 1)) Then
                Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & ","
                DecimalAllow = False
                Exit Sub
            Else
                Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & "0,"
                DecimalAllow = False
                Exit Sub
            End If
        Else
            Me.TxtBoxReadOut.Text = "0,"
            DecimalAllow = False
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdPlusMinus_Click()
Dim MinusPos As Long, I As Long
    If Me.TxtBoxReadOut.Text = "0" Then
        Me.TxtBoxReadOut.Text = "-"
    Else
        sFormula = Me.TxtBoxReadOut.Text
        If Not OperatorsPresent(sFormula) Then
            MinusPos = InStrRev(sFormula, "-")
            If MinusPos = 0 Then
                Me.TxtBoxReadOut.Text = "-" & sFormula
                Exit Sub
            Else
                If MinusPos = 1 Then
                    Me.TxtBoxReadOut.Text = Right(sFormula, Len(sFormula) - 1)
                    Exit Sub
                End If
                Me.TxtBoxReadOut.Text = Left(sFormula, Len(sFormula) - MinusPos) & Right(sFormula, MinusPos - 1)
                Exit Sub
            End If
        Else
            For I = Len(sFormula) To 1 Step -1
                If Mid(sFormula, I, 1) = "-" Then
                    If Mid(sFormula, I - 1, 1) Like "[-%+/*]" Then
                        Me.TxtBoxReadOut.Text = Left(sFormula, I - 1) & Right(sFormula, Len(sFormula) - I)
                        Exit Sub
                    End If
                    If IsNumeric(Mid(sFormula, I - 1, 1)) Then
                        Me.TxtBoxReadOut.Text = Left(sFormula, I) & "-" & Right(sFormula, Len(sFormula) - I)
                        Exit Sub
                    End If
                End If
                If Mid(sFormula, I, 1) Like "[%+/*]" Then
                    Me.TxtBoxReadOut.Text = Left(sFormula, I) & "-" & Right(sFormula, Len(sFormula) - I)
                    Exit Sub
                End If
            Next I
        End If
        If CanAddMinus(Me.TxtBoxReadOut.Text) Then Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & "-"
    End If
End Sub

Private Sub cmdSquareRoot_Click()
    If Me.TxtBoxReadOut.Text = "0" Then Exit Sub
    If OperatorsPresent(Me.TxtBoxReadOut.Text) Then
        MsgBox "Неверная формула!", 48, "Ошибка"
        Exit Sub
    End If
    If CDbl(Me.TxtBoxReadOut.Text) < 0 Then
        MsgBox "Нельзя взять квадратный корень отрицательного числа!", 48, "Ошибка"
        Exit Sub
    End If
    Me.TxtBoxReadOut.Text = WorksheetFunction.Power(CDbl(Me.TxtBoxReadOut.Text), (1 / 2))
End Sub

Private Sub cmdSubtract_Click()
    If Me.TxtBoxReadOut.Text <> "0" Then
        HandleOperatorClick "-"
    End If
End Sub

Private Sub cmdInsert_Click()
    If OperatorsPresent(Me.TxtBoxReadOut.Text) Then
        On Error Resume Next
        sFormula = Replace(Me.TxtBoxReadOut.Text, ",", ".")
        Me.TxtBoxReadOut.Text = Evaluate(sFormula)
        If Err <> 0 Then
            MsgBox "Неверная формула!", 48, "Ошибка"
            Err.Clear
        End If
    End If
    ActiveCell.value = CDbl(Me.TxtBoxReadOut.Text)
    NewCalc = True
    If CloseProgByInsert Then Unload Me
End Sub

Private Sub CommandButton1_Click()
Me.TxtBoxReadOut.Text = Worksheets("Калькулятор").Range("$C$35").value
End Sub

Private Sub CommandButton10_Click()

End Sub

Private Sub CommandButton18_Click()

End Sub

Private Sub lblSettings_Click()
    Me.Hide
    frmSettings.Show
End Sub

Private Sub lblValueFromCell_Click()
Dim Rng As Range

    If Not NewCalc Then
        If Me.TxtBoxReadOut <> "0" Then
            If Not Repeating Then
                MsgBox "Сперва введите знак математического действия ( / + - * )", 48, "Ошибка"
                Exit Sub
            End If
        End If
    End If
    Me.Hide

    On Error Resume Next

    Set Rng = Application.InputBox("Укажите мышкой на ячейку со значением", "Укажите ячейку", Selection.address, , , , , 8)

    If Rng Is Nothing Then GoTo ExitMrk

    On Error GoTo 0

    If Rng.Cells.Count > 1 Then
        MsgBox "Укажите только одну ячейку!", 48, "Ошибка"
        GoTo ExitMrk
    End If

    If Not IsNumeric(Rng.value) Then
        MsgBox "Ячека должна содержать число!", 48, "Ошибка"
        Me.Show vbModeless
        Exit Sub
    End If
    If Rng.value = 0 Then GoTo ExitMrk

    If Me.TxtBoxReadOut.Text = "0" Then
        Me.TxtBoxReadOut.Text = Rng.value
    Else
        If NewCalc Then
            Me.TxtBoxReadOut.Text = Rng.value
        Else
            Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & Rng.value
        End If
    End If

ExitMrk:
    NewCalc = False
    Me.Show vbModeless
End Sub

Private Sub lblAbout_Click()
    Me.Hide
    frmAbout.Show
End Sub

Private Sub ListBoxHistory_Click()
Dim I As Long

    With Me.ListBoxHistory
        If .ListCount > 0 Then
            Me.TxtBoxReadOut.Text = Left(.value, InStr(1, .value, "=", vbTextCompare) - 1)
        End If
        For I = 0 To .ListCount - 1
            .Selected(I) = False
        Next I
    End With

    Me.TxtBoxReadOut.SetFocus
    Editing = True
    NewCalc = False
End Sub

Private Sub TxtBoxReadOut_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmdEquals = True
        Me.TxtBoxReadOut.SetFocus
        Exit Sub
    End If

    If KeyCode = vbKeyBack Then
        If Editing Then Exit Sub
        KeyCode = 0
        cmdBack = True
        Me.TxtBoxReadOut.SetFocus
        Exit Sub
    End If

    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        InitCalc
        Me.TxtBoxReadOut.SetFocus
        Exit Sub
    End If

    If KeyCode = vbKeyInsert Then
        KeyCode = 0
        cmdInsert = True
        Exit Sub
    End If

    HandleNumberPress KeyCode, Shift
    If Not Editing Then KeyCode = 0
    Me.TxtBoxReadOut.SetFocus
End Sub

Private Sub UserForm_Initialize()
Dim lngFrmHndl As Long, lngStyle As Long
Dim iFile As String, ihIcon As Long

    lngFrmHndl = FindWindow(vbNullString, Me.Caption)

    'иконка
    iFile = "C:\WINDOWS\system32\calc.exe"
    ihIcon = ExtractIcon(lngFrmHndl, iFile, 0&)
    SetClassLong lngFrmHndl, -14&, ihIcon
    DestroyIcon ihIcon

    'меню
    lngStyle = GetWindowLong(lngFrmHndl, GWL_STYLE)
    lngStyle = lngStyle Or WS_SYSMENU
    lngStyle = lngStyle Or WS_MINIMIZEBOX
    'lngStyle = lngStyle Or WS_MAXIMIZEBOX
    SetWindowLong lngFrmHndl, GWL_STYLE, (lngStyle)
    DrawMenuBar lngFrmHndl
End Sub

Private Sub HandleBackspace()
    If Me.TxtBoxReadOut.Text = "" Then Exit Sub
    If Me.TxtBoxReadOut.Text = "Деление на нуль запрещено" Then
        Me.TxtBoxReadOut.value = 0
        Exit Sub
    End If

    If Me.TxtBoxReadOut.Text <> "0" Then
        Me.TxtBoxReadOut.Text = Left$(TxtBoxReadOut.Text, Len(TxtBoxReadOut.Text) - 1)
    End If
    If Me.TxtBoxReadOut.Text = "" Then Me.TxtBoxReadOut.value = 0
End Sub

Private Sub HandleClearClick()
    InitCalc
End Sub

Private Sub HandleClearEntry()
    Me.TxtBoxReadOut.value = 0
End Sub

Private Sub HandleNumberClick(strNum As String)
    If NewCalc Then Me.TxtBoxReadOut.value = 0
    If Me.TxtBoxReadOut.Text = "0" Then
        Me.TxtBoxReadOut.Text = strNum
    Else
        Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & strNum
    End If
    iLastInputValue = FindLastInputValue(Me.TxtBoxReadOut.Text)
    NewCalc = False
End Sub

Private Function FindLastInputValue(sFormula) As Single
Dim I As Long
    For I = Len(sFormula) To 1 Step -1
        If Mid(sFormula, I, 1) Like "[-+/*]" Then
            FindLastInputValue = Right(sFormula, Len(sFormula) - I)
            Exit Function
        End If
    Next I
    FindLastInputValue = sFormula
End Function

Private Sub HandleNumberPress(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim strChar As String
Dim intShiftDown As Integer

    If Editing Then Exit Sub
'    If Editing Then
'        Select Case KeyCode
'            Case 65 To 122
'                KeyCode = 0
'            Case 192 To 255
'                KeyCode = 0
'        End Select
'        Exit Sub
'    End If
    'if / * - +
    If KeyCode = 106 Or KeyCode = 107 Or KeyCode = 109 Or KeyCode = 111 Then NewCalc = False
    If NewCalc Then Me.TxtBoxReadOut.value = 0
    intShiftDown = Shift
    Select Case KeyCode
            '0 - 9
        Case 48 To 57
            If intShiftDown Then
                Select Case KeyCode
                        'Handle Shift-8 ("*")
                    Case 48 + 8
                        HandleOperatorClick "*"
                    Case 48    ' (
                        strChar = Chr$(41)
                        HandleNumberClick strChar
                    Case 57    ' )
                        strChar = Chr$(40)
                        HandleNumberClick strChar
                End Select
            Else
                strChar = Chr$(KeyCode)
                HandleNumberClick strChar
            End If

            'Backspace
        Case vbKeyBack
            HandleBackspace

        Case vbKeyDelete
            If Editing Then Exit Sub
            HandleClearEntry

            'Numpad 0 - Numpad 9
        Case vbKeyNumpad0 To vbKeyNumpad9
            strChar = Chr$(KeyCode - vbKeyNumpad0 + Asc("0"))
            HandleNumberClick strChar

            'Period and Decimal
        Case vbKeyDecimal, vbKeyPeriod
            HandleDecimalClick

        Case vbKeySubtract, vbKeyMinus
            HandleOperatorClick "-"

        Case vbKeyMultiply
            HandleOperatorClick "*"

        Case vbKeyAdd
            HandleOperatorClick "+"

        Case vbKeyDivide, vbKeySlash
            HandleOperatorClick "/"

        Case vbKeyEqual, vbKeyReturn
            If intShiftDown Then
                HandleOperatorClick "+"
            Else
                HandleOperatorClick "="
            End If

        Case vbKeyC
            HandleClearClick

        Case 65 To 122
            KeyCode = 0

        Case 192 To 255
            KeyCode = 0
    End Select
End Sub

Private Sub HandleOperatorClick(strOp As String)
    If strOp Like "[-/+%*]" Then DecimalAllow = True
    If strOp = "%" Then
        If Right(Me.TxtBoxReadOut.Text, 1) = "%" Then Exit Sub
    End If
    If strOp Like "[-+/*]" Then iLastInputSign = strOp
    If Me.TxtBoxReadOut.Text = "0" Then Exit Sub
    
    If strOp = "=" And Not OperatorsPresent(Me.TxtBoxReadOut.Text) Then
        If CDbl(Me.TxtBoxReadOut.Text) <> "0" Then
            If iLastInputValue <> 0 And iLastInputSign <> "" Then
                'add to history
                If Me.ListBoxHistory.ListCount = 5 Then
                    Me.ListBoxHistory.RemoveItem (Me.ListBoxHistory.ListCount - 1)
                End If
                sFormula = Replace(Me.TxtBoxReadOut.Text, ",", ".")
                On Error Resume Next
                Me.ListBoxHistory.AddItem Me.TxtBoxReadOut.Text & _
                        iLastInputSign & iLastInputValue & "=" & _
                        Evaluate(CDbl(Me.TxtBoxReadOut.Text) & iLastInputSign & _
                        CDbl(iLastInputValue)), (0)
                Me.TxtBoxReadOut.Text = Evaluate(CDbl(Me.TxtBoxReadOut.Text) & _
                        iLastInputSign & CDbl(iLastInputValue))
                If Err <> 0 Then
                    MsgBox "Неверная формула!", 48, "Ошибка"
                    Err.Clear
                End If
                NewCalc = True
                Editing = False
                Exit Sub
            End If
        End If
    End If

    If strOp = "=" Then
        EqualBtn
        Exit Sub
    End If

    If strOp = "-" Then
        If CanAddMinus(Me.TxtBoxReadOut.Text) Then Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & strOp
    End If

    If Not Repeating Then Me.TxtBoxReadOut.Text = Me.TxtBoxReadOut.Text & strOp
    NewCalc = False
End Sub

Private Sub EqualBtn()
    sFormula = Me.TxtBoxReadOut.Text
    If Repeating Then Me.TxtBoxReadOut.Text = Left(sFormula, Len(sFormula) - 1)
    If OperatorsPresent(Me.TxtBoxReadOut.Text) Then Call AddToHistory
    On Error Resume Next
    sFormula = Replace(Me.TxtBoxReadOut.Text, ",", ".")
    Me.TxtBoxReadOut.Text = Evaluate(sFormula)
    If Err <> 0 Then
        MsgBox "Неверная формула!", 48, "Ошибка"
        Err.Clear
    End If
    NewCalc = True
    Editing = False
    Exit Sub
End Sub

Private Sub InitCalc()
    Me.TxtBoxReadOut.Text = "0"
End Sub

Private Function Repeating() As Boolean
    Repeating = Right(Me.TxtBoxReadOut.Text, 1) Like "[-+,./*]"
End Function

Private Function OperatorsPresent(sFormula As String) As Boolean
Dim I As Long
    For I = 1 To Len(sFormula)
        If Mid(sFormula, I, 1) Like "[-%+/*]" Then
            If Mid(sFormula, I, 1) <> "," Then
                If I <> 1 Or Mid(sFormula, I, 1) <> "-" Then
                    OperatorsPresent = True
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Private Function CanAddMinus(sFormula As String) As Boolean
    If IsNumeric(Right(sFormula, 1)) Then
        CanAddMinus = True
        Exit Function
    End If
    If Left(Right(sFormula, 2), 1) Like "[-%+/*]" Then
        CanAddMinus = False
        Exit Function
    End If
    If Right(sFormula, 1) Like "[-%+/*]" Then
        CanAddMinus = True
        Exit Function
    End If
End Function

Private Sub AddToHistory()
    If Me.ListBoxHistory.ListCount = 5 Then
        Me.ListBoxHistory.RemoveItem (Me.ListBoxHistory.ListCount - 1)
    End If
    sFormula = Replace(Me.TxtBoxReadOut.Text, ",", ".")
    On Error Resume Next
    Me.ListBoxHistory.AddItem Me.TxtBoxReadOut.Text & "=" & Evaluate(sFormula), (0)
End Sub

Private Function CloseProgByInsert() As String
    CloseProgByInsert = GetSetting(PROGID, "Settings", "CloseByInsertKey", False)
End Function

Private Sub UserForm_Terminate()
Dim iFile As String, ihIcon As Long, lngFrmHndl As Long
    lngFrmHndl = FindWindow(vbNullString, Me.Caption)
    iFile = ""
    ihIcon = ExtractIcon(lngFrmHndl, iFile, 0&)
    SetClassLong lngFrmHndl, -14&, ihIcon
    DestroyIcon ihIcon
End Sub
