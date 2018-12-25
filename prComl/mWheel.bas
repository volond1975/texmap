Attribute VB_Name = "mWheel"
Option Explicit
'To be able to scroll with mouse wheel within Userform

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                                        ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal Wparam As Long, _
                                        ByVal Lparam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                                       ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'To get hWnd long value of the UserForm
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                                    ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Dim LocalHwnd As Long
Dim LocalPrevWndProc As Long
Dim MyForm As UserForm

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long
'Обработчик событий мыши (CallBack процедура вызываемая из оконного события)
    Dim MouseKeys As Long
    Dim Rotation As Long

    If Lmsg = WM_MOUSEWHEEL Then
        MouseKeys = Wparam And 65535
        Rotation = Wparam / 65536
        'My Form s MouseWheel function
        frmMain.MouseWheel Rotation
    End If
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, Wparam, Lparam)
End Function

Public Sub WheelHook(PassedForm As UserForm)
'Хук для получения событий UserForm
    On Error Resume Next

    Set MyForm = PassedForm
    LocalHwnd = FindWindow("ThunderDFrame", MyForm.Caption)
    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook()
'Отключаем наш обработчик событий
    Dim WorkFlag As Long

    On Error Resume Next
    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
    Set MyForm = Nothing
End Sub
