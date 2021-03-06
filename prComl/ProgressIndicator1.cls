VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressIndicator1"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
 
Sub ProgressIndicator()
 '---------------------------------------------------------------------------------------
 ' Class Module      : ProgressIndicator
 ' �����     : EducatedFool  (�����)                    ����: 07.10.2012
 ' ���������� �������� ����� ��������� ��� Microsoft Excel
 ' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
 ' ��������� ��� ������ ������: http://ExcelVBA.ru/payments
 '---------------------------------------------------------------------------------------

 Public FP As New F_Progress

 Public SubActionIndex As Single, SubActionsCount As Single
 Private FPVisible As Boolean, FPStartTime As Date, Position As Integer
 Private PrS As Integer, PrE As Integer, Percent As Double, LogString As String

 Public Parent As ProgressIndicator
 Public ShowPercents As Boolean, ShowTime As Boolean, ShowTimeInLog As Boolean
 Public Children As New Collection

 Function AddChildIndicator(ByVal Caption As String, Optional ByVal FPPosition As Integer = 1) As ProgressIndicator
     ' ������ �������� ���������, � ���������� ���
     On Error Resume Next
     Set AddChildIndicator = New ProgressIndicator
     Set AddChildIndicator.Parent = Me
     AddChildIndicator.Show Caption, FPPosition
     Children.Add AddChildIndicator
 End Function

 Private Sub Class_Initialize()
     ' ��������� �� ��������� ��� ����� ������������ ����������
     Set FP = New F_Progress: ShowPercents = True: FPVisible = True
     PrS = 0: PrE = 100: Set_ProgressBar 0: FP.PrBar.Caption = ""
     FPStartTime = Now: ShowTime = True: ShowPercents = True
     Set FP.indicator = Me
 End Sub

 Sub Show(ByVal Caption As String, Optional ByVal FPPosition As Integer = 0, _
          Optional LogSize As Long = 0)
     ' ���������� ��������-���
     On Error Resume Next
     SetProgressFormCaption Caption: On Error Resume Next:
     FP.PrBar.Width = ProgressBar_Default_Width
     Position = FPPosition
     FP.Tag = Caption: FP.Show:
     If Position <> 0 Then Move Position
     FP.Repaint: DoEvents
     SetLogSize LogSize
 End Sub

 Sub Hide(): Unload FP: FPVisible = False: End Sub    ' ������� ��������-����

 Sub Repaint()
     FP.Repaint: DoEvents
     If Application.ScreenUpdating = False Then
         Application.ScreenUpdating = True: Application.ScreenUpdating = False
     End If
 End Sub

 Sub Move(ByVal Position As Integer)    ' ������� ��������-���� �� ������ �� ���������
     If Abs(Position) > 3 Then Exit Sub
     h = FP.Height
     If Not Me.Parent Is Nothing Then h = Me.Parent.FP.Height
     FP.Top = FP.Top + (h + 3) * Position
 End Sub

 Public Property Get Visible(): Visible = FPVisible: End Property

 ' ��������� ��������� ����� � �������� �� ����������
 Public Property Let line1(ByVal NewValue As String): FP.L1.Caption = ProcessLabel(NewValue): Me.FP.Repaint: DoEvents: End Property
 Public Property Let line2(ByVal NewValue As String): FP.L2.Caption = ProcessLabel(NewValue): Me.FP.Repaint: DoEvents: End Property
 Public Property Let line3(ByVal NewValue As String): FP.L3.Caption = ProcessLabel(NewValue): Me.FP.Repaint: DoEvents: End Property
 Public Property Let Caption(ByVal NewValue As String): SetProgressFormCaption NewValue: End Property

 Private Sub SetProgressFormCaption(Optional ByVal Caption As String = "")
     ' ������������� ��������� ����� ��������-����
     ' � ������ �������� (����������� ������� � �������� ����������)
     If Len(Caption) > 0 Then FP.Tag = Caption
     txt = Trim(FP.Tag): If ShowPercents Then txt = Fix(Percent) & " %   " & txt
     dt = Format(Now - FPStartTime, "HH:NN:SS")
     If ShowTime Then txt = "( " & dt & " )    " & txt
     FP.Caption = txt
 End Sub

 Private Function TimeToFinish() As String
     If Percent < 15 Then Exit Function    ' ������ ����������� �����, ����� �� ������ ����������...
     dt = (Now - FPStartTime) * (100 - Percent) / Percent
     TimeToFinish = IIf(Minute(dt) > 0, Minute(dt) & " ���. ", "") & Second(dt) & " ���."
     If dt < TimeSerial(0, 0, 1) Then TimeToFinish = "����� �������"
     TimeToFinish = "�������� �� ����������: " & TimeToFinish
 End Function

 Sub SetFocus()    ' ������ ����� ��������-���� ��������
     FP.Show 0: If Position <> 0 Then Move Position
 End Sub

 Private Sub UpdateLabels(Optional ByVal L1_txt$, Optional ByVal L2_txt$, Optional ByVal L3_txt$)
     ' ��������� ������� �� ��������-���� (������� ������ �������� ������)
     If L1_txt$ <> "" Then FP.L1.Caption = ProcessLabel(L1_txt$)
     If L2_txt$ <> "" Or L1_txt$ <> "" Then FP.L2.Caption = ProcessLabel(L2_txt$)
     If L3_txt$ <> "" Or L2_txt$ <> "" Or L1_txt$ <> "" Then FP.L3.Caption = ProcessLabel(L3_txt$)
 End Sub

 Private Function ProcessLabel(ByVal txt As String) As String
     ' �������� �������� ����� � ������ txt �� �������� ���������� ����������
     txt = Replace(txt, "$index", SubActionIndex)
     txt = Replace(txt, "$count", SubActionsCount)
     txt = Replace(txt, "$time", TimeToFinish)
     ProcessLabel = txt
 End Function

 Sub SubAction(Optional ByVal L1_txt$ = "", Optional ByVal L2_txt$ = "", Optional ByVal L3_txt$ = "")
     ' ������ �������� ������ ��������� ��������
     On Error Resume Next
     If SubActionsCount = 0 Then SubActionsCount = 1
     SubActionIndex = SubActionIndex + 1
     If SubActionIndex > SubActionsCount Then SubActionIndex = SubActionsCount
     Percent = PrS + (PrE - PrS) * ((SubActionIndex - 1) / SubActionsCount)
     UpdateLabels L1_txt$, L2_txt$, L3_txt$
     Set_ProgressBar Percent: DoEvents
 End Sub

 Sub StartNewAction(Optional ByVal Pr_Start As Integer = 0, Optional ByVal Pr_End As Integer = 100, _
                    Optional ByVal L1_txt$ = "", Optional ByVal L2_txt$ = " ", Optional ByVal L3_txt$ = " ", _
                    Optional ByVal Actions_Count As Long = 0)
     ' ������ ������ �������� - �� ����� ���������� �� Pr_Start % �� Pr_End %
     ' � ���������� Actions_Count �������� ���������� �������� ������ ��������
     On Error Resume Next
     PrS = Pr_Start: PrE = Pr_End: SubActionIndex = 0: SubActionsCount = Actions_Count
     UpdateLabels L1_txt$, L2_txt$, L3_txt$
     Set_ProgressBar PrS
 End Sub

 Sub UpdateFromChild(ByVal ChildPercent As Double)
     ' ����������� ��������� �� ������������ ����� ��� ���������� �� ��������
     If SubActionsCount = 0 Then
         Percent = PrS + (PrE - PrS) * (ChildPercent / 100)
     Else
         ' ��������, SubActionIndex = 3: SubActionsCount=10: PrS = 50: PrE = 100: ChildPercent=40
         ' ��������� �.�. � ��������� �� 60 �� 65, � ������ ����� 62
         Percent = PrS + (PrE - PrS) / SubActionsCount * (SubActionIndex - 1) + _
                   (PrE - PrS) / SubActionsCount * (ChildPercent / 100)
     End If
     Set_ProgressBar Percent
 End Sub

 Private Sub Set_ProgressBar(ByVal NewPercent As Double)
     ' ��������� ������ ����������
     Percent = NewPercent
     If NewPercent > 100 Then Percent = 100
     If NewPercent < 0 Then Percent = 0
     FP.PrBar.Width = Int(Percent * ProgressBar_Default_Width / 100)
     SetProgressFormCaption
     FP.Repaint
     If Not Parent Is Nothing Then Parent.UpdateFromChild Percent
 End Sub

 'Private Function GetCurrentProgress() As Long    ' ���������� ������� ������� ����������
 '    If FP.PrBar.Width = 0 Then Exit Function
 '    GetCurrentProgress = FP.PrBar.Width / ProgressBar_Default_Width * 100
 'End Function

 Private Sub Class_Terminate()    ' ����������� ���������� ������
     On Error Resume Next
     Unload FP: FPVisible = False
 End Sub

 Private Function ProgressBar_Default_Width() As Double    ' ��������� ������� ������ �� ������� �����
     ProgressBar_Default_Width = FP.Width - 18
 End Function

 Function CancelButton() As MSForms.CommandButton
     Set CancelButton = FP.CommandButton_stop
 End Function

 ' ============================== ���������� �� 23.02.2012 =========================================
 Sub SetLogSize(ByVal n As Long)
     On Error Resume Next
     If n < 0 Then n = 0
     If n > 5 Then n = 5
     FP.SpinButton_log.value = n
     FP.SpinButton_log.Visible = n > 0
 End Sub

 Sub Log(ByVal txt$)
     On Error Resume Next
     If ShowTimeInLog Then currtime$ = Time & vbTab
     LogString = LogString & vbNewLine & currtime$ & txt
     FP.TextBox_Log.Text = Mid(LogString, 3)
     If FP.SpinButton_log.value = 0 Then FP.SpinButton_log.value = 2: FP.SpinButton_log.Visible = True
     FP.CommandButton_stop.SetFocus: FP.TextBox_Log.SetFocus
 End Sub

 Sub ClearLog()
     LogString = "": FP.TextBox_Log.Text = ""
 End Sub

 Sub ShowLog()
     On Error Resume Next: Err.Clear
     Filename$ = Environ("TEMP") & "\macro_log.txt"    ' ��������� ��� ��� ���������� �����
     With CreateObject("scripting.filesystemobject").CreateTextFile(Filename, True)
         .Write Mid(LogString, 3): .Close    ' ��������� ����� � ����
     End With
     CreateObject("wscript.shell").Run """" & Filename$ & """"    ' ��������� ��������� ����
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

 ' ============================== ���������� �� 07.10.2012 =========================================

 Sub AddButton(ByVal Caption$, ByVal Macro$)    ' ���������� ������ ������� �������
     dd = 18
     If FP.SpinButton_log = 0 Then FP.SpinButton_log = 1
     With Me.FP.CommandButton_RunMacro
         .Caption = Caption$
         .Visible = True
         .Top = FP.Height - .Height - dd - 20
         .Left = FP.Width - .Width - dd - 15
     End With
     FP.ButtonMacro = Macro$
 End Sub

 Function MacroButton() As MSForms.CommandButton
     Set MacroButton = FP.CommandButton_RunMacro
 End Function

 Sub QueryClose()    ' ���������� �� �����, ��� ������� � ��������
     On Error Resume Next
     Dim pi As ProgressIndicator
     For Each pi In Children
         pi.QueryClose
         pi.Hide
     Next pi
 End Sub

End Sub
