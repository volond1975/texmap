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
 ' Автор     : EducatedFool  (Игорь)                    Дата: 07.10.2012
 ' Разработка макросов любой сложности для Microsoft Excel
 ' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
 ' Реквизиты для оплаты работы: http://ExcelVBA.ru/payments
 '---------------------------------------------------------------------------------------

 Public FP As New F_Progress

 Public SubActionIndex As Single, SubActionsCount As Single
 Private FPVisible As Boolean, FPStartTime As Date, Position As Integer
 Private PrS As Integer, PrE As Integer, Percent As Double, LogString As String

 Public Parent As ProgressIndicator
 Public ShowPercents As Boolean, ShowTime As Boolean, ShowTimeInLog As Boolean
 Public Children As New Collection

 Function AddChildIndicator(ByVal Caption As String, Optional ByVal FPPosition As Integer = 1) As ProgressIndicator
     ' создаёт дочерний индикатор, и отображает его
     On Error Resume Next
     Set AddChildIndicator = New ProgressIndicator
     Set AddChildIndicator.Parent = Me
     AddChildIndicator.Show Caption, FPPosition
     Children.Add AddChildIndicator
 End Function

 Private Sub Class_Initialize()
     ' параметры по умолчанию для вновь создаваемого индикатора
     Set FP = New F_Progress: ShowPercents = True: FPVisible = True
     PrS = 0: PrE = 100: Set_ProgressBar 0: FP.PrBar.Caption = ""
     FPStartTime = Now: ShowTime = True: ShowPercents = True
     Set FP.indicator = Me
 End Sub

 Sub Show(ByVal Caption As String, Optional ByVal FPPosition As Integer = 0, _
          Optional LogSize As Long = 0)
     ' отображает прогресс-бар
     On Error Resume Next
     SetProgressFormCaption Caption: On Error Resume Next:
     FP.PrBar.Width = ProgressBar_Default_Width
     Position = FPPosition
     FP.Tag = Caption: FP.Show:
     If Position <> 0 Then Move Position
     FP.Repaint: DoEvents
     SetLogSize LogSize
 End Sub

 Sub Hide(): Unload FP: FPVisible = False: End Sub    ' скрытие прогресс-бара

 Sub Repaint()
     FP.Repaint: DoEvents
     If Application.ScreenUpdating = False Then
         Application.ScreenUpdating = True: Application.ScreenUpdating = False
     End If
 End Sub

 Sub Move(ByVal Position As Integer)    ' позиция прогресс-бара на экране по вертикали
     If Abs(Position) > 3 Then Exit Sub
     h = FP.Height
     If Not Me.Parent Is Nothing Then h = Me.Parent.FP.Height
     FP.Top = FP.Top + (h + 3) * Position
 End Sub

 Public Property Get Visible(): Visible = FPVisible: End Property

 ' установка заголовка формы и надписей на индикаторе
 Public Property Let line1(ByVal NewValue As String): FP.L1.Caption = ProcessLabel(NewValue): Me.FP.Repaint: DoEvents: End Property
 Public Property Let line2(ByVal NewValue As String): FP.L2.Caption = ProcessLabel(NewValue): Me.FP.Repaint: DoEvents: End Property
 Public Property Let line3(ByVal NewValue As String): FP.L3.Caption = ProcessLabel(NewValue): Me.FP.Repaint: DoEvents: End Property
 Public Property Let Caption(ByVal NewValue As String): SetProgressFormCaption NewValue: End Property

 Private Sub SetProgressFormCaption(Optional ByVal Caption As String = "")
     ' устанавливает заголовок формы прогресс-бара
     ' с учёток настроек (отображение таймера и процента выполнения)
     If Len(Caption) > 0 Then FP.Tag = Caption
     txt = Trim(FP.Tag): If ShowPercents Then txt = Fix(Percent) & " %   " & txt
     dt = Format(Now - FPStartTime, "HH:NN:SS")
     If ShowTime Then txt = "( " & dt & " )    " & txt
     FP.Caption = txt
 End Sub

 Private Function TimeToFinish() As String
     If Percent < 15 Then Exit Function    ' сложно предсказать время, когда всё только начинается...
     dt = (Now - FPStartTime) * (100 - Percent) / Percent
     TimeToFinish = IIf(Minute(dt) > 0, Minute(dt) & " мин. ", "") & Second(dt) & " сек."
     If dt < TimeSerial(0, 0, 1) Then TimeToFinish = "менее секунды"
     TimeToFinish = "Осталось до завершения: " & TimeToFinish
 End Function

 Sub SetFocus()    ' делает форму прогресс-бара активной
     FP.Show 0: If Position <> 0 Then Move Position
 End Sub

 Private Sub UpdateLabels(Optional ByVal L1_txt$, Optional ByVal L2_txt$, Optional ByVal L3_txt$)
     ' обновляем надписи на прогресс-баре (выводит только непустые строки)
     If L1_txt$ <> "" Then FP.L1.Caption = ProcessLabel(L1_txt$)
     If L2_txt$ <> "" Or L1_txt$ <> "" Then FP.L2.Caption = ProcessLabel(L2_txt$)
     If L3_txt$ <> "" Or L2_txt$ <> "" Or L1_txt$ <> "" Then FP.L3.Caption = ProcessLabel(L3_txt$)
 End Sub

 Private Function ProcessLabel(ByVal txt As String) As String
     ' заменяет ключевые слова в строке txt на значения параметров индикатора
     txt = Replace(txt, "$index", SubActionIndex)
     txt = Replace(txt, "$count", SubActionsCount)
     txt = Replace(txt, "$time", TimeToFinish)
     ProcessLabel = txt
 End Function

 Sub SubAction(Optional ByVal L1_txt$ = "", Optional ByVal L2_txt$ = "", Optional ByVal L3_txt$ = "")
     ' запуск операции внутри основного действия
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
     ' запуск нового действия - на шкале индикатора от Pr_Start % до Pr_End %
     ' в переменной Actions_Count получает количество операций внутри действия
     On Error Resume Next
     PrS = Pr_Start: PrE = Pr_End: SubActionIndex = 0: SubActionsCount = Actions_Count
     UpdateLabels L1_txt$, L2_txt$, L3_txt$
     Set_ProgressBar PrS
 End Sub

 Sub UpdateFromChild(ByVal ChildPercent As Double)
     ' отображение изменений на родительской форме при изменениях на дочерней
     If SubActionsCount = 0 Then
         Percent = PrS + (PrE - PrS) * (ChildPercent / 100)
     Else
         ' например, SubActionIndex = 3: SubActionsCount=10: PrS = 50: PrE = 100: ChildPercent=40
         ' результат д.б. в диапазоне от 60 до 65, а именно равен 62
         Percent = PrS + (PrE - PrS) / SubActionsCount * (SubActionIndex - 1) + _
                   (PrE - PrS) / SubActionsCount * (ChildPercent / 100)
     End If
     Set_ProgressBar Percent
 End Sub

 Private Sub Set_ProgressBar(ByVal NewPercent As Double)
     ' изменение ширины индикатора
     Percent = NewPercent
     If NewPercent > 100 Then Percent = 100
     If NewPercent < 0 Then Percent = 0
     FP.PrBar.Width = Int(Percent * ProgressBar_Default_Width / 100)
     SetProgressFormCaption
     FP.Repaint
     If Not Parent Is Nothing Then Parent.UpdateFromChild Percent
 End Sub

 'Private Function GetCurrentProgress() As Long    ' возвращает текущий процент выполнения
 '    If FP.PrBar.Width = 0 Then Exit Function
 '    GetCurrentProgress = FP.PrBar.Width / ProgressBar_Default_Width * 100
 'End Function

 Private Sub Class_Terminate()    ' уничтожение экземпляра класса
     On Error Resume Next
     Unload FP: FPVisible = False
 End Sub

 Private Function ProgressBar_Default_Width() As Double    ' установка размера полосы по размеру формы
     ProgressBar_Default_Width = FP.Width - 18
 End Function

 Function CancelButton() As MSForms.CommandButton
     Set CancelButton = FP.CommandButton_stop
 End Function

 ' ============================== обновление от 23.02.2012 =========================================
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
     Filename$ = Environ("TEMP") & "\macro_log.txt"    ' формируем имя для временного файла
     With CreateObject("scripting.filesystemobject").CreateTextFile(Filename, True)
         .Write Mid(LogString, 3): .Close    ' сохраняем текст в файл
     End With
     CreateObject("wscript.shell").Run """" & Filename$ & """"    ' открываем созданный файл
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

 ' ============================== обновление от 07.10.2012 =========================================

 Sub AddButton(ByVal Caption$, ByVal Macro$)    ' добавление кнопки запуска макроса
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

 Sub QueryClose()    ' вызывается из формы, при попытке её закрытия
     On Error Resume Next
     Dim pi As ProgressIndicator
     For Each pi In Children
         pi.QueryClose
         pi.Hide
     Next pi
 End Sub

End Sub
