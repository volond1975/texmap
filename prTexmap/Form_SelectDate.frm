VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_SelectDate 
   OleObjectBlob   =   "Form_SelectDate.frx":0000
   Caption         =   "����� ���� � �������"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   209
End
Attribute VB_Name = "Form_SelectDate"
Attribute VB_Base = "0{2B54A5D6-F573-4D1D-ABC7-29D4262762AF}{2A4F567B-B169-4048-A3F2-EF700B3FAA5A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False




Dim mozno As Boolean


Private Sub chb_Time_Click()
    Me.Frame_time.Visible = Me.chb_Time.value
End Sub
Private Sub UserForm_Initialize()
    Form_DateTime_Showed = True
    Me.chb_Time.value = False    ' �� ��������� ����� �� ����������
    Me.Frame_time.Visible = False
    '���������� �������� ���� ��� ������ ������
    If IsDate(SelectedDate) Then
        dt_1 = CDate(SelectedDate)
    Else
        If IsDate(DefaultDate) Then dt_1 = CDate(DefaultDate) Else dt_1 = Now
    End If
    dt_2 = dt_1

    '���������� ������ ComboBox_Month
    ComboBox_Month.AddItem "������"
    ComboBox_Month.AddItem "�������"
    ComboBox_Month.AddItem "����"
    ComboBox_Month.AddItem "������"
    ComboBox_Month.AddItem "���"
    ComboBox_Month.AddItem "����"
    ComboBox_Month.AddItem "����"
    ComboBox_Month.AddItem "������"
    ComboBox_Month.AddItem "��������"
    ComboBox_Month.AddItem "�������"
    ComboBox_Month.AddItem "������"
    ComboBox_Month.AddItem "�������"

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False



    Dim TimeNow_i As String
    '���������� ������ ComboBox_Hour
    For i = 0 To 23
        If Len(str(i)) = 2 Then TimeNow_i = "0" + Right(str(i), 1) Else TimeNow_i = Right(str(i), 2)
        ComboBox_Hour.AddItem TimeNow_i
    Next i
    ComboBox_Hour.AddItem "00"

    '���������� ������� ComboBox_Minute � ComboBox_Second
    For i = 0 To 59
        If Len(str(i)) = 2 Then TimeNow_i = "0" + Right(str(i), 1) Else TimeNow_i = Right(str(i), 2)
        ComboBox_Minute.AddItem TimeNow_i
        ComboBox_Second.AddItem TimeNow_i
    Next i
    ComboBox_Minute.AddItem "00"
    ComboBox_Second.AddItem "00"


    '��������� Label_Hour_Minute_Second
    Set_Label_Hour_Minute_Second (dt_1)
    mozno = True
    '��������� ComboBox_Hour
    Set_ComboBox_Hour (dt_1)
    '��������� ScrollBar_Hour
    Set_ScrollBar_Hour (dt_1)
    '��������� ComboBox_Minute
    Set_ComboBox_Minute (dt_1)
    '��������� ScrollBar_Minute
    Set_ScrollBar_Minute (dt_1)
    '��������� ComboBox_Second
    Set_ComboBox_Second (dt_1)
    mozno = False

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '�������������� �������� ���� ��� ������ ������
    If CloseMode = 0 Then Cancel = 1
    Form_DateTime_Showed = False
End Sub

Private Sub Cmd_Cancel_Click()
    '������� - �������� ����� ���� � ������� �����
    If IsDate(SelectedDate) Then SelectedDate = CStr(DateValue(dt_2))
    Unload Me
End Sub
Private Sub Cmd_Select_Click()
    '������� - ��������� ����� ���� � ������� �����
    SelectedDate = CStr(DateValue(dt_1))
   Form_SelectDate.Hide
   'Unload Me
End Sub


Private Sub Set_TextBox_����(MyDate As Date)
    '��������� TextBox_����
    TextBox_����.value = Format(MyDate, "dd.mm.yyyy")
End Sub

Private Sub Set_TextBox_Year(MyDate As Date)
    '��������� TextBox_Year
    TextBox_Year.value = Format(MyDate, "yyyy")
End Sub

Private Sub Set_M�nth(MyDate As Date)
    '��������� ComboBox_Month � ���������
    MyYear = Year(MyDate)
    MyMonth = Month(MyDate)
    MyDay = Day(MyDate)

    Label_Day.Caption = MyDay
    Label_Month.Caption = MyMonth
    Label_Year.Caption = MyYear

    ComboBox_Month.ListIndex = MyMonth - 1

    MyWeekDay = Weekday(DateSerial(MyYear, MyMonth, 1), vbMonday)
    MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)

    l_start = 2 - MyWeekDay
    For i = 1 To 6
        For j = 1 To 7

            If l_start >= 1 And l_start <= MyCountDay Then
                Me.Controls("Cell_" & i & "_" & j).Caption = l_start
            Else
                Me.Controls("Cell_" & i & "_" & j).Caption = ""
            End If

            If l_start = MyDay Then
                Set_On_Off CInt(i), CInt(j)
            End If

            l_start = l_start + 1

        Next j, i

        Cmd_Select.SetFocus
    End Sub


Private Sub Set_Label_Hour_Minute_Second(MyDate As Date)
    '��������� ������� Label_Hour_Minute_Second
    Label_Hour.Caption = Format(MyDate, "hh")
    Label_Minute.Caption = Format(MyDate, "nn")
    Label_Second.Caption = Format(MyDate, "ss")
End Sub

Private Sub Set_ComboBox_Hour(MyDate As Date)
    '��������� ComboBox_Hour
    MyHour = Hour(MyDate)
    ComboBox_Hour.ListIndex = MyHour
End Sub

Private Sub Set_ScrollBar_Hour(MyDate As Date)
    '��������� ScrollBar_Hour
    MyHour = Hour(MyDate)
    ScrollBar_Hour.value = MyHour
End Sub

Private Sub Set_ComboBox_Minute(MyDate As Date)
    '��������� ComboBox_Minute
    MyMinute = Minute(MyDate)
    ComboBox_Minute.ListIndex = MyMinute
End Sub

Private Sub Set_ScrollBar_Minute(MyDate As Date)
    '��������� ScrollBar_Minute
    MyMinute = Minute(MyDate)
    ScrollBar_Minute.value = MyMinute
End Sub

Private Sub Set_ComboBox_Second(MyDate As Date)
    '��������� ComboBox_Second
    MySecond = Second(MyDate)
    ComboBox_Second.ListIndex = MySecond
End Sub






Private Sub Cmd_�������_����_Click()
    '������� - ���������� ����, ��������������� �������� ���
    'MyYear = Year(dt_1)
    'MyMonth = Month(dt_1)
    'MyDay = Day(dt_1)
    MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    dt_1 = Date + TimeSerial(MyHour, MyMinute, MySecond)

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False
End Sub

Private Sub Cmd_�����_����_Click()
    '������� - ���������� ����, �� ���� �����
    dt_1 = dt_1 - 1

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False
End Sub

Private Sub Cmd_������_����_Click()
    '������� - ���������� ����, �� ���� �����
    dt_1 = dt_1 + 1

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False
End Sub


Private Sub Set_����(iRow As Integer, jCol As Integer)
    '������� - ���������� ����, ��������� � ���������

    If Me.Controls("Cell_" & iRow & "_" & jCol).Caption = "" Then Exit Sub

    MyYear = Year(dt_1)
    MyMonth = Month(dt_1)
    MyDay = CInt(Me.Controls("Cell_" & iRow & "_" & jCol).Caption)
    MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(MyHour, MyMinute, MySecond)

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False
End Sub


Private Sub ComboBox_Month_Change()
    If mozno Then Exit Sub

    '������� - ���������� ����, ��������� � ��������� (����� ������)

    MyYear = Year(dt_1)
    MyMonth = CInt(ComboBox_Month.ListIndex + 1)
    MyDay = Day(dt_1)
    MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)
    If MyDay > MyCountDay Then MyDay = MyCountDay

    MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(MyHour, MyMinute, MySecond)

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False
End Sub



Private Sub SpinButton_Year_SpinDown()
    '������� - ���������� ����, ��������� � ��������� (����� ���� -1)

    MyYear = Year(dt_1) - 1
    MyMonth = Month(dt_1)
    MyDay = Day(dt_1)
    MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)
    If MyDay > MyCountDay Then MyDay = MyCountDay

    MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(MyHour, MyMinute, MySecond)

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False
End Sub

Private Sub SpinButton_Year_SpinUp()
    '������� - ���������� ����, ��������� � ��������� (����� ���� +1)

    MyYear = Year(dt_1) + 1
    MyMonth = Month(dt_1)
    MyDay = Day(dt_1)
    MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)
    If MyDay > MyCountDay Then MyDay = MyCountDay

    MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(MyHour, MyMinute, MySecond)

    '��������� TextBox_����
    Set_TextBox_���� (dt_1)
    '��������� TextBox_Year
    Set_TextBox_Year (dt_1)
    '��������� ComboBox_Month � ���������
    mozno = True
    Set_M�nth (dt_1)
    mozno = False
End Sub

Private Sub Label_Hour_Click()
    ComboBox_Hour.DropDown
    ComboBox_Hour.SetFocus
End Sub
Private Sub Label_Minute_Click()
    ComboBox_Minute.DropDown
    ComboBox_Minute.SetFocus
End Sub
Private Sub Label_Second_Click()
    ComboBox_Second.DropDown
    ComboBox_Second.SetFocus
End Sub

Private Sub ComboBox_Hour_Change()
    If mozno Then Exit Sub
    '������� - ���������� �����, ����� �������� ����
    MyYear = Year(dt_1)
    MyMonth = Month(dt_1)
    MyDay = Day(dt_1)
    'MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(ComboBox_Hour.value, MyMinute, MySecond)

    mozno = True
    '��������� Label_Hour_Minute_Second
    Set_Label_Hour_Minute_Second (dt_1)
    '��������� ComboBox_Hour
    Set_ComboBox_Hour (dt_1)
    '��������� ScrollBar_Hour
    Set_ScrollBar_Hour (dt_1)
    '��������� ComboBox_Minute
    Set_ComboBox_Minute (dt_1)
    '��������� ScrollBar_Minute
    Set_ScrollBar_Minute (dt_1)
    '��������� ComboBox_Second
    Set_ComboBox_Second (dt_1)
    mozno = False
End Sub

Private Sub ComboBox_Minute_Change()
    If mozno Then Exit Sub
    '������� - ���������� �����, ����� �������� �����
    MyYear = Year(dt_1)
    MyMonth = Month(dt_1)
    MyDay = Day(dt_1)
    MyHour = Hour(dt_1)
    'MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(MyHour, ComboBox_Minute.value, MySecond)

    mozno = True
    '��������� Label_Hour_Minute_Second
    Set_Label_Hour_Minute_Second (dt_1)
    '��������� ComboBox_Hour
    Set_ComboBox_Hour (dt_1)
    '��������� ScrollBar_Hour
    Set_ScrollBar_Hour (dt_1)
    '��������� ComboBox_Minute
    Set_ComboBox_Minute (dt_1)
    '��������� ScrollBar_Minute
    Set_ScrollBar_Minute (dt_1)
    '��������� ComboBox_Second
    Set_ComboBox_Second (dt_1)
    mozno = False
End Sub

Private Sub ComboBox_Second_Change()
    If mozno Then Exit Sub
    '������� - ���������� �����, ����� �������� ������
    MyYear = Year(dt_1)
    MyMonth = Month(dt_1)
    MyDay = Day(dt_1)
    MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    'MySecond = Second(dt_1)

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(MyHour, MyMinute, ComboBox_Second.value)

    mozno = True
    '��������� Label_Hour_Minute_Second
    Set_Label_Hour_Minute_Second (dt_1)
    '��������� ComboBox_Hour
    Set_ComboBox_Hour (dt_1)
    '��������� ScrollBar_Hour
    Set_ScrollBar_Hour (dt_1)
    '��������� ComboBox_Minute
    Set_ComboBox_Minute (dt_1)
    '��������� ScrollBar_Minute
    Set_ScrollBar_Minute (dt_1)
    '��������� ComboBox_Second
    Set_ComboBox_Second (dt_1)
    mozno = False
End Sub

Private Sub ScrollBar_Hour_Change()
    If mozno Then Exit Sub
    '������� - ���������� �����, ����� �������� ����
    MyYear = Year(dt_1)
    MyMonth = Month(dt_1)
    MyDay = Day(dt_1)
    'MyHour = Hour(dt_1)
    MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    ' ������ ����� �������
    If ScrollBar_Hour.value = 24 Then
        mozno = True
        ScrollBar_Hour.value = 0
        mozno = False
    Else
        If ScrollBar_Hour.value = -1 Then
            mozno = True
            ScrollBar_Hour.value = 23
            mozno = False
        End If
    End If

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(ScrollBar_Hour.value, MyMinute, MySecond)

    mozno = True
    '��������� Label_Hour_Minute_Second
    Set_Label_Hour_Minute_Second (dt_1)
    '��������� ComboBox_Hour
    Set_ComboBox_Hour (dt_1)
    '��������� ScrollBar_Hour
    Set_ScrollBar_Hour (dt_1)
    '��������� ComboBox_Minute
    Set_ComboBox_Minute (dt_1)
    '��������� ScrollBar_Minute
    Set_ScrollBar_Minute (dt_1)
    '��������� ComboBox_Second
    Set_ComboBox_Second (dt_1)
    mozno = False
End Sub
Private Sub ScrollBar_Minute_Change()
    If mozno Then Exit Sub
    '������� - ���������� �����, ����� �������� �����
    MyYear = Year(dt_1)
    MyMonth = Month(dt_1)
    MyDay = Day(dt_1)
    MyHour = Hour(dt_1)
    'MyMinute = Minute(dt_1)
    MySecond = Second(dt_1)

    ' ������ ����� �������
    If ScrollBar_Minute.value = 60 Then
        mozno = True
        ScrollBar_Minute.value = 0
        mozno = False
    Else
        If ScrollBar_Minute.value = -1 Then
            mozno = True
            ScrollBar_Minute.value = 59
            mozno = False
        End If
    End If

    dt_1 = DateSerial(MyYear, MyMonth, MyDay) + TimeSerial(MyHour, ScrollBar_Minute.value, MySecond)

    mozno = True
    '��������� Label_Hour_Minute_Second
    Set_Label_Hour_Minute_Second (dt_1)
    '��������� ComboBox_Hour
    Set_ComboBox_Hour (dt_1)
    '��������� ScrollBar_Hour
    Set_ScrollBar_Hour (dt_1)
    '��������� ComboBox_Minute
    Set_ComboBox_Minute (dt_1)
    '��������� ScrollBar_Minute
    Set_ScrollBar_Minute (dt_1)
    '��������� ComboBox_Second
    Set_ComboBox_Second (dt_1)
    mozno = False
End Sub










Private Sub Cell_1_1_Click()
    Set_���� 1, 1
End Sub
Private Sub Cell_1_2_Click()
    Set_���� 1, 2
End Sub
Private Sub Cell_1_3_Click()
    Set_���� 1, 3
End Sub
Private Sub Cell_1_4_Click()
    Set_���� 1, 4
End Sub
Private Sub Cell_1_5_Click()
    Set_���� 1, 5
End Sub
Private Sub Cell_1_6_Click()
    Set_���� 1, 6
End Sub
Private Sub Cell_1_7_Click()
    Set_���� 1, 7
End Sub
Private Sub Cell_2_1_Click()
    Set_���� 2, 1
End Sub
Private Sub Cell_2_2_Click()
    Set_���� 2, 2
End Sub
Private Sub Cell_2_3_Click()
    Set_���� 2, 3
End Sub
Private Sub Cell_2_4_Click()
    Set_���� 2, 4
End Sub
Private Sub Cell_2_5_Click()
    Set_���� 2, 5
End Sub
Private Sub Cell_2_6_Click()
    Set_���� 2, 6
End Sub
Private Sub Cell_2_7_Click()
    Set_���� 2, 7
End Sub
Private Sub Cell_3_1_Click()
    Set_���� 3, 1
End Sub
Private Sub Cell_3_2_Click()
    Set_���� 3, 2
End Sub
Private Sub Cell_3_3_Click()
    Set_���� 3, 3
End Sub
Private Sub Cell_3_4_Click()
    Set_���� 3, 4
End Sub
Private Sub Cell_3_5_Click()
    Set_���� 3, 5
End Sub
Private Sub Cell_3_6_Click()
    Set_���� 3, 6
End Sub
Private Sub Cell_3_7_Click()
    Set_���� 3, 7
End Sub
Private Sub Cell_4_1_Click()
    Set_���� 4, 1
End Sub
Private Sub Cell_4_2_Click()
    Set_���� 4, 2
End Sub
Private Sub Cell_4_3_Click()
    Set_���� 4, 3
End Sub
Private Sub Cell_4_4_Click()
    Set_���� 4, 4
End Sub
Private Sub Cell_4_5_Click()
    Set_���� 4, 5
End Sub
Private Sub Cell_4_6_Click()
    Set_���� 4, 6
End Sub
Private Sub Cell_4_7_Click()
    Set_���� 4, 7
End Sub
Private Sub Cell_5_1_Click()
    Set_���� 5, 1
End Sub
Private Sub Cell_5_2_Click()
    Set_���� 5, 2
End Sub
Private Sub Cell_5_3_Click()
    Set_���� 5, 3
End Sub
Private Sub Cell_5_4_Click()
    Set_���� 5, 4
End Sub
Private Sub Cell_5_5_Click()
    Set_���� 5, 5
End Sub
Private Sub Cell_5_6_Click()
    Set_���� 5, 6
End Sub
Private Sub Cell_5_7_Click()
    Set_���� 5, 7
End Sub
Private Sub Cell_6_1_Click()
    Set_���� 6, 1
End Sub
Private Sub Cell_6_2_Click()
    Set_���� 6, 2
End Sub
Private Sub Cell_6_3_Click()
    Set_���� 6, 3
End Sub
Private Sub Cell_6_4_Click()
    Set_���� 6, 4
End Sub
Private Sub Cell_6_5_Click()
    Set_���� 6, 5
End Sub
Private Sub Cell_6_6_Click()
    Set_���� 6, 6
End Sub
Private Sub Cell_6_7_Click()
    Set_���� 6, 7
End Sub

Private Sub Set_On_Off(iRow As Integer, jCol As Integer)

    If Me.Controls("Cell_" & iRow & "_" & jCol).Caption = "" Then Exit Sub

    ' �������� ��� ������
    For i = 1 To 6
        For j = 1 To 7
            Me.Controls("Cell_" & i & "_" & j).BackColor = RGB(255, 255, 255)
            Me.Controls("Cell_" & i & "_" & j).BorderColor = RGB(255, 255, 255)
        Next j
    Next i

    ' �������� ������� ������
    Me.Controls("Cell_" & iRow & "_" & jCol).BackColor = RGB(204, 255, 204)
    Me.Controls("Cell_" & iRow & "_" & jCol).BorderColor = RGB(150, 150, 150)

End Sub
