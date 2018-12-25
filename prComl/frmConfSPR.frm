VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfSPR 
   OleObjectBlob   =   "frmConfSPR.frx":0000
   Caption         =   "Конфигуратор справочников"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10545
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   436
End
Attribute VB_Name = "frmConfSPR"
Attribute VB_Base = "0{B4EE159C-6537-4970-AB68-358369A231F9}{9CA949CD-1BE7-4E9E-ACB6-C5A4ADE812CD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub FiltrDataDo_Click()
Dim I As Integer
Dim Criterial()
Dim A
Dim b

Dim CriterialFiltrTextFiltr()
Dim CriterialFiltrDataFiltrFrom()
Dim CriterialFiltrFiltrTo()

Dim TextFiltr()
Dim DataFiltrFrom() As String
Dim DataFiltrTo() As String
Dim CritIndex

frmConfSPR.ProgressProc.Visible = True
frmConfSPR.ProgressLabel.Caption = "Установка фильтров "
DoEvents

' Создаём переменные и присваиваем им значения
    For I = 2 To ColumnsTable - 1 Step 1
    Progress = I / (ColumnsTable - 1) * 0.999 * 100
    frmConfSPR.ProgressProc = Round(Progress, 0)
    
    
                    ReDim Preserve TextFiltr(I)
                    ReDim Preserve CriterialFiltrTextFiltr(I)
                    ReDim Preserve CriterialFiltrDataFiltrFrom(I)
                    ReDim Preserve CriterialFiltrFiltrTo(I)
                    ReDim Preserve DataFiltrFrom(I)
                    ReDim Preserve DataFiltrTo(I)
                    CritIndex = frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I - 1).ListIndex
        A = ""
        b = ""
        If ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value = "" Then
        ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value = 1
        
            If IsDate(ActiveSheet.ListObjects(NameTable).Range(3, I).value) Then
                'Фильтр даты
                ' Как фильтровать
                CriterialFiltrDataFiltrFrom(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 4).value
                CriterialFiltrFiltrTo(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 5).value
                'Значение фильтра
                    If frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I - 1).value <> "" Then
                        DataFiltrFrom(I) = frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I - 1).value
                    Else
                        DataFiltrFrom(I) = "01.01.1900"
                    End If
                    
                    If frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I - 1).value <> "" Then
                        DataFiltrTo(I) = frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I - 1).value
                    Else
                        DataFiltrTo(I) = Date
                    End If
            Else
                If IsNumeric(ActiveSheet.ListObjects(NameTable).Range(3, I).value) Then
                    'Фильтр числовых значений
                    ' Как фильтровать
                    CriterialFiltrTextFiltr(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 7, 3).value
                    'Значение фильтра
                    TextFiltr(I) = frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I - 1).value
                Else
                    'Фильтр текстовых значений
                    ' Как фильтровать
                    CriterialFiltrTextFiltr(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 3).value
                    A = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 4).value
                    b = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 5).value
                    'Значение фильтра
                    TextFiltr(I) = A & frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I - 1).value & b
                    A = ""
                    b = ""
                End If
            End If
        ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value = ""
        Else
            If IsDate(ActiveSheet.ListObjects(NameTable).Range(3, I).value) Then
                'Фильтр даты
                ' Как фильтровать
                CriterialFiltrDataFiltrFrom(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 4).value
                CriterialFiltrFiltrTo(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 5).value
                'Значение фильтра
                    If frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I - 1).value <> "" Then
                        DataFiltrFrom(I) = frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I - 1).value
                    Else
                        DataFiltrFrom(I) = "01.01.1900"
                    End If
                    
                    If frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I - 1).value <> "" Then
                        DataFiltrTo(I) = frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I - 1).value
                    Else
                        DataFiltrTo(I) = Date
                    End If
            Else
            If IsNumeric(ActiveSheet.ListObjects(NameTable).Range(3, I).value) Then
                    'Фильтр числовых значений
                    ' Как фильтровать)
                    CriterialFiltrTextFiltr(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 7, 3).value
                    'Значение фильтра
                    TextFiltr(I) = frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I - 1).value
                Else
                    'Фильтр текстовых значений
                    ' Как фильтровать
                    CriterialFiltrTextFiltr(I) = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 3).value
                    A = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 4).value
                    b = Sheets("Справочники").ListObjects("Критерии").Range(CritIndex + 3, 5).value
                    'Значение фильтра
                    TextFiltr(I) = A & frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I - 1).value & b
                    A = ""
                    b = ""
                End If
            End If
        End If
        
        'Ставлю новые фильтры
        If frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I - 1).value <> "" Then
            ' Текстовые поля и числа
            ActiveSheet.ListObjects(NameTable).Range.AutoFilter Field:=I, Criteria1:= _
            CriterialFiltrTextFiltr(I) & TextFiltr(I)
        Else
            If frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I - 1).value <> "" Or frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I - 1).value <> "" Then
                'Даты
                ActiveSheet.ListObjects(NameTable).Range.AutoFilter Field:=I, Criteria1:= _
                CriterialFiltrDataFiltrFrom(I) & Format(CStr(CDate(DataFiltrFrom(I))), "#"), Operator:=xlAnd, Criteria2:= _
                CriterialFiltrFiltrTo(I) & Format(CStr(CDate(DataFiltrTo(I))), "#")
            Else
                ActiveSheet.ListObjects(NameTable).Range.AutoFilter Field:=I
            End If
        End If
    Next I
    frmConfSPR.ProgressLabel.Caption = "Загрузка результатов "
    DoEvents
    frmMain.ListBoxSotrudnikiRefresh
    frmConfSPR.ProgressProc.Visible = False
    Unload Me
End Sub
' Очистка полей
Private Sub ClearAllData_Click()
Dim I As Integer
    For I = 1 To ColumnsTable - 1 Step 1
        frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I).value = ""
        frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I).value = ""
        frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I).value = ""
        Next I
End Sub
'Работа с календарём
Private Sub Calendar1_DblClick()
Dim Y1

    If Left(ButtonActive, 10) = "CalendFrom" Then
        Y1 = Right(ButtonActive, (Len(ButtonActive) - 10))
        TextBoxTo = "DateBoxFltrFrom" & Y1
    Else
        Y1 = Right(ButtonActive, (Len(ButtonActive) - 8))
        TextBoxTo = "DateBoxFltrTo" & Y1
    End If
    frmConfSPR.FrameFiltr.Controls(TextBoxTo).value = Me.Calendar1.Day & "." & Me.Calendar1.Month & "." & Me.Calendar1.Year
    frmConfSPR.FrameFiltr.Controls(TextBoxTo) = Format(frmConfSPR.FrameFiltr.Controls(TextBoxTo), "dd.mm.yyyy")
    frmConfSPR.FrameFiltr.Calendar1.Visible = False
End Sub
