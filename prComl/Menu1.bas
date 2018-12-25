Attribute VB_Name = "Menu1"
Sub Pusk()
frmMain.Show
End Sub

' Обработка формы
Sub CreateFiltr()
Dim I
Dim ObgTextBox As MSForms.TextBox, NamTextBox As String 'ЭУ, его имя
Dim ObgLabel As MSForms.Label, NamLabel As String
Dim ObgComboBox As MSForms.ComboBox, NamComboBox As String
Dim ObgCommandButton As MSForms.CommandButton, NamCommandButton As String
Dim cCBFiltrFrom() As New CBclass
Dim cCBFiltrTo() As New CBclass
Dim cTBFiltrFrom() As New DataClass
Dim cTBFiltrTo() As New DataClass
Dim cTB() As New TBclas
Dim cClear() As New ClearData
Dim FormatD()
Dim TopPoz
TopPoz = 6
On Error Resume Next
Application.GoTo Reference:=NameTable
If Err = 0 Then
Application.GoTo Reference:=NameTable
ColumnsTable = ActiveSheet.ListObjects(NameTable).ListColumns.Count
RowsTable = ActiveSheet.ListObjects(NameTable).ListRows.Count
For I = 1 To ColumnsTable - 1 Step 1
        
        'Lобавление Label (наименования полей = наименованиям столбцов)
        NamLabel = "LabelFltr" & I
        Set ObgLabel = frmConfSPR.FrameFiltr.Controls.Add("Forms.Label.1", NamLabel)
            ObgLabel.Height = 15.75
            ObgLabel.Left = 9
            ObgLabel.Top = TopPoz
            ObgLabel.Width = 160
            ObgLabel.Caption = ThisWorkbook.Sheets(frmMain.ComboBox_Лист.value).ListObjects(NameTable).ListColumns(I + 1)
                         
        ' Контрол выбора способа сортировки
        NamComboBox = "ComboBoxFltr" & I
        Set ObgComboBox = frmConfSPR.FrameFiltr.Controls.Add("Forms.ComboBox.1", NamComboBox)
            ObgComboBox.Height = 15.75
            ObgComboBox.Left = 170
            ObgComboBox.Top = TopPoz
            ObgComboBox.Width = 90
            ObgComboBox.Style = fmStyleDropDownList
            ObgComboBox.TabStop = False
            
       'добавление TextBox (начало интервала) для дат
        NamTextBox = "DateBoxFltrFrom" & I
        Set ObgTextBox = frmConfSPR.FrameFiltr.Controls.Add("Forms.TextBox.1", NamTextBox)
            ObgTextBox.Height = 15.75
            ObgTextBox.Left = 280
            ObgTextBox.Top = TopPoz
            ObgTextBox.Width = 60
            ObgTextBox.MaxLength = 10
            ObgTextBox = Format(ObgTextBox, "dd.mm.yyyy")
            ObgTextBox.Locked = True
        ReDim Preserve cTBFiltrFrom(I)
        Set cTBFiltrFrom(I).TCBData = frmConfSPR.FrameFiltr.Controls(NamTextBox)
            
        'добавление TextBox (конец интервала) для дат
        NamTextBox = "DateBoxFltrTo" & I
        Set ObgTextBox = frmConfSPR.FrameFiltr.Controls.Add("Forms.TextBox.1", NamTextBox)
            ObgTextBox.Height = 15.75
            ObgTextBox.Left = 390
            ObgTextBox.Top = TopPoz
            ObgTextBox.Width = 60
            ObgTextBox.MaxLength = 10
            ObgTextBox = Format(ObgTextBox, "dd.mm.yyyy")
            ObgTextBox.Locked = True
        ReDim Preserve cTBFiltrTo(I)
        Set cTBFiltrTo(I).TCBData = frmConfSPR.FrameFiltr.Controls(NamTextBox)
        
        'добавление TextBox для числовых и текстовых значений
        NamTextBox = "ComboBoxZn" & I
        Set ObgComboBox = frmConfSPR.FrameFiltr.Controls.Add("Forms.ComboBox.1", NamTextBox)
            ObgComboBox.Height = 15.75
            ObgComboBox.Left = 260
            ObgComboBox.Top = TopPoz
            ObgComboBox.Width = 215
            ObgComboBox.Style = fmStyleDropDownCombo
            ObgComboBox.TabStop = False
           v = ThisWorkbook.Sheets(frmMain.ComboBox_Лист.value).ListObjects(frmMain.ComboBox_Таблица.value).ListColumns(I + 1).Range
            ObgComboBox.list = v
'        ReDim Preserve cTB(i)
'        Set cTB(i).TB = frmConfSPR.FrameFiltr.Controls(NamTextBox)

        'добавление кнопки для вызова календаря (начало интервала)
        NamCommandButton = "CalendFrom" & I
        Set ObgCommandButton = frmConfSPR.FrameFiltr.Controls.Add("Forms.CommandButton.1", NamCommandButton)
            ObgCommandButton.Height = 15.75
            ObgCommandButton.Left = 345
            ObgCommandButton.Top = TopPoz
            ObgCommandButton.Width = 20
            ObgCommandButton.Caption = "..."
            ObgCommandButton.Font.Size = 6
            ObgCommandButton.TabStop = False
        ReDim Preserve cCBFiltrFrom(I)
        Set cCBFiltrFrom(I).CB = frmConfSPR.FrameFiltr.Controls(NamCommandButton)

        'добавление кнопки для вызова календаря (конец интервала)
        NamCommandButton = "CalendTo" & I
        Set ObgCommandButton = frmConfSPR.FrameFiltr.Controls.Add("Forms.CommandButton.1", NamCommandButton)
            ObgCommandButton.Height = 15.75
            ObgCommandButton.Left = 455
            ObgCommandButton.Top = TopPoz
            ObgCommandButton.Width = 20
            ObgCommandButton.Caption = "..."
            ObgCommandButton.Font.Size = 6
            ObgCommandButton.TabStop = False
        ReDim Preserve cCBFiltrTo(I)
        Set cCBFiltrTo(I).CB = frmConfSPR.FrameFiltr.Controls(NamCommandButton)
   
        'Пояснение "с"
        NamLabel = "LabelFrom" & I
        Set ObgLabel = frmConfSPR.FrameFiltr.Controls.Add("Forms.Label.1", NamLabel)
            ObgLabel.Height = 15.75
            ObgLabel.Left = 270
            ObgLabel.Top = TopPoz
            ObgLabel.Width = 10
            ObgLabel.Caption = "c"
            ObgLabel.Font.FontStyle = "полужирный"
            
        'Пояснение "по"
        NamLabel = "LabelTo" & I
        Set ObgLabel = frmConfSPR.FrameFiltr.Controls.Add("Forms.Label.1", NamLabel)
            ObgLabel.Height = 15.75
            ObgLabel.Left = 375
            ObgLabel.Top = TopPoz
            ObgLabel.Width = 10
            ObgLabel.Caption = "по"
            ObgLabel.Font.FontStyle = "полужирный"
            
        'добавление кнопок для очистки TextBox-ов
        NamCommandButton = "CommandCleare" & I
        Set ObgCommandButton = frmConfSPR.FrameFiltr.Controls.Add("Forms.CommandButton.1", NamCommandButton)
            ObgCommandButton.Height = 15.75
            ObgCommandButton.Left = 480
            ObgCommandButton.Top = TopPoz
            ObgCommandButton.Width = 20
            ObgCommandButton.Caption = "Х"
            ObgCommandButton.ForeColor = &HFF&
            ObgCommandButton.Font.Size = 6
            ObgCommandButton.Font.FontStyle = "полужирный"
            ObgCommandButton.TabStop = False
        ReDim Preserve cClear(I)
        Set cClear(I).Cleare = frmConfSPR.FrameFiltr.Controls(NamCommandButton)
        

        'Проверка значений таблицы
        ReDim Preserve FormatD(I)
        If ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value = "" Then
        ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value = 1
            If IsDate(ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value) Then
                FormatD(I) = "Дата"
            Else
                If IsNumeric(ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value) Then
                    FormatD(I) = "Число"
                Else
                    FormatD(I) = "Текст"
                End If
            End If
        ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value = ""
        Else
            If IsDate(ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value) Then
                FormatD(I) = "Дата"
            Else
                If IsNumeric(ActiveSheet.ListObjects(NameTable).Range(3, I + 1).value) Then
                    FormatD(I) = "Число"
                Else
                    FormatD(I) = "Текст"
                End If
            End If
        End If

        If FormatD(I) = "Дата" Then 'Если значение является датой
            frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).list = ThisWorkbook.Sheets("Справочники").Range("b2:b2").value
                frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).ListIndex = 0
            frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I).Visible = True
            frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I).Visible = True
            frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I).Visible = False
            frmConfSPR.FrameFiltr.Controls("CalendFrom" & I).Visible = True
            frmConfSPR.FrameFiltr.Controls("CalendTo" & I).Visible = True
            frmConfSPR.FrameFiltr.Controls("LabelFrom" & I).Visible = True
            frmConfSPR.FrameFiltr.Controls("LabelTo" & I).Visible = True
        Else
            If FormatD(I) = "Число" Then 'Если значение является числом
                frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).Visible = True
                    frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).list = ThisWorkbook.Sheets("Справочники").Range("b7:b12").value
                    frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).ListIndex = 0
                frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I).Visible = True
                frmConfSPR.FrameFiltr.Controls("CalendFrom" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("CalendTo" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("LabelFrom" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("LabelTo" & I).Visible = False
            Else 'Если значение является текстом
                frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).Visible = True
                    frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).list = ThisWorkbook.Sheets("Справочники").Range("b3:b8").value
                    frmConfSPR.FrameFiltr.Controls("ComboBoxFltr" & I).ListIndex = 0
                frmConfSPR.FrameFiltr.Controls("DateBoxFltrFrom" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("DateBoxFltrTo" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("TextBoxFltr" & I).Visible = True
                frmConfSPR.FrameFiltr.Controls("CalendFrom" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("CalendTo" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("LabelFrom" & I).Visible = False
                frmConfSPR.FrameFiltr.Controls("LabelTo" & I).Visible = False
            End If
        End If
        TopPoz = TopPoz + 18
    Next I
               
    ' Запуск сформированной формы, выставляем высоту окна и Frame1
    If TopPoz > 324 Then
    frmConfSPR.Height = 405
    frmConfSPR.FrameFiltr.Height = 324
    Else
        If TopPoz < 42 Then
        frmConfSPR.Height = 85
        frmConfSPR.FrameFiltr.Height = 31
        Else
        frmConfSPR.Height = TopPoz + 91
        frmConfSPR.FrameFiltr.Height = TopPoz + 10
        End If
    End If
    ' Выставляем высоту прокрутки на случай если она понадобиться
    frmConfSPR.FrameFiltr.ScrollHeight = TopPoz
    ' Проверяем необходимоть включения прокрутки во Frame1
    If TopPoz > frmConfSPR.FrameFiltr.Height Then
    frmConfSPR.FrameFiltr.ScrollBars = fmScrollBarsVertical
    Else
    frmConfSPR.FrameFiltr.ScrollBars = fmScrollBarsNone
    End If
    frmConfSPR.MultiPageConf.Pages("Filtr").Caption = frmConfSPR.MultiPageConf.Pages("Filtr").Caption & " " & NameTable
    frmConfSPR.Show
Else
MsgBox ("Данный объект в книге отсутствует")
End If
End Sub
