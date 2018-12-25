VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   OleObjectBlob   =   "frmMain.frx":0000
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   3038
End
Attribute VB_Name = "frmMain"
Attribute VB_Base = "0{90B1F5E0-3803-49D8-AF12-8DA5E09E337E}{71291595-D10E-4B73-9B82-2BA22FC38CF1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub ComboBox_Лист_Change()
Dim wsh As Worksheet
If Me.ComboBox_Лист.value <> "" Then
Set wsh = ThisWorkbook.Worksheets(Me.ComboBox_Лист.value)
Me.ComboBox_Таблица.list = ArrayWorksheetsListObjectName(wsh)
Else
Me.ComboBox_Таблица.value = ""
End If
End Sub

Private Sub ComboBox_Таблица_Change()
bListBoxScrool = False
    frmConfSPR.FrameFiltr.Calendar1.value = CStr(Date)
    DelFiltrForAllTables
    ListBoxSotrudnikiRefresh
    WheelHook Me    'Начинаем отслеживать колесо мыши
End Sub

Private Sub FiltrSotrudniki_Click()
Menu1.CreateFiltr
End Sub

Private Sub UserForm_Activate()
'bListBoxScrool = False
'    frmConfSPR.FrameFiltr.Calendar1.Value = CStr(Date)
'    DelFiltrForAllTables
'    ListBoxSotrudnikiRefresh
'    WheelHook Me    'Начинаем отслеживать колесо мыши
End Sub

Private Sub UserForm_Initialize()
Me.ComboBox_Лист.list = ArrayWorksheetsName(ThisWorkbook)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    WheelUnHook    'Отменяем контроль отслеживания мыши при закрытии формы
End Sub

Private Sub UserForm_Deactivate()
    WheelUnHook    'Отменяем контроль отслеживания мыши при деактивации формы
End Sub

Private Sub frmMain_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
bListBoxScrool = False
Set objActiveListBox = Nothing
End Sub
'Private Sub ListBoxSotrudniki_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'bListBoxScrool = True
'Set objActiveListBox = frmMain.ListBoxSotrudniki
'End Sub

'Очищаем автофильтры таблиц
Sub DelFiltrForAllTables()
Application.GoTo Reference:=Me.ComboBox_Таблица.value
Selection.AutoFilter
Selection.AutoFilter
End Sub

'Загружаем справочник Сотрудники в ListBox
Sub ListBoxSotrudnikiRefresh()
Dim I As Range
Dim p
NameTable = Me.ComboBox_Таблица.value
ListBoxSotrudniki.Clear
'ListBoxSotrudniki.ColumnWidths = "1.2 cm;1.7 cm;6.55 cm;1.07 cm;1.07 cm;1.07 cm"
On Error Resume Next
Application.GoTo Reference:=NameTable
    If Err = 0 Then
    Application.GoTo Reference:=NameTable
    RowsTable = ActiveSheet.ListObjects(NameTable).ListRows.Count
    Selection.Columns(1).Select
        For Each I In Selection.SpecialCells(xlVisible)
        p = I.Row
            With frmMain.ListBoxSotrudniki
            .AddItem
            .Column(0, .ListCount - 1) = ActiveSheet.ListObjects(NameTable).Range(p, 1).value
            .Column(1, .ListCount - 1) = ActiveSheet.ListObjects(NameTable).Range(p, 2).value
            .Column(2, .ListCount - 1) = ActiveSheet.ListObjects(NameTable).Range(p, 3).value
            .Column(3, .ListCount - 1) = ActiveSheet.ListObjects(NameTable).Range(p, 4).value
            .Column(4, .ListCount - 1) = ActiveSheet.ListObjects(NameTable).Range(p, 5).value
            .Column(5, .ListCount - 1) = ActiveSheet.ListObjects(NameTable).Range(p, 6).value
            End With
        Progress = (p / (RowsTable + 1) * 99) + 1
        frmAuth.ProgressBar = Round(Progress, 0)
        frmConfSPR.ProgressProc = Round(Progress, 0)
        Next I
        frmMain.ListBoxSotrudniki.AddItem
    Else
    MsgBox ("Таблица " & NameTable & " в книге не найдена!")
    End If
End Sub
'Прокрутка справочников
Sub MouseWheel(ByVal Rotation As Long)
If bListBoxScrool = False Then
Exit Sub
Else
    If objActiveListBox.ListCount > 3 Then
        If Rotation > 0 Then
            'Scroll up
            If objActiveListBox.TopIndex > 3 Then
            objActiveListBox.TopIndex = objActiveListBox.TopIndex - 3
            Else
            objActiveListBox.TopIndex = 0
            End If
    Else
            'Scroll down
            If objActiveListBox.TopIndex = objActiveListBox.TopIndex + 3 > objActiveListBox.ListCount - 2 Then
            objActiveListBox.TopIndex = objActiveListBox.ListCount - 2
            Else
            objActiveListBox.TopIndex = objActiveListBox.TopIndex + 3
            End If
        End If
    Else
    Exit Sub
    End If
End If
End Sub

Function ArrayWorksheetsName(wb As Workbook)
Dim sh As Worksheet
Dim v()
Dim k
k = 0
ReDim v(wb.Worksheets.Count)
For Each sh In wb.Worksheets
v(k) = sh.name
k = k + 1
Next

ArrayWorksheetsName = v


End Function
Function ArrayWorksheetsListObjectName(wsh As Worksheet)
Dim lo As ListObject
Dim v()
Dim k
k = 0
ReDim v(wsh.ListObjects.Count)
For Each lo In wsh.ListObjects
v(k) = lo.name
k = k + 1
Next

ArrayWorksheetsListObjectName = v


End Function
