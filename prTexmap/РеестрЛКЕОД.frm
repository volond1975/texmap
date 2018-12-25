VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} РеестрЛКЕОД 
   OleObjectBlob   =   "РеестрЛКЕОД.frx":0000
   Caption         =   "UserForm1"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   690
End
Attribute VB_Name = "РеестрЛКЕОД"
Attribute VB_Base = "0{72AD6E03-F796-46C7-8275-A921DA408475}{6C675F44-294B-4874-B410-8B8A6E00380C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

 Private NotUse As Boolean   ' switch
 Private Arr()               ' array


Private Sub cmbboxЛісництво_Change()
'Dim ws As Worksheet
'Dim rf As Range
'Dim crf As Range
'Dim lis As Range
'Dim loj As clsmListObjs
'Dim wb As Workbook
'Dim v As Range
'Dim lop As ListObject
'Set wb = ActiveWorkbook
'Dim mufcntr As Object
'Set loj = New clsmListObjs
''Set muf = OpenUserForms(UFName)
'Set mufcntr = UserFormControl(Me, "cmbboxЛісництво")
'
''mufcntr.Clear
'With loj
'.Initialize wb
'Set wa = .ValueListObject("Форма", "Параметр", "Адрес", "Лісництво")
'Range(wa) = mufcntr.Value
'Set lop = .Items(Me.name)
' If mufcntr.Value <> "" Then Set w = ControlValueAddListObject(wb, Me, "cmbboxЛісництво", "Форма", "Параметр", "Значение", "Лісництво", cmboxNLK.Value)
''Set w = .ValueListObject("Форма", "Параметр", "Значение", "Лісництво")
''w.value = cmboxNLK.value
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxCKR", "Лісництва", "Сокр")
'Set z = ControlAddalue(Me, "cmboxCKR", "Лісництва", "Підрозділ", "Сокр", mufcntr.Value)

'End If
''If ListObjName Like "*[!_0-9]*" Then
'Set lop = .Items(Me.name)
'Set ws = lop.parent
'If lop.AutoFilter.FilterMode = True Then lop.AutoFilter.ShowAllData
'il = .IndexColumn(lop, "Лісництво")
'If mufcntr.value = "" Then
'lop.Range.AutoFilter Field:=il
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxNLK", Me.name, "Номер")
'Else
'lop.Range.AutoFilter Field:=il, _
'        Criteria1:=mufcntr.value
'        End If
' Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxNLK", Me.name, "Номер", True)
' Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxCKR", "Лісництва", "Сокр")
''
'Set z = ControlAddalue(Me, "cmboxCKR", "Лісництва", "Підрозділ", "Сокр", mufcntr.Value)
''' Лісництва [Сокр]Підрозділ
'''End If
'End With
End Sub

Private Sub cmboxCKR_Change()

End Sub

Private Sub cmboxFileName_Change()
Dim ps As PathSplitString
ps = PathSplit(Me.cmboxPath & "\" & Me.cmboxFileName)
If Not ps.bFileExists Then
LabelPath.Caption = "Файл не найден: "
Else
Изменение = DataCreatedFile(Me.cmboxPath & "\" & Me.cmboxFileName, "Последнее изменение")
LabelFileName.Caption = "Файл изменен: " & Изменение
End If
End Sub

Private Sub cmboxNDil_Change()
'Dim ws As Worksheet
'Dim rf As Range
'Dim crf As Range
'Dim lis As Range
'Dim loj As clsmListObjs
'Dim wb As Workbook
'Dim v As Range
'Dim lop As ListObject
'Set wb = ActiveWorkbook
'Dim mufcntr As Object
'Dim lkcntr As Object
'Dim cntrs As Object
'Set loj = New clsmListObjs
''Set muf = OpenUserForms(UFName)
'Set mufcntr = UserFormControl(Me, "cmboxNDil")
'Set lkcntr = UserFormControl(Me, "cmboxNLK")
'Сведено = mufcntr.Value & " " & lkcntr.Value
'cnr = Array("tbPl", "tboxVidRubki", "tboxS", "tboxPo", "tboxDilova", "tboxDrov", "tboxLekvid")
'
''mufcntr.Clear
'With loj
'.Initialize wb
''If ListObjName Like "*[!_0-9]*" Then
'Set lop = .Items(Me.name)
''Set loTexkarta = .Items("Техкарта")
'Set SvedT = .ValueListObject("Техкарта", "Сведено", "Сведено", Сведено)
'If Not SvedT Is Nothing Then Me.cbVzyatSTexkartu.Enabled = True
'Set ws = lop.parent
'If lop.AutoFilter.FilterMode = True Then lop.AutoFilter.ShowAllData
'il = .IndexColumn(lop, "Сведено")
'
'
'If mufcntr.Value = "" Then
' z = .myFilterListObject(lop, "Ділянка")
'lop.Range.AutoFilter Field:=il
'For i = 0 To UBound(cnr)
' Set ctrls = UserFormControl(Me, cnr(i))
'ctrls.Value = ""
'     Next i
'Else
''z = .myFilterListObject(lop, "Ділянка", mufcntr.value)
''Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxNDil", Me.name, "Ділянка", True)
'lop.Range.AutoFilter Field:=il, _
'        Criteria1:=Сведено
'     naz = Array("Площа", "Вид рубки", "Дата початку", "Дата закінчення", "Ділова", "Дров'яна", "Всього куб. м.")
'
''     i = 1
'     For i = 0 To UBound(cnr)
'     Set z = ControlAddalue(Me, cnr(i), Me.name, "Сведено", naz(i), Сведено)
''     If mufcntr.Value <> "" Then Set w = ControlValueAddListObject(wb, Me, cnr(i), "Форма", "Параметр", "Значение", naz(i), cmboxNLK.Value)
'     Next i
'        End If
'
''End If
'End With
End Sub

Private Sub cmboxNLK_Change()
'Dim ws As Worksheet
'Dim rf As Range
'Dim crf As Range
'Dim lis As Range
'Dim loj As clsmListObjs
'Dim wb As Workbook
'Dim v As Range
'Dim lop As ListObject
'Set wb = ActiveWorkbook
'Dim mufcntr As Object
'Set loj = New clsmListObjs
''Set muf = OpenUserForms(UFName)
'Set mufcntr = UserFormControl(Me, "cmboxNLK")
'
''mufcntr.Clear
'With loj
'.Initialize wb
'Set w = .ValueListObject("Форма", "Параметр", "Значение", "Лісорубний квиток №")
'Set wa = .ValueListObject("Форма", "Параметр", "Адрес", "Лісорубний квиток №")
''cmboxNLK.ControlSource = wa.Value
''w.Value = cmboxNLK.Value
''If ListObjName Like "*[!_0-9]*" Then
'Set lop = .Items(Me.name)
'Set ws = lop.parent
'ws.Activate
'If lop.AutoFilter.FilterMode = True Then lop.AutoFilter.ShowAllData
'il = .IndexColumn(lop, "Номер")
'
'
'If mufcntr.Value = "" Then
' z = .myFilterListObject(lop, "Ділянка")
'lop.Range.AutoFilter Field:=il
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxNDil", Me.name, "Ділянка")
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmbboxЛісництво", Me.name, "Лісництво")
'Else
''z = .myFilterListObject(lop, "Ділянка", mufcntr.value)
'lop.AutoFilter.ShowAllData
'lop.Range.AutoFilter Field:=il, _
'        Criteria1:=mufcntr.Value
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxNDil", Me.name, "Ділянка", True)
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmbboxЛісництво", Me.name, "Лісництво", True)
'Set mufcntrl = UserFormControl(Me, "cmbboxЛісництво")
'If mufcntrl.ListCount = 2 Then
'mufcntrl.ListIndex = 0
'        End If
'
'End If
'End With
End Sub

Private Sub cmboxPath_Change()
Dim ps As PathSplitString
ps = PathSplit(cmboxPath.value & "\")
If Not ps.bFolderExists Then
LabelPath.Caption = "Путь не найден: "
End If
End Sub

Private Sub cmboxTypeRubki_Change()
Dim locls As clsmListObjs
Dim lo_forma As ListObject
Dim loc As ListColumn
Dim locAddress As ListColumn
Dim lo As ListObject
Dim r As Range
Dim AR As Range
Dim n As Range
Dim wb As Workbook
Set wb = ThisWorkbook
Set locls = New clsmListObjs
With locls
.Initialize wb
Set lo_forma = .items("Форма")
Set w = .ValueListObject("Форма", "Параметр", "Значение", "Рубка Лист")

 If cmboxTypeRubki.value = "" Then
 wb.Worksheets("Форма").Activate
 Else
 w.value = cmboxTypeRubki.value
  wb.Worksheets(cmboxTypeRubki.value).Activate

Set loc = lo_forma.ListColumns(cmboxTypeRubki.value)
Set locAddress = lo_forma.ListColumns("Адрес")
Set locControlName = lo_forma.ListColumns("ControlName")
Set locParametrName = lo_forma.ListColumns("Параметр")
For Each r In loc.DataBodyRange.Cells
If r = 1 Then
Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).Enabled = True


If locAddress.DataBodyRange.Cells(r.Row - 1, 1) Like "* *" Then
Y = "'" & locAddress.DataBodyRange.Cells(r.Row - 1, 1).value
Y = VBA.Replace(Y, "!", "'!")
Else

Y = locAddress.DataBodyRange.Cells(r.Row - 1, 1).value
End If

If locAddress.DataBodyRange.Cells(r.Row - 1, 1) <> "" Then
Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).value = Range(Y).value
'Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).Value).ControlSource = Y
If Not Range(Y).Formula Like "=*" Then Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).ControlSource = Y
'Set w = .ValueListObject("Форма", "Параметр", "Значение", locParametrName.DataBodyRange.Cells(r.Row - 1, 1).Value)
'w.Value = Range(Y).Value 'Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).Value)
Else
 Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).Enabled = False
End If
'
'
Else
If locControlName.DataBodyRange.Cells(r.Row - 1, 1).value <> "" Then Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).Enabled = False
End If
'End If
Next
'
'
  
 End If
End With
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub CommandButton21_Click()
Call RefreshQuery(Me.cmboxQuerys)
End Sub



Private Sub CommandButton22_Click()
Dim locls As clsmListObjs
Dim lo_forma As ListObject
Dim loc As ListColumn
Dim locAddress As ListColumn
Dim lo As ListObject
Dim r As Range
Dim AR As Range
Dim n As Range
Dim wb As Workbook
Set wb = ThisWorkbook
Set locls = New clsmListObjs
With locls
.Initialize wb
Set lo_forma = .items("Форма")
Set w = .ValueListObject("Форма", "Параметр", "Значение", "Рубка Лист")

 If cmboxTypeRubki.value = "" Then
 wb.Worksheets("Форма").Activate
 Else
 w.value = cmboxTypeRubki.value
  wb.Worksheets(cmboxTypeRubki.value).Activate

Set loc = lo_forma.ListColumns(cmboxTypeRubki.value)
Set locAddress = lo_forma.ListColumns("Адрес")
Set locControlName = lo_forma.ListColumns("ControlName")
Set locParametrName = lo_forma.ListColumns("Параметр")
For Each r In loc.DataBodyRange.Cells
If r = 1 Then
Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).Enabled = True


If locAddress.DataBodyRange.Cells(r.Row - 1, 1) Like "* *" Then
Y = "'" & locAddress.DataBodyRange.Cells(r.Row - 1, 1).value
Y = VBA.Replace(Y, "!", "'!")
Else

Y = locAddress.DataBodyRange.Cells(r.Row - 1, 1).value
End If

If locAddress.DataBodyRange.Cells(r.Row - 1, 1) <> "" Then
Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).value = Range(Y).value
'Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).Value).ControlSource = Y
If Not Range(Y).Formula Like "=*" Then Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).ControlSource = Y
'Set w = .ValueListObject("Форма", "Параметр", "Значение", locParametrName.DataBodyRange.Cells(r.Row - 1, 1).Value)
'w.Value = Range(Y).Value 'Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).Value)
Else
 Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).Enabled = False
End If
'
'
Else
If locControlName.DataBodyRange.Cells(r.Row - 1, 1).value <> "" Then Me.Controls(locControlName.DataBodyRange.Cells(r.Row - 1, 1).value).Enabled = False
End If
'End If
Next
'
'
  
 End If
End With
End Sub

Private Sub CommandButton23_Click()
Call ВТеккарту
End Sub

Private Sub CommandButton6_Click()

Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs
Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
Dim loc As ListColumn
Set wb = ActiveWorkbook
Dim mufcntr As Object
Set loj = New clsmListObjs

Set mufcntr = UserFormControl(Me, "cmboxTypeRubki")

'Set mufcntradd = UserFormControl(Me, Me.MPage.Pages(Me.MPage.Value).Frame2.ActiveControl.name)
'If Not Me.MPage.Pages(Me.MPage.Value).Frame2.ActiveControl.name Like "T*" Then Exit Sub
With loj
.Initialize wb

Set lop = .items("Форма")

 If mufcntr.value <> "" Then
Set loc = lop.ListColumns(mufcntr.value)
 wb.Worksheets(mufcntr.value).Activate
 Set ПрефиксИмени = .ValueListObject("Шаблон", "Наименование", "Имя", mufcntr.value)
' Me.Controls("T" & j & i).ControlSource = Range(ПрефиксИмени.Value & "T_" & j & i).Address
For i = 1 To 3
For j = 1 To 4
Me.Controls("T" & j & i).ControlSource = Range(ПрефиксИмени.value & "T_" & j & i).Address
Range(ПрефиксИмени.value & "T_" & j & i).value = ""
Next j
Next i
End If


End With








End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub MPage_Change()
Select Case MPage.SelectedItem.name
    Case "Page1"
    Case "Page5"
    
End Select
End Sub

Private Sub T11_Change()

'Dim ws As Worksheet
'Dim rf As Range
'Dim crf As Range
'Dim lis As Range
'Dim loj As clsmListObjs
'Dim wb As Workbook
'Dim v As Range
'Dim lop As ListObject
'Set wb = ActiveWorkbook
'Dim mufcntr As Object
'Set loj = New clsmListObjs
'
'Set mufcntr = UserFormControl(Me, "cmboxTypeRubki")
'
'Set mufcntradd = UserFormControl(Me, Me.MPage.Pages(Me.MPage.Value).Frame2.ActiveControl.name)
'
'With loj
'.Initialize wb
'
'Set lop = .Items("Шаблон")
' If mufcntr.Value <> "" Then
' wb.Worksheets(mufcntr.Value).Activate
' Set ПрефиксИмени = .ValueListObject("Шаблон", "Наименование", "Имя", mufcntr.Value)
'Range(ПрефиксИмени.Value & VBA.Replace(Me.MPage.Pages(Me.MPage.Value).Frame2.ActiveControl.name, "T", "T_")).Value = mufcntradd.Value
'End If
'
'
'End With
End Sub

Private Sub T12_Change()

End Sub

Private Sub T13_Change()

End Sub

Private Sub T21_Change()

End Sub

Private Sub T22_Change()

End Sub

Private Sub T23_Change()

End Sub

Private Sub T31_Change()

End Sub

Private Sub T32_Change()

End Sub

Private Sub T33_Change()

End Sub

Private Sub T41_Change()

End Sub

Private Sub T42_Change()

End Sub

Private Sub T43_Change()

End Sub

Private Sub tbMasEOD_Change()
tboxLekvid = tbMasEOD
End Sub

Private Sub tbKilDer_Change()
Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs
Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
Set wb = ActiveWorkbook
Dim mufcntr As Object
Set loj = New clsmListObjs
'Set muf = OpenUserForms(UFName)
Set mufcntr = UserFormControl(Me, "tbCrHl")
'mufcntr.Clear
With loj
.Initialize wb
Set w = .ValueListObject("Форма", "ControlName", "Адрес", "tbCrHl")
If w Like "* *" Then
Y = "'" & w.value
Y = VBA.Replace(Y, "!", "'!")
Else
Y = w.value
End If
mufcntr.value = Range(Y).value
End With
'cmboxTypeRubki_Change
End Sub

Private Sub tbNel_Change()
tboxNeLekvid = tbNel
End Sub

Private Sub tboxNeLekvid_Change()
tbNel = tboxNeLekvid
End Sub

Private Sub tboxPlocha_Change()
tbPl = tboxPlocha
End Sub

Private Sub tboxS_Change()
Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs
Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
Set wb = ActiveWorkbook
Dim mufcntr As Object
Set loj = New clsmListObjs
'Set muf = OpenUserForms(UFName)
Set mufcntr = UserFormControl(Me, "tboxS")
'mufcntr.Clear
With loj
.Initialize wb
Set lop = .items(Me.name)
 If mufcntr.value <> "" Then
 z = VBA.Month(VBA.Format(mufcntr.value, "dd-mm-yyyy"))
Set w = .ValueListObject("Форма", "Параметр", "Значение", "Месяц ЛК")
w.value = z
Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxMothLk", "1_12_1", "Сокр")
Set z = ControlAddalue(Me, "cmboxMothLk", "Форма", "Параметр", "Значение", "Месяц ЛК")

End If
End With
End Sub

Private Sub tbPlEOD_Change()
 tboxPlocha = tbPlEOD
End Sub

Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
'RefreshQuery (Me.name)
Dim locls As clsmListObjs
Dim lo_forma As ListObject
Dim loc As ListColumn
Dim lo As ListObject
Dim r As Range
Dim AR As Range
Dim n As Range
Dim wb As Workbook
'Dim ps As PathSplitString
Set wb = ThisWorkbook
Set locls = New clsmListObjs
With locls
.Initialize wb
Me.MPage.value = 0
Set lo_forma = .items("Форма")
Set lo_forma_path = .items("Путь")
Set w = .ValueListObject("Форма", "Параметр", "Значение", "Рубка Лист")
Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxPath", "Путь", "Путь к книге отчета")
Set mufcntrl = UserFormControl(Me, "cmboxPath")
If mufcntrl.ListCount = 2 Then
mufcntrl.ListIndex = 0

        End If
Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxFileName", "Путь", "Назва")
Set mufcntrl = UserFormControl(Me, "cmboxFileName")
mufcntrl.ListIndex = 0
       
Set loc = lo_forma.ListColumns("Control")
For Each r In loc.DataBodyRange.Cells
If r.value <> "" Or Not IsEmpty(r.value) Then
Set Таблица = .ValueListObject("Форма", "Control", "Таблица", r.value)
Set Заголовки = .ValueListObject("Форма", "Control", "Заголовки", r.value)
Set По_Умолчанию = .ValueListObject("Форма", "Control", "По Умолчанию", r.value)
Set Значение2 = .ValueListObject("Форма", "Control", "Значение2", r.value)
'If Значение2 = "" Or IsEmpty(По_Умолчанию.value) Then
Call ControlAddListVisibleListObjectColumnValue(Me, r.value, Таблица.value, Заголовки.value)
If По_Умолчанию.value <> "" Or Not IsEmpty(По_Умолчанию.value) Then
' Select Case TypeName(По_Умолчанию.value)
'Case "Date"
'DataText = VBA.CDate(VBA.Day(По_Умолчанию.value) & "." & VBA.Month(По_Умолчанию.value) & "." & VBA.Year(По_Умолчанию.value))
'End Select
Call ControlAddalue(Me, r.value, "Форма", "Control", "По Умолчанию", r.value)
End If
End If
'End If
Next
Set r = Range("Шаблон[Наименование]").Find(ActiveSheet.name)
If Not r Is Nothing Then
Set mufcntr = UserFormControl(Me, "cmboxTypeRubki")
mufcntr.value = ActiveSheet.name
w.value = ActiveSheet.name
End If
End With
Me.cmboxTypeRubki.SetFocus

Dim q
Me.cmboxQuerys.Clear
For Each q In wb.Queries
Me.cmboxQuerys.AddItem q.name
Next
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmbboxЛісництво", "РеестрЛКЕОДConect", "Лісництво")
''Тарифи[№]
'Call ControlAddListVisibleListObjectColumnValue(Me, "cmboxDataTarif", "РеестрЛКЕОДConect", "Лісництво")

End Sub
 '-------------------------------------------
 Private Function getFormWidth(ByVal config As Variant) As Single
     If IsNumeric(config) Then
         If Not IsEmpty(config) Then
             If config >= 100 And config <= Application.Width / 2 Then
                 getFormWidth = CSng(config)
                 Exit Function
             End If
         End If
     End If
     getFormWidth = 210
 End Function


 Private Function getFormCaprion(ByVal config As Variant) As String
     If IsEmpty(config) Then
         getFormCaprion = "Unique records: "
         Exit Function
     End If
     getFormCaprion = config
 End Function


 Private Function getPattern(ByVal config As Variant) As String
     If VarType(config) = vbString Then
         If InStr(1, config, "request", vbTextCompare) Then
             getPattern = Replace(config, "request", Me.ComboBox1.Text)
             Exit Function
         End If
     End If
     getPattern = "*" & Me.ComboBox1.Text & "*"
 End Function


 Private Function getRegister(ByVal config As Variant) As Boolean
     If VarType(config) = vbBoolean Then
         getRegister = config
         Exit Function
     End If
     getRegister = False
 End Function


 Private Function getCase(ByVal Text As String, ByVal register As Boolean) As String
     If register Then getCase = Text Else getCase = LCase(Text)
 End Function


 Private Function getSearchCaption(ByVal config As Variant) As String
     If IsEmpty(config) Then
         getSearchCaption = "Search result: "
         Exit Function
     End If
     getSearchCaption = config
 End Function


 Private Function searchEnteredValue(ByVal config As Variant) As Boolean
     If VarType(config) = vbBoolean Then
         searchEnteredValue = config
         Exit Function
     End If
     searchEnteredValue = True
 End Function
Private Sub meComboBoxChange(contrName)
     If NotUse Then Exit Sub
     If Me.ComboBox1.Text = "" Then
         Me.Caption = getFormCaprion(DLLSheetSettings.Range("F41").value) & UBound(Arr) + 1
         Me.ComboBox1.List = Arr
         Exit Sub
     End If
     Dim elem, pattern As String, register As Boolean
     pattern = getPattern(DLLSheetSettings.Range("F2").value)
     register = getRegister(DLLSheetSettings.Range("F29").value)
     pattern = getCase(pattern, register)
     With CreateObject("Scripting.Dictionary")
         For Each elem In Arr
             If getCase(elem, register) Like pattern Then .Add CStr(elem), elem
         Next
         Me.Caption = getSearchCaption(DLLSheetSettings.Range("F50").value) & .Count
         Me.ComboBox1.List = .items
     End With
 End Sub

Sub TShange()
Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs
Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
Set wb = ActiveWorkbook
Dim mufcntr As Object
Set loj = New clsmListObjs

Set mufcntr = UserFormControl(Me, "cmboxTypeRubki")

Set mufcntradd = UserFormControl(Me, Me.MPage.Pages(Me.MPage.value).Frame2.ActiveControl.name)
If Not Me.MPage.Pages(Me.MPage.value).Frame2.ActiveControl.name Like "T*" Then Exit Sub
With loj
.Initialize wb

Set lop = .items("Шаблон")
 If mufcntr.value <> "" Then
 wb.Worksheets(mufcntr.value).Activate
 Set ПрефиксИмени = .ValueListObject("Шаблон", "Наименование", "Имя", mufcntr.value)
mufcntradd.ControlSource = Range(ПрефиксИмени.value & VBA.Replace(Me.MPage.Pages(Me.MPage.value).Frame2.ActiveControl.name, "T", "T_")).Address
End If


End With
End Sub
