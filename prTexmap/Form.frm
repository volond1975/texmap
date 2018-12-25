VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   OleObjectBlob   =   "Form.frx":0000
   Caption         =   "ТехКарта"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7665
   TypeInfoVer     =   691
End
Attribute VB_Name = "Form"
Attribute VB_Base = "0{8D1A5248-6E0D-47E1-955F-E262840D6DA9}{0DE96912-93C7-42EE-8ACA-6CDE18A6D53A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub cbOsh_Click()

End Sub

Private Sub butDataDog_Click()
'Range("valCalend").ClearContents
'Range("FormCalend").Value = "Form"
'Range("ContrCalend").Value = "cbDataDog"
'Call KalendarShow
'Kalendar.Calendar1.Value = Me.cbDataDog.Value
''Range("cbDataDog").Value = Range("valCalend").Value
'Me.Controls(Range("ContrCalend").Value).Value = MonthTK(VBA.Day(Range("valCalend").Value)) & "." & MonthTK(VBA.Month(Range("valCalend").Value)) & "." & VBA.Year(Range("valCalend").Value)
Me.cbDataDog = Get_Date1(Me.cbDataDog, Now)

End Sub

Private Sub cbDataDog_Change()

End Sub

Private Sub cbDataS_Change()

End Sub

Private Sub cbDataT_Change()
cbDataT.value = Format(cbDataT.value, "dd-mm-yyyy")
Range("умДатаТарифы") = VBA.CDate(MonthTK(VBA.Day(cbDataT.Text)) & "." & MonthTK(VBA.Month(cbDataT.Text)) & "." & VBA.Year(cbDataT.Text))
Call CalculationA
Call CalculationM
End Sub

Private Sub cbFiltr_Click()

End Sub

Private Sub cbK_Skladanny_Change()
Range("mK_Skladanny") = cbK_Skladanny.value
End Sub

Private Sub cbLisn_Change()
s = СокрTK(Form.cbLisn.value)

Set r = NTK
Form.cbN.value = r.value
'r.Value = r + 1
'Заполняем відповидальними
s = СокрTK(Form.cbLisn.value)
With Form
Select Case .cbShablon.Text

Case "РГК"
z = "g"
Case "РД"
z = "d"
Case "Молодняк"
z = "m"
End Select
End With
Me.cbVukon.ControlSource = "=" & z & "Vukon"
Me.tbVukon.ControlSource = "=" & z & "Vukon"
Me.tbK_Vukon.ControlSource = "=" & z & "K_Vukon"
Me.cbVukon.value = VukonPoLisn(Me.cbLisn.value)
Me.tbK_Vukon.value = KoefPoVukon(Me.cbVukon.value)
'LeterSheet & "K_Vukon"
Form.cbVidpovidal.Clear
Set r = Range("ma" & s)
For Each r In r.Cells
Form.cbVidpovidal.AddItem r.value
Next r

End Sub

Private Sub cbMonth_Change()
Form.cbNTK = НомерТК
If Form.cbShablon.Text = "" Then Exit Sub
With Form
Select Case .cbShablon.Text
Case "РГК"
z = "g"
    
Case "РД"
    z = "d"
   
 Case "Молодняк"
z = "m"
  
    
   End Select
 End With
 Range(z & "NTK").value = Form.cbNTK
End Sub

Private Sub cbMonthPR_Change()
Dim z
Form.cbNTK = НомерТК
If Form.cbShablon.Text = "" Then Exit Sub
With Form
Select Case .cbShablon.Text
Case "РГК"
z = "g"
    
Case "РД"
    z = "d"
   
 Case "Молодняк"
z = "m"
  
    
   End Select
 End With
 Range(z & "NTK").value = Form.cbNTK
End Sub

Private Sub cbN_Change()
Form.cbNTK = НомерТК
With Form
Select Case .cbShablon.Text
Case "РГК"
z = "g"
    
Case "РД"
    z = "d"
   
 Case "Молодняк"
z = "m"
  
    
   End Select
 End With
 Range(z & "NTK").value = Form.cbNTK
End Sub

Private Sub cbNTK_Change()

End Sub

Private Sub cbPodcherk_Change()
If Me.cbPodcherk.value = "_" Then
cbVidRubki.value = cbVidRubki.value & Me.cbPodcherk.value
Else
Dim k As String
k = cbVidRubki.value
Mid(k, Len(cbVidRubki.value), 1) = " "
cbVidRubki.value = Trim(k)
End If
End Sub

Private Sub cbPodcherk_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.cbPodcherk.value = "_" Then
cbVidRubki.value = cbVidRubki.value & Me.cbPodcherk.value
Else
Dim k As String
k = cbVidRubki.value
Mid(k, Len(cbVidRubki.value), 1) = " "
cbVidRubki.value = Trim(k)
End If
End Sub

Private Sub cbPodcherkN_Change()
If Me.cbPodcherkN.value = "Подчеркнуть" Then
cbN.value = cbN.value & "_" & cbMonth
Else
Dim k As String
k = cbN.value
q = 0
х = 0
For i = Len(cbN.value) To 1 Step -1
z = Mid(k, i, 1)
If z = "_" Then
q = i
х = 1
Exit For
End If
Next i
If X = 0 Then Exit Sub
'Mid(k, Len(cbVidRubki.Value) - 1, 2) = " "
For i = q To Len(cbN.value)
Mid(k, i, 1) = " "
Next i
cbN.value = Trim(k)
End If
End Sub

Private Sub cbPodcherkN_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.cbPodcherkN.value = "Подчеркнуть" Then
cbN.value = cbN.value & "_" & cbMonth
Else
Dim k As String
k = cbN.value
q = 0
х = 0
For i = Len(cbN.value) To 1 Step -1
z = Mid(k, i, 1)
If z = "_" Then
q = i
х = 1
Exit For
End If
Next i
If q = 0 Then Exit Sub
'Mid(k, Len(cbVidRubki.Value) - 1, 2) = " "
For i = q To Len(cbN.value)
Mid(k, i, 1) = " "
Next i
cbN.value = Trim(k)
End If
End Sub

Private Sub cbShablon_Change()
Call CalculationM
Call ЗаполнитьФормуТехкарты(Me.cbShablon.Text)
Me.mp1.value = 1
Range("cVidShablona") = cbShablon.value




End Sub

Private Sub cbSpRubki_Change()
Range("mSpRob") = cbSpRubki.value
End Sub

Private Sub cbVidRubki_Change()

End Sub

Private Sub cbVukon_Change()
Me.tbVukon = Me.cbVukon.value
Me.tbK_Vukon = KoefPoVukon(Me.cbVukon.value)
Range("cVukon").value = Me.tbVukon.value
'Call CalculationA
'Call CalculationM
End Sub

Private Sub chbCkl_Click()
If Me.chbCkl Then
Range("mVRSkl").value = 1
Form.cbK_Skladanny.Visible = True
Form.lK_Skladanny.Visible = True
Form.cbK_Skladanny = 0.3
Form.cbK_Skladanny.value = Range("mK_Skladanny")
Else
Range("mVRSkl").value = 0
'-----------------------------
Form.cbK_Skladanny.Visible = False
Form.lK_Skladanny.Visible = False

'--------------------------
End If
End Sub

Private Sub cmbTenderData_Change()
cbDataS = cmbTenderData
End Sub

Private Sub cmbTender№_Change()

Form.cbNTK = НомерТК
If Form.cbShablon.Text = "" Then Exit Sub
With Form
Select Case .cbShablon.Text
Case "РГК"
z = "g"
    
Case "РД"
    z = "d"
   
 Case "Молодняк"
z = "m"
  
    
   End Select
 End With
 Range(z & "NTK").value = Form.cbNTK
 
 Dim loj As clsmListObjs
 Dim wb As Workbook
 Set wb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize wb
Me.cmbTenderData.Text = ""
Set d = loj.ValueListObject("Тендер", "Тендер", "Дата", cmbTender№.Text)
 If Not d Is Nothing Then Me.cmbTenderData.Text = d.value
Set Виконавець = loj.ValueListObject("Тендер", "Тендер", "Виконавець", cmbTender№.Text)
 If Not Виконавець Is Nothing Then Me.cbVukon.Text = Виконавець.value
End With
 
 
 
 
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub CommandButton1_Click()
Dim sh As Worksheet
Set sh = ThisWorkbook.Worksheets(Form.cbShablon.Text)
sh.Activate
Call РасчитатьТехКарту
sh.PageSetup.BlackAndWhite = Form.chbBlecAndWitePrint
Call ПечатьВыделенное

End Sub

Private Sub CommandButton10_Click()
Set r = NTK

r.value = r + 1
Form.cbN.value = r.value

q = LastRow("Тарифи")

'Установим актуальные тарифные ставки
ats = Application.WorksheetFunction.max(Range(Worksheets("Тарифи").Cells(2, 1), Worksheets("Тарифи").Cells(q, 1)))
Range("умДатаТарифы") = ats
Range("mK_Skladanny").value = 0.5
cbDataT.value = Range("умДатаТарифы")
Me.mp1.value = 2

Me.CommandButton6.SetFocus
End Sub

Private Sub CommandButton11_Click()
If Me.cbVidRubki = "" Then
MsgBox "Від рубки не выбран"
Me.cbMonth.SetFocus
Exit Sub
End If
If Me.cbLisn = "" Then
MsgBox "Лісництво не вибрано"
Me.cbLisn.SetFocus
Exit Sub
End If
If Me.tbKV = "" Then
MsgBox "Квартал не вибран"
Me.cbN.SetFocus
Exit Sub
End If

If Me.tbVud = "" Then
MsgBox "Видел не вибран"
Me.cbN.SetFocus
Exit Sub
End If


n = ID_ТК
Set r = Worksheets("Техкарты").Columns("AH:AH").Find(What:=n, LookAt:=xlWhole)

If r Is Nothing Then
MsgBox "Техкарта с  номером " & n & " не найдена"
Else
Worksheets("Техкарты").Activate
r.Activate
End If
End Sub

Private Sub CommandButton12_Click()
v = Split(mp1.SelectedItem.name, "_")
If v(1) = 4 Then
mp1.value = 0
Else
mp1.value = v(1) + 1 - 1
End If
End Sub

Private Sub CommandButton13_Click()
v = Split(mp1.SelectedItem.name, "_")
If v(1) = 1 Then
mp1.value = 3
Else
mp1.value = v(1) - 2
End If
End Sub


Private Sub CommandButton14_Click()

'Range("valCalend").ClearContents
'Range("FormCalend").Value = "Form"
'Range("ContrCalend").Value = "cbDataS"
'Call KalendarShow
''Range("cbDataDog").Value = Range("valCalend").Value
'Me.Controls(Range("ContrCalend").Value).Value = MonthTK(VBA.Day(Range("valCalend").Value)) & "." & MonthTK(VBA.Month(Range("valCalend").Value)) & "." & VBA.Year(Range("valCalend").Value)
Me.cbDataS = Get_Date1(Me.cbDataS, Now)
End Sub

Private Sub CommandButton15_Click()
Dim B As Workbook
Set B = mywbBook(Me.ComboBox_cDog, ThisWorkbook.Path & "\")
If B Is Nothing Then MsgBox ("Файл " & ThisWorkbook.Path & "\" & Me.ComboBox_cDog)
ThisWorkbook.Save
ThisWorkbook.Close
'With Dokument
'.Show 0
'.cbLisn = Form.cbLisn
'.cbVidRubki = Form.cbVidRubki
'.tbKV = Form.tbKV
'.tbVud = Form.tbVud
'.cbN = Form.cbN
''Dokument.CommandButton1.SetFocus
'
'.mp1.Value = 1
'.CommandButton11.SetFocus
''.CommandButton11.Value = True
'.tbSumaDog.Value = Form.tbSumDog.Value
'.tbZalushok.Value = Form.tbZalushok.Value
'.cbMonth.Value = Form.cbMonth.Value
'.cbMonthPR.Value = Form.cbMonthPR.Value
'.cbNLK.Value = Form.tbNLK.Value
'.cbDataDog.Value = Form.cbDataDog.Value
'.cbDataS.Value = Form.cbDataS.Value
'.tbPlocha.Value = Form.tbPlocha.Value
'.tbSumZalushok.Value = Form.tbZalushokSum.Value
'.tbPlochaRob.Value = Form.tbZalushokPlosha.Value
'If .tbPlochaRob <> "" Then .tbPlocha = .tbPlochaRob
'End With
End Sub

Private Sub CommandButton16_Click()
Dim tbook As Workbook
Dim kbook As Workbook
Dim kbook_path_name As String
Set tbook = ThisWorkbook
kbook_path_name = tbook.Path & "\" & "Кошторис.xls"
k = Worksheets("Константы").Range("cVukPDV").value
Set kbook = mywbBook("Кошторис.xls", kbook_path_name)
With Form
Select Case .cbShablon.Text

Case "РГК"
z = "g"
tbook.Activate
del = Val(T11.value) + Val(T12.value) + Val(T13.value) _
+ Val(T21.value) + Val(T22.value) + Val(T23.value)
drov = Val(T31.value) + Val(T32.value) + Val(T33.value) _
+ Val(T41.value) + Val(T42.value) + Val(T43.value)
kbook.Activate
Worksheets(cbShablon.value & " " & tbVukon.value).Activate
'Worksheets(cbShablon.Value & " " & cbVukon.Value).Activate
'Корегування массі
If Val(Me.tbZalushok) <> 0 Then
Worksheets(cbShablon.value & " " & tbVukon.value).Range(cbShablon.value & "_" & tbVukon.value & "_kil").value = _
Me.tbZalushok.value
Else
Range(cbShablon.value & "_" & tbVukon.value & "_kil").value = _
Me.tbMas.value
End If
If Val(tbK_Vukon.value) = k Then
Range(cbShablon.value & "_" & tbVukon.value & "_opls").NumberFormat = "#,##0.000"
Range(cbShablon.value & "_" & tbVukon.value & "_opls").value = _
Val(tbVutrKbm.value) - Val(tbVutrKbmPDV.value)


Range(cbShablon.value & "_" & tbVukon.value & "_opls").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_opls").value, 3)
Range(cbShablon.value & "_" & cbVukon.value & "_opls").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & cbVukon.value & "_opls").value, 3)
Else
Range(cbShablon.value & "_" & tbVukon.value & "_opls").NumberFormat = "#,##0.000"

Range(cbShablon.value & "_" & tbVukon.value & "_opls").value = _
Val(tbVutrKbm.Text)


Range(cbShablon.value & "_" & tbVukon.value & "_opls").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_opls").value, 3)

End If

If Val(tbK_Vukon.value) = k Then
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").NumberFormat = "#,##0.000"
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = _
Val(tbVutrHl.value) - Val(tbVutrHlPDV.value)

Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value, 3)
Else
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").NumberFormat = "#,##0.000"
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = Val(tbVutrHl.value)

Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value, 3)
End If

Case "РД"
z = "d"
tbook.Activate
del = Val(T11.value) + Val(T12.value) + Val(T13.value) _
+ Val(T21.value) + Val(T22.value) + Val(T23.value)
drov = Val(T31.value) + Val(T32.value) + Val(T33.value) _
+ Val(T41.value) + Val(T42.value) + Val(T43.value)
kbook.Activate
'mas = Range(z & "Mas").Value
kbook.Worksheets(cbShablon.value & " " & tbVukon.value).Activate
'Корегування массі
If Val(Me.tbZalushok) <> 0 Then
Range(cbShablon.value & "_" & tbVukon.value & "_kil").value = _
Me.tbZalushok.value
Else
Range(cbShablon.value & "_" & tbVukon.value & "_kil").value = _
Me.tbMas.value
End If
'Range(cbShablon.Value & "_" & tbVukon.Value & "_kil").Value = _
Me.tbMas



If Val(tbK_Vukon.value) = k Then
Range(cbShablon.value & "_" & tbVukon.value & "_opl").NumberFormat = "#,##0.000"
Range(cbShablon.value & "_" & tbVukon.value & "_opl").value = _
Val(tbVutrKbm.value) - Val(tbVutrKbmPDV.value) '+ tbVutrKbm.Value * 0.2

Range(cbShablon.value & "_" & tbVukon.value & "_opl").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_opl").value, 3)

Else
Range(cbShablon.value & "_" & tbVukon.value & "_opl").NumberFormat = "#,##0.000"
Range(cbShablon.value & "_" & tbVukon.value & "_opl").value = _
Val(tbVutrKbm.value)

Range(cbShablon.value & "_" & tbVukon.value & "_opl").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_opl").value, 3)
End If
kbook.Activate
Case "Молодняк"
z = "m"
kbook.Activate

'Корегування массі
If Val(Me.tbZalushok) <> 0 Then
Range(cbShablon.value & "_" & tbVukon.value & "_kilx").value = _
Me.tbZalushok.value
del = Me.tbZalushok.value
Else
Range(cbShablon.value & "_" & tbVukon.value & "_kilx").value = _
Me.tbMas.value
del = Me.tbMas.value
End If




'Range(cbShablon.Value & "_" & tbVukon.Value & "_kilx").Value = _
Me.tbMas.Value
'del = Me.tbMas.Value

drov = _
Val(T41.value) + Val(T42.value) + Val(T43.value)
kbook.Activate

'If tbook.Range(z & "Mas") = 1 Then
'Range(Worksheets("Формулы").Cells(11, 2) & "_" & Worksheets("Формулы").Cells(2, 2) & "_opls").Value = _
'tbook.Range(z & "VutrKbm") + tbook.Range(z & "VutrKbm") * 0.2
'Else
'Range(Worksheets("Формулы").Cells(11, 2) & "_" & Worksheets("Формулы").Cells(2, 2) & "_opls").Value = _
'tbook.Range(z & "VutrKbm")
'End If

If Val(tbK_Vukon.value) = k Then
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").NumberFormat = "#,##0.000"
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = _
Val(tbVutrHl.value) - Val(tbVutrHlPDV.value) * 1

Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value, 3)
Else
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").NumberFormat = "#,##0.000"
Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = _
Val(tbVutrHl.value) * 1

Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value = Application.WorksheetFunction.Round(Range(cbShablon.value & "_" & tbVukon.value & "_oplx").value, 3)
End If






End Select

kbook.Worksheets("Формулы").Cells(1, 2) = Me.cbNTK
kbook.Worksheets("Формулы").Cells(2, 2) = tbVukon.value

kbook.Worksheets("Формулы").Cells(3, 2) = Me.cbLisn
kbook.Worksheets("Формулы").Cells(4, 2) = Me.tbNLK.value
kbook.Worksheets("Формулы").Cells(5, 2) = Me.cbDataDog.value
kbook.Worksheets("Формулы").Cells(6, 2) = Me.tbKV
kbook.Worksheets("Формулы").Cells(7, 2) = Me.tbVud
'kbook.Worksheets("Формулы").Cells(8, 2) = tbook.Range(z & "Mas").Value
kbook.Worksheets("Формулы").Cells(9, 2) = del
kbook.Worksheets("Формулы").Cells(10, 2) = drov
kbook.Worksheets("Формулы").Cells(11, 2) = Me.cbShablon
End With


Call CalculationA
If Val(Me.tbZalushok) <> 0 Then
Me.tbZalushokSum.value = Range("опл_" & cbShablon.value & "_" & tbVukon.value).value
Else
Me.tbSumDog.value = Range("опл_" & cbShablon.value & "_" & tbVukon.value).value
End If
kbook.Save
kbook.Close
Set kbook = mywbBook("Кошторис.xls", kbook_path_name)
Call CalculationM

FilenName = ThisWorkbook.Path & "\" & "Техкарты" & "\" & Me.cbNTK.value & "_Кошторис" & ".xls"
Call CopyAndReplaseFormul(kbook, cbShablon.value & " " & tbVukon.value, FilenName)

kbook.Save
kbook.Close
tbook.Activate
Me.CommandButton18.Enabled = True
End Sub

Private Sub CommandButton17_Click()
'Range("valCalend").ClearContents
'Range("FormCalend").Value = "Form"
'Range("ContrCalend").Value = "cbDataT"
'Call KalendarShow

'Me.Controls(Range("ContrCalend").Value).Value = VBA.CDate(MonthTK(VBA.Day(Range("valCalend").Value)) & "." & MonthTK(VBA.Month(Range("valCalend").Value)) & "." & VBA.Year(Range("valCalend").Value))
z = Get_Date1(Me.cbDataT, Now)
'Me.Controls(Range("ContrCalend").Value).Value = VBA.CDate(MonthTK(VBA.Day(z)) & "." & MonthTK(VBA.Month(z)) & "." & VBA.Year(z))

Me.cbDataT = z
End Sub

Private Sub CommandButton18_Click()
FilenName = ThisWorkbook.Path & "\" & "Техкарты" & "\" & Me.cbNTK.value & "_Кошторис" & ".xls"
Set kbook = Workbooks.Open(FilenName)
kbook.Worksheets(1).Activate
Call Печать1_1
kbook.Close

End Sub

Private Sub CommandButton19_Click()

End Sub

Private Sub CommandButton2_Click()
Me.Hide
End Sub

Private Sub CommandButton20_Click()
cmbTender№.Text = ""
End Sub

Private Sub CommandButton21_Click()
If Me.cbVidRubki = "" Then
MsgBox "Від рубки не выбран"
Me.cbMonth.SetFocus
Exit Sub
End If
If Me.cbLisn = "" Then
MsgBox "Лісництво не вибрано"
Me.cbLisn.SetFocus
Exit Sub
End If
If Me.tbKV = "" Then
MsgBox "Квартал не вибран"
Me.cbN.SetFocus
Exit Sub
End If

If Me.tbVud = "" Then
MsgBox "Видел не вибран"
Me.cbN.SetFocus
Exit Sub
End If


n = ID_ТК
Set r = Worksheets("Техкарты").Columns("AH:AH").Find(What:=n, LookAt:=xlWhole)

If r Is Nothing Then
MsgBox "Техкарта с  номером " & n & " не найдена"
Else
Worksheets("Техкарты").Activate
r.Activate
End If
Dim loj As clsmListObjs
 Dim wb As Workbook
 Set wb = ThisWorkbook
Set loj = New clsmListObjs
With loj
.Initialize wb
.ActiveListObjectRowDelete

End With
End Sub

Private Sub CommandButton3_Click()

Call OpenTexKart1

Call РасчитатьТехКарту
'If Form.cbShablon.Text <> "" Then
'Worksheets(Form.cbShablon.Text).Activate
Call ЗаполнитьФормуТехкарты(ActiveSheet.name) 'Worksheets(Form.cbShablon.Text).cell(ActiveCell.row, 1)
'End If
'Worksheets(Form.cbShablon.Text).Activate
'Call ObnovitKart
End Sub

Private Sub CommandButton4_Click()
Call СохранитьТехКарту

Select Case Me.cbShablon.value

Case "РГК"
z = "g"
Case "РД"
z = "d"
Case "Молодняк"
z = "m"
Case Else

Exit Sub
End Select
Worksheets("Константы").Range("cNTK").value = Me.cbNTK.value
Worksheets("Константы").Range("cLK").value = Me.tbNLK.value
Worksheets("Константы").Range("cKV").value = Me.tbKV.value
Worksheets("Константы").Range("cVud").value = Me.tbVud.value

НоваяПапка = NewFolderName
ИмяФайла = NewFaleName
РасширениеСоздаваемыхФайлов = ".xls"
FilenName = НоваяПапка & "\" & ИмяФайла & РасширениеСоздаваемыхФайлов
'FilenName = ThisWorkbook.Path & "\" & "Техкарты" & "\" & Range(Z & "NTK") & ".xls"
Debug.Print FilenName
Call CopyAndReplaseFormul(ThisWorkbook, Me.cbShablon.value, FilenName)

Worksheets(Me.cbShablon.value).Activate
End Sub

Private Sub CommandButton5_Click()
Call РасчитатьТехКарту
Call RaschitatEnableTrue
Call CalculationA
Select Case Me.cbShablon.value

Case "РГК"
z = "g"
Case "РД"
z = "d"
Case "Молодняк"
z = "m"
Case Else

Exit Sub
End Select
Call CalculationM
tbVukon.value = Range(z & "Vukon")
tbK_Vukon = Range(z & "K_Vukon")
If tbK_Vukon = "1.05" Or tbK_Vukon = "1,05" Then
If z = "m" Then
Range(z & "Skrut1").Font.Color = vbBlack
Range(z & "Skrut2").Font.Color = vbBlack
Else
Range(z & "Skrut").Font.Color = vbBlack
End If











Else
If z = "m" Then
Range(z & "Skrut1").Font.Color = vbWhite
Range(z & "Skrut2").Font.Color = vbWhite

Else
Range(z & "Skrut").Font.Color = vbWhite
End If
'Range (z & "SumDog")
End If

End Sub

Private Sub CommandButton6_Click()
'For i = 1 To 3
'For j = 1 To 4
'Form.Controls("T" & j & i).Value = 0
'
'
'Next j
'Next i
'If Form.cbShablon.Value = "РД" Or Form.cbShablon.Value = "РГК" Then Form.Controls("tbNel").Value = 0


Call ОчиститьДанныеТехкарты
Me.T11.SetFocus
End Sub

Private Sub CommandButton9_Click()
If Me.cbMonth = "" Then
MsgBox "Месяц не выбран"
Me.cbMonth.SetFocus
Exit Sub
End If
If Me.cbLisn = "" Then
MsgBox "Лісництво не вибрано"
Me.cbLisn.SetFocus
Exit Sub
End If
If Me.cbN = "" Then
MsgBox "Номер не вибран"
Me.cbN.SetFocus
Exit Sub
End If
n = НомерТК
Set r = Worksheets("Техкарты").Columns("AE:AE").Find(What:=n, LookAt:=xlWhole)

If r Is Nothing Then
MsgBox "Техкарта с таким номером не найдена"
Else
Worksheets("Техкарты").Activate
r.Activate
End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Range("умДатаТарифы").value = Me.DTPicker1.value
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame6_Click()

End Sub

Private Sub Label18_Click()
If Me.Height = 57 Then Me.Height = 561 Else Me.Height = 57
End Sub

Private Sub Label40_Click()

End Sub

Private Sub SpinButton1_Change()

End Sub

Private Sub SpinButton1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub SpinButton1_SpinDown()
If Me.cbLisn.value <> "" Then
Set r = Range("ЛісВа").Find(Me.cbLisn.value)
If r Is Nothing Then
Exit Sub
Else
'SpinButton1.ControlSource = "=" & r.Address


r.Offset(columnoffset:=1).value = r.Offset(columnoffset:=1).value - 1
Call CalculationA
Call CalculationM
Me.cbNTK.value = Range(LeterSheet & "NTK").value
End If
End If
End Sub

Private Sub SpinButton1_SpinUp()
If Me.cbLisn.value <> "" Then
Set r = Range("ЛісВа").Find(Me.cbLisn.value)
If r Is Nothing Then
Exit Sub
Else
'SpinButton1.ControlSource = "=" & r.Address


r.Offset(columnoffset:=1).value = r.Offset(columnoffset:=1).value + 1
Call CalculationA
Call CalculationM
Me.cbNTK.value = Range(LeterSheet & "NTK").value
End If
End If
End Sub

Private Sub T11_Change()

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

Private Sub T43_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Form.cbShablon.value = "РД" Or Form.cbShablon.value = "РГК" Then
Form.tbNel.SetFocus
Else
CommandButton12.SetFocus
End If
End Sub

Private Sub tbMas_Change()

End Sub

Private Sub tbNel_Change()

End Sub

Private Sub tbPlocha_Change()

End Sub

Private Sub tbPlocha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
txt = Me.tbPlocha ' читаем текст из поля (для недопущения ввода двух и более запятых)
    Select Case KeyAscii
        Case 13: Me.tbPlocha = txt ' нажат Enter
        Case 8: ' нажат Backspace - ничего не делаем
        ' если запятая уже есть - отменяем ввод символа ' заменяем при вводе точку на запятую
        Case 44, 46: KeyAscii = IIf(InStr(1, txt, ",") > 0, 0, 44)
       Case 48 To 57:
      KeyAscii = KeyAscii    'если введена цифра'And Len(Txt) - InStrRev(Txt, ",") > 1
        Case 95:
        KeyAscii = IIf(InStr(1, txt, "_") > 0, 0, KeyAscii)   'если введена цифра'And Len(Txt) - InStrRev(Txt, ",") > 1
        Case Else:   KeyAscii = 0 ' иначе отменяем ввод символа
    End Select
End Sub

Private Sub tbSumDog_Change()
z = LeterSheet & "SumDog"
a = tbSumDog.value
'Range(z).Value = tbSumDog.Value
End Sub

Private Sub tbVud_Change()

End Sub

Private Sub tbVud_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
txt = Me.tbVud ' читаем текст из поля (для недопущения ввода двух и более запятых)
    Select Case KeyAscii
        Case 13: Me.tbVud = txt ' нажат Enter
        Case 8: ' нажат Backspace - ничего не делаем
        ' если запятая уже есть - отменяем ввод символа ' заменяем при вводе точку на запятую
        Case 44, 46: KeyAscii = IIf(InStr(1, txt, "_") > 0, 0, 95)
       Case 48 To 57:
      KeyAscii = KeyAscii    'если введена цифра'And Len(Txt) - InStrRev(Txt, ",") > 1
        Case 95:
        KeyAscii = IIf(InStr(1, txt, "_") > 0, 0, KeyAscii)   'если введена цифра'And Len(Txt) - InStrRev(Txt, ",") > 1
        Case Else:   KeyAscii = 0 ' иначе отменяем ввод символа
    End Select
End Sub

Private Sub tbVukon_Change()

End Sub

Private Sub tbVutrKbm_Change()

End Sub

Private Sub tbZalushok_Change()

End Sub

Private Sub tbZalushokPlosha_Change()

End Sub

Private Sub tbZalushokPlosha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
txt = Me.tbZalushokPlosha ' читаем текст из поля (для недопущения ввода двух и более запятых)
    Select Case KeyAscii
        Case 13: Me.tbZalushokPlosha = txt ' нажат Enter
        Case 8: ' нажат Backspace - ничего не делаем
        ' если запятая уже есть - отменяем ввод символа ' заменяем при вводе точку на запятую
        Case 44, 46: KeyAscii = IIf(InStr(1, txt, ",") > 0, 0, 46)
       Case 48 To 57:
      KeyAscii = KeyAscii    'если введена цифра'And Len(Txt) - InStrRev(Txt, ",") > 1
        Case 95:
        KeyAscii = IIf(InStr(1, txt, "_") > 0, 0, KeyAscii)   'если введена цифра'And Len(Txt) - InStrRev(Txt, ",") > 1
        Case Else:   KeyAscii = 0 ' иначе отменяем ввод символа
    End Select
End Sub

Private Sub tbZalushokSum_Change()

End Sub

Private Sub UserForm_Activate()
ThisWorkbook.Activate
hWnd = GetActiveWindow

WndStyle = GetWindowLong(hWnd, GWL_STYLE)

WndStyle = SetWindowLong(hWnd, GWL_STYLE, WndStyle Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Dim MyDate, DataText
ThisWorkbook.Activate
Dim i, q, ats
'Call CalculationM
Me.mp1.value = 0
Me.cbMonthPR.Clear
Me.cbMonth.Clear
For i = 1 To 12
Me.cbMonth.AddItem i
Me.cbMonthPR.AddItem i
Me.cbMonthPR.value = VBA.Month(VBA.Now())

'Me.cbShablon.
Next i
Me.cbN.Clear
For i = 1 To 20

Me.cbN.AddItem i

'Me.cbShablon.
Next i

Me.cbK_Skladanny.Clear
For i = 0 To 1 Step 0.1

Me.cbK_Skladanny.AddItem i

'Me.cbShablon.
Next i

cbSpRubki.Clear
cbSpRubki.AddItem 0
cbSpRubki.AddItem 1




'Заполним дати тарифних ставок
Me.cbDataT.Clear
q = LastRow("Тарифи")
For i = 2 To q
MyDate = Worksheets("Тарифи").Cells(i, 1)
DataText = VBA.CDate(MonthTK(VBA.Day(MyDate)) & "." & MonthTK(VBA.Month(MyDate)) & "." & VBA.Year(MyDate))
Me.cbDataT.AddItem VBA.Format(DataText, "dd-mm-yyyy")
Next i
'Установим актуальные тарифные ставки
ats = Application.WorksheetFunction.max(Range(Worksheets("Тарифи").Cells(2, 1), Worksheets("Тарифи").Cells(q, 1)))
Range("умДатаТарифы") = VBA.Format(ats, "dd-mm-yyyy")
Me.Height = 57
Me.Top = 2.25
cbPodcherk.AddItem "_"
cbPodcherk.AddItem ""
cbPodcherkN.AddItem "Подчеркнуть"
cbPodcherkN.AddItem ""
Call RaschitatEnableFalse
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Call CalculationA
End Sub
 
Function LeterSheet()
With Form
Select Case .cbShablon.Text

Case "РГК"
LeterSheet = "g"
Case "РД"
LeterSheet = "d"
Case "Молодняк"
LeterSheet = "m"
End Select
End With
End Function

Sub RaschitatEnableFalse()
Me.CommandButton4.Enabled = False
Me.CommandButton1.Enabled = False
End Sub
Sub RaschitatEnableTrue()
Me.CommandButton4.Enabled = True
Me.CommandButton1.Enabled = True
End Sub
 
