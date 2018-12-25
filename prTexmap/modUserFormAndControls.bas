Attribute VB_Name = "modUserFormAndControls"
'http://www.ozgrid.com/VBA/control-loop.htm
'http://www.mrexcel.com/forum/excel-questions/557158-visual-basic-applications-getting-control-type-please-help.html
 Private NotUse As Boolean   ' switch
 Private Arr()               ' array

Dim mufZVIT As Object


Function ShowUFLink(UFName, mdl)
Dim muf As Object
Dim muf1 As Object
Set muf = OpenUserForms(UFName)
muf.Show (mdl)
Set ShowUFLink = muf
End Function
Sub ShowUF()
Dim muf As Object
Dim muf1 As Object
Set mufZVIT = ShowUFLink("UF_Shapes", 0)
'UF.Show 0
End Sub
Function OpenUserForms(UFName)
    Dim VBComp As Object
    Dim muf As Object
    For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
        If VBComp.Type = 3 Then '3 = vbext_ct_MSForm
        If UFName = VBComp.name Then
        Set OpenUserForms = VBA.UserForms.Add(VBComp.name)
        Exit Function
            End If
        End If
    Next
End Function
Function UserFormControl(muf, contrName)
Dim mufcntr As Object
For Each mufcntr In muf.Controls
If contrName = mufcntr.name Then
        Set UserFormControl = mufcntr
        Exit Function
            End If
Next
End Function

Sub AllChekenUnCekenControls(contr, zn As Boolean)
For li = 0 To contr.ListCount - 1
 contr.Selected(li) = zn
Next li
End Sub
Sub OpenSelectedBook(muf, contrName)
 Dim wb As Workbook
 Set contr = UserFormControl(muf, contrName)
 
'  Application.UpdateLinks (False)
 z = mySelectedCount(muf, contrName)
If Not z = 0 Then
 For li = 0 To contr.ListCount - 1

Application.AskToUpdateLinks = False
If contr.Selected(li) Then
mUpdateLinks = 0
If Not IsBookOpen(contr.List(li)) Then
 Workbooks.Open fileName:=contr.List(li), UpdateLinks:=0
 End If
 End If
Next li
Application.AskToUpdateLinks = True
Else
Call MsgBox("Вы не выбрали не одного файла!", vbExclamation, "Внимание")

End If
 End Sub
Sub AddNameOpenBook(muf, contrName)
 Dim wb As Workbook
 Set contr = UserFormControl(muf, contrName)
' Application.AskToUpdateLinks = False
contr.Clear
   For Each wb In Workbooks
   contr.AddItem wb.FullName
   Next
 
'Application.AskToUpdateLinks = True
 End Sub
 
 Function mySelectedCount(muf, contrName)
 
 
 Dim intIndex As Integer
    Dim intCount As Integer
  Set contr = UserFormControl(muf, contrName)
    With contr
        For intIndex = 0 To .ListCount - 1
            If .Selected(intIndex) Then intCount = intCount + 1
        Next
    End With
  mySelectedCount = intCount
    
End Function
Sub myControlAdPar(muf, contrName, col, Par)

Dim obj As Object
Set contr = UserFormControl(muf, contrName)

contr.Clear
For Each obj In col
 v = GetCallByName(obj, Par, VbGet)
contr.AddItem v

Next
End Sub
 Sub ControlAddListVisibleListObjectColumnValue(muf, contrName, ListObjName, ListObjColumnName, Optional mIsVisible As Boolean = False, Optional mIsUnik As Boolean = True, Optional mAddZerro As Boolean = True, Optional mAddAll As Boolean = False)
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
Set mufcntr = UserFormControl(muf, contrName)
mufcntr.Clear
With loj
.Initialize wb
If ListObjName Like "*[!_0-9.]*" Then
Set lop = .items(ListObjName)

If mIsVisible Then
Set v = lop.ListColumns(ListObjColumnName).DataBodyRange.SpecialCells(xlCellTypeVisible)
Else
Set v = lop.ListColumns(ListObjColumnName).DataBodyRange
End If
If mAddAll Then mufcntr.AddItem "(Все)"

Dim elem
     On Error Resume Next
     With CreateObject("Scripting.Dictionary")
         .compaode = vbTextCompare
         For Each elem In v
    If VarType(elem) = vbDate Then elem = str(elem)
    If VarType(elem) = VBA.vbDouble Then elem = str(elem)
             If VarType(elem) = vbString Then elem = Trim(elem)
             
             
             
             If Not IsError(elem) Then
                 If Len(elem) > 0 Then .Add CStr(elem), elem
             End If
         Next
         Arr = .items
     End With
'     With Me
         Call QuickSortNonRecursive(Arr)
'         .Width = getFormWidth(DLLSheetSettings.Range("F59").value)
'         .ComboBox1.Width = .Width - 10
'         .Caption = getFormCaprion(DLLSheetSettings.Range("F41").value) & UBound(Arr) + 1
'         If UBound(Arr) Then .ComboBox1.DropDown
'         Call SetFormPosition(Me, DDLCell)
         mufcntr.List = Arr
'     End With


















'If mIsUnik Then
'z = UnicumRange(v)
'For w = 0 To UBound(z)
'
'
'mufcntr.AddItem z(w)
'
'Next
'
'Else
'For Each rf In v.Cells
'
'
'mufcntr.AddItem rf.value
'
'Next
'End If
Else
Dim f
f = Val("0.1")
startc = Val(Substring(ListObjName, "_", 1))
endc = Val(Substring(ListObjName, "_", 2))
stepc = Val(Substring(ListObjName, "_", 3))
If startc > endc Then stepc = steps * -1
For i = startc To endc Step stepc
mufcntr.AddItem i
Next i
End If
If mAddZerro Then mufcntr.AddItem ""

End With


 
 
 
 
 End Sub

Sub ControlAddListVisibleListObjectColumnValueAll(muf, cntrlname, ReportName)
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

Set mufcntr = UserFormControl(muf, cntrlname)

With loj
.Initialize wb

Set l = .ValueListObject("LinkControl", "contrName", "SaveValueLoName", cntrlname)
Set c = .ValueListObject("LinkControl", "contrName", "SaveValueLoClName", cntrlname)
Set v = .ValueListObject(l.value, "Report_Name", c.value, ReportName)
v.value = mufcntr.value
End With
End Sub










 Function ControlAddalue(muf, contrName, ListObjName, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)
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
Set mufcntr = UserFormControl(muf, contrName)
'mufcntr.Clear
With loj
.Initialize wb
Set lop = .items(ListObjName)
Set rValue = .ValueListObject(ListObjName, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)
If Not rValue Is Nothing Then
 Select Case TypeName(rValue.value)
Case "Date"
rValueV = VBA.CDate(VBA.Day(rValue.value) & "." & VBA.Month(rValue.value) & "." & VBA.Year(rValue.value))
If Not rValue Is Nothing Then muf.Controls(contrName).value = VBA.str(rValue.value)
Case Else
rValueV = rValue.value
If Not rValue Is Nothing Then mufcntr.value = rValueV
End Select




Set ControlAddalue = rValue
End If
End With


 
 
 
 
 End Function

 Function ControlValueAddListObjName(muf, contrName, Report_Name)

Dim ADO As New ADO
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
Set mufcntr = UserFormControl(muf, contrName)
'mufcntr.Clear
With loj
.Initialize wb
'Set lop = .Items(ListObjName)
'W = " WHERE "
'W = W & "(Report_Name='" & Report_Name & "') AND " & "(contrName='" & contrName & "')"
'W = " WHERE Report_Name=" & Report_Name

'ADO.Header = True
sQuery = "SELECT * FROM [" & "LinkControl" & "$]" ' & W
Debug.Print sQuery
ADO.Query (Query)

Set oRS = ADO.Recordset

oRS.MoveFirst
'oRS.Fields("SaveValueLoName").value
'oRS.Fields("SaveValueLoName").value

Set rValue = .ValueListObject(oRS.Fields("SaveValueLoName").value, "Report_Name", oRS.Fields("SaveValueLoName").value, Report_Name)
'Set rListObjName = .ValueListObject("LinkControl", "contrName", "SaveValueLoName", contrName)
'Set rИмяСтолбцаПоиска = .ValueListObject("LinkControl", "contrName", "SaveValueLoName", contrName)

rValue = mufcntr.value

'Set ControlAddalue = rValue

End With


 
 
 
 
 End Function
Function ControlValueAddListObject(wb As Workbook, muf, contrName, Index As Variant, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска, value)


Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs
'Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
'Set wb = ActiveWorkbook
Dim mufcntr As Object
Set loj = New clsmListObjs
'Set muf = OpenUserForms(UFName)
Set mufcntr = UserFormControl(muf, contrName)
'mufcntr.Clear
With loj
.Initialize wb
Set w = .ValueListObject(Index, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)
If Not w Is Nothing Then
w.value = value
Set ControlValueAddListObject = w
Else
Set ControlValueAddListObject = Nothing
End If



End With

End Function
























Sub MakeForm()
    Dim TempForm As Object    'VBComponent
    Dim NewButton As MSForms.CommandButton
    Dim line As Integer
    Dim TheForm
    Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs
Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
Dim vWidth()
Set wb = ActiveWorkbook
Dim mufcntr As Object
Set loj = New clsmListObjs
'Set muf = OpenUserForms(UFName)
'Set mufcntr = UserFormControl(muf, contrName)
'mufcntr.Clear
With loj
.Initialize wb
Set lop = .items(.ActiveListObject)
Set v = .ActiveListObjectHederRowRange
'
'утв цшер
'    Application.VBE.MainWindow.visible = False
    ' ???????? ??????????? ???? UserForm
    Set TempForm = ThisWorkbook.VBProject. _
                   VBComponents.Add(3)    'vbext_ct_MSForm
    With TempForm
        .Properties("Caption") = lop.name
        .Properties("Width") = 200
        .Properties("Height") = 300
    End With
    ReDim vWidth(v.Count, 2)
    For i = 20 To v.Count * 20 Step 20
    z = i / 20
                    k = 5
      If maxLabel < Len(v.Cells(z).value) * k Then maxLabel = Len(v.Cells(z).value) * k
'       maxLabel = Len(v.Cells(z).value) * k
'Debug.Print Len(v.Cells(z).value) * k
    Next i
    Debug.Print "______________"
   Debug.Print maxLabel
   Debug.Print "______________"
    For i = 20 To v.Count * 20 Step 20
    Set NewLabel = TempForm.Designer.Controls _
                    .Add("forms.label.1")
                    z = i / 20
                    k = 5
           vWidth(z - 1, 0) = Len(v.Cells(z).value) * k
           
    With NewLabel
        .Caption = v.Cells(z).value
        .Left = 10
        .Top = i
        .Width = maxLabel
         vWidth(z - 1, 0) = Len(v.Cells(z).value) * k + .Width + 10
         Debug.Print vWidth(z - 1, 0)
    End With
  TempForm.Properties("Height") = z * 20 + 10
    WithCMB = 150
    Set NewCombobox = TempForm.Designer.Controls _
                    .Add("forms.ComboBox.1")
    With NewCombobox
'        .Caption = "Test"
        .Left = maxLabel + 20
        .Top = i
        .Width = WithCMB
'    Call ControlAddListVisibleListObjectColumnValue(TempForm.Name, NewCombobox.Name, lop.Name, v.Cells(z).value, , , False)
    End With
    
    
    
    
    
    
    
    
    Next
    ' ?????????? ???????? ?????????? CommandButton
    Set NewButton = TempForm.Designer.Controls _
                    .Add("forms.CommandButton.1")
    With NewButton
        .Caption = "??????!"
        .Left = 60
        .Top = v.Count * 20 + 10 + 10
    End With

    With TempForm.tCodeModule
    
'  "Public collControls As Collection"
'"Private cMultipleControls As clsMultipleControls"
'
'
'
'"Private Sub UserForm_Activate()"
'"Dim ctl As MSForms.Control"
'"Dim i As Long"
'"Set collControls = New Collection"
'"For Each ctl In Me.Controls"
'    "Set cMultipleControls = New clsMultipleControls"
'    "Set cMultipleControls.ctl = ctl"
'    "collControls.Add cMultipleControls"
'"Next ctl"
'"End Sub"
    
    
    
    
    
    
    
    
    
    
        line = .CountOfLines
        .InsertLines line + 1, "Sub CommandButtonl_Click()"
        .InsertLines line + 2, "MsgBox ""??????!"""
        .InsertLines line + 3, "Unload Me"
        .InsertLines line + 4, "End Sub"
    End With
   TempForm.Properties("Height") = v.Count * 20 + 70
   TempForm.Properties("Width") = maxLabel + WithCMB + 30
  End With
    
    ' ??????????? ??????????? ???? UserForm
    VBA.UserForms.Add(TempForm.name).Show
    ' ???????? ??????????? ???? UserForm
    ThisWorkbook.VBProject.VBComponents.Remove TempForm
End Sub
Sub testisValidDatePeriod(txtEndDate, txtStartDate, msg As Boolean)
Dim r As Boolean
r = isValidDatePeriod(txtEndDate, txtStartDate, msg)
If r Then
MsgBox "Период дат выбран верно!", vbOKOnly, "Период дат!"
 
Else
Exit Sub
End If
End Sub

Function isValidDatePeriod(txtEndDate, txtStartDate, Optional msg As Boolean = False)






'Ошибка проверки - дата начала не может быть позднее, чем дата окончания
EndDate = Right(txtEndDate, 4) & Mid(txtEndDate, 4, 2) & Left(txtEndDate, 2)
StartDate = Right(txtStartDate, 4) & Mid(txtStartDate, 4, 2) & Left(txtStartDate, 2)
If EndDate < StartDate Then
If msg Then MsgBox "Дата начала не может быть позднее, чем Дата окончания!", vbOKOnly, "Ошибка в выборе периода дат!"
isValidDatePeriod = False
Exit Function
End If
isValidDatePeriod = True
End Function
Sub Заполнить_FilenamesCollection(muf, contrName, folder$, Шабл)
    On Error Resume Next
    Dim coll As Collection

 Set mufcntr = UserFormControl(muf, contrName)

mufcntr.Clear
'    folder$ = ThisWorkbook.Path & "\Платежи\"
    If Dir(folder$, vbDirectory) = "" Then
        MsgBox "Не найдена папка «" & folder$ & "»", vbCritical, "Нет папки ПЛАТЕЖИ"
        Exit Sub        ' выход, если папка не найдена
    End If
 
    Set coll = FilenamesCollection(folder$, Шабл)        ' получаем список файлов XLS из папки
    If coll.Count = 0 Then
'        MsgBox "В папке «" & Split(folder$, "\")(UBound(Split(folder$, "\")) - 1) & "» нет ни одного подходящего файла!", _
               vbCritical, "Файлы для обработки не найдены"
        Exit Sub        ' выход, если нет файлов
    End If
 
    ' перебираем все найденные файлы
    For Each file In coll
        mufcntr.AddItem file        ' выводим имя файла в окно Immediate
    Next
End Sub
Sub ЗаполнитьИменаЛистов(muf, contrName, FullPath$, Optional Шабл = "*")
Dim ADO As New ADO
Set mufcntr = UserFormControl(muf, contrName)
mufcntr.Clear

ADO.DataSource = FullPath$
ExistListObjectName = False
f = ADO.OpenSchema()
Set oRS = ADO.Recordset
 oRS.MoveFirst
Do While Not oRS.EOF

        tn = oRS("TABLE_NAME")
        If Left(tn, 1) = "'" Then tn = Mid(tn, 2)
        If Len(tn) < InStr(1, tn, "$") + 2 Then
            tn = Left(tn, InStr(1, tn, "$"))
            Debug.Print tn
If tn Like Шабл Then
mufcntr.AddItem Replace(tn, "$", "")
Exit Do
End If
        End If
       oRS.MoveNext
    Loop
    oRS.Close
End Sub





Sub UserFormLinkControl(muf)
Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs

Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
Dim lop_lr As ListRow
Dim lop_lc As ListColumn
Set wb = ThisWorkbook
Dim mufcntr As Object
Dim ps As PathSplitString
Set loj = New clsmListObjs


Dim bIsVisible As Boolean
Dim bIsUnik As Boolean
Dim bAddZerro As Boolean
Dim bEnable As Boolean
Dim bAddAll As Boolean
With loj
.Initialize wb
Set lop = .ListObjectActivate("LinkControl")

For Each v In lop.ListColumns("contrName").DataBodyRange
Set lop_lc = lop.ListColumns("FormName")
Set lop_lr = lop.ListRows(v.Row - lop.HeaderRowRange.Row)
Set s = Intersect(lop_lc.Range, lop_lr.Range)
s.Select
If s.value = muf.name Then
Debug.Print v.value
Set ListObjName = .ValueListObject("LinkControl", "contrName", "ListObjName", v.value)
Set ListObjColumnName = .ValueListObject("LinkControl", "contrName", "ListObjColumnName", v.value)
Set df = .ValueListObject("LinkControl", "contrName", "defValue", v.value)

bIsVisible = LogikVBA(.ValueListObject("LinkControl", "contrName", "bIsVisible", v.value))
bIsUnik = LogikVBA(.ValueListObject("LinkControl", "contrName", "bIsUnik", v.value))
bAddZerro = LogikVBA(.ValueListObject("LinkControl", "contrName", "bAddZerro", v.value))
bEnable = LogikVBA(.ValueListObject("LinkControl", "contrName", "bEnable", v.value))
bAddAll = LogikVBA(.ValueListObject("LinkControl", "contrName", "bAddAll", v.value))

ListObjColumnNameWhot = .ValueListObject("LinkControl", "contrName", "ListObjColumnNameWhot", v.value)
contrNameWhot = .ValueListObject("LinkControl", "contrName", "contrNameWhot", v.value)
bAddZerroWhot = LogikVBA(.ValueListObject("LinkControl", "contrName", "bAddZerroWhot", v.value))
bAddAllWhot = LogikVBA(.ValueListObject("LinkControl", "contrName", "bAddAllWhot", v.value))

If ListObjName.value <> "" Then Call ControlAddListVisibleListObjectColumnValue(muf, v.value, ListObjName.value, ListObjColumnName.value, bIsVisible, bIsUnik, bAddZerro, bAddAll)
Set mufcntr = UserFormControl(muf, v.value)
mufcntr.value = df.value
If bEnable Then mufcntr.Enabled = False
End If
Next


End With
End Sub

Sub CurrentPageSelectControl(muf, contrName, wb As Workbook, TableName, CurrentPageName)
'contrName = "cbLis_ufRealizaciya"
'TableName = "СводнаяРеализация"
'CurrentPageName = "Підрозділ"



Dim tableSheet As Worksheet
Dim ptj As clsmPivotTables
Dim pc As PivotCache
Dim pt As PivotTable
Dim pts As PivotTables
Dim pf As PivotField
Dim pi As PivotItem
Dim pis As PivotItem
Dim obj As Object
'Dim contrName




Set contr = UserFormControl(muf, contrName)
Set ptj = New clsmPivotTables
With ptj
.Initialize wb
'Debug.Print "Count PivotCaches-" & .Count_PivotCaches
Debug.Print ".Exists('СводнаяРеализация')-" & .Exists("СводнаяРеализация")
Set pt = .items("СводнаяРеализация")
Set tableSheet = pt.parent
tableSheet.Activate
Select Case contr.value
Case ""
Case "(Все)"
Call .CurrentPageSelect(TableName, CurrentPageName)
Case Else
Call .CurrentPageSelect(TableName, CurrentPageName, contr.value)
End Select
End With

End Sub
