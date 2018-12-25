VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fAkt 
   OleObjectBlob   =   "fAkt.frx":0000
   Caption         =   "UserForm2"
   ClientHeight    =   10170
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9510
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   414
End
Attribute VB_Name = "fAkt"
Attribute VB_Base = "0{493AE1D8-4A19-4ADC-A666-F40273384684}{22E85829-7D04-4644-9CEF-7C8FB0F91DB6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub cbKorectCall3_Click()
'[fВитрати_хлисти_мех] = [fВитрати_сорт_ручне]
'TextBox_ДилоВ = [fmЗнач2] + [fmЗнач3]
'TextBox_ДровВ = [fmЗнач5] + [fmЗнач6]
'TextBox_ХмизВ = [fmЗнач8] + [fmЗнач9]
End Sub

Private Sub CheckBox_Печать_Click()

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox_з_ПДВ_Click()
Set r = FindAll(Worksheets("Таблица Щоденник").Cells, "ПДВ")
If Me.CheckBox_з_ПДВ Then
r.Offset(columnoffset:=7).value = 0.2
Else
r.Offset(columnoffset:=7).value = 0

End If
Call ЗаполнитьСуммами

End Sub

Private Sub ComboBox_Full_Dilyanka_Change()
Dim FullDil As String
FullDil = ComboBox_Full_Dilyanka.Text
v = МасивДилянок(FullDil)
Me.ComboBox_Квартал = v(0)
Me.ComboBox_Виділ = v(1)
Me.ComboBox_Ділянка = v(2)
Me.TextBox_Cокр_ЛісВА.Text = Range("fSokrLis").value
Dim twb As Workbook
Dim shMastera As Worksheet
Set twb = ThisWorkbook
Set shMastera = twb.Worksheets("Мастера")
'Dim ЗначениеУмнойТаблицы(ИмяЛиста, ИмяТаблицы, ИмяСтолбцаПоиска, ИмяСтолбцаЗначения, ЗначениеПоиска)

Me.TextBox_ДилоВ.value = ""
Me.TextBox_ДровВ.value = ""
Me.TextBox_ХмизВ.value = ""


End Sub

Private Sub ComboBox_Lisnuctvo_Change()

Dim twb As Workbook
Dim shPriymannya As Worksheet
Dim col As Collection
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("Приймання")
nrow = 9
erow = LastRow("Приймання")

Me.ComboBox_Full_Dilyanka.Clear
Set dil = shPriymannya.Cells.Find("Ділянка")
Set lis = shPriymannya.Cells.Find("Лісництво")
Set dp = shPriymannya.Cells.Find("Договір, Номер договору")
Me.ComboBox_Full_Dilyanka.Clear
Me.ComboBox_№ДП.Clear
For I = nrow To erow
If shPriymannya.Cells(I, lis.Column).value = Me.ComboBox_Lisnuctvo.value Then
Me.ComboBox_Full_Dilyanka.AddItem shPriymannya.Cells(I, dil.Column).value
z = Split(shPriymannya.Cells(I, dp.Column).value, ",")
Me.ComboBox_№ДП.AddItem z(1)
End If
Next I
Range("fЛісництво").value = Me.ComboBox_Lisnuctvo.value
Me.TextBox_Cокр_ЛісВА.Text = Range("fSokrLis").value

Me.ListBox_Мастера.RowSource = "=Мастера[" & Me.TextBox_Cокр_ЛісВА.value & "]"
Me.ComboBox_МастераВ.RowSource = "=Мастера[" & Me.TextBox_Cокр_ЛісВА.value & "]"

'D = FiltUT(ThisWorkbook, "Реестр_ЛК", "Реестр_ЛК", vf, vc)
'D = FiltUT(ThisWorkbook, "Техкарты", "Техкарты", vf, vc)

vf = Array("2")


vc = Array(fAkt.ComboBox_Lisnuctvo)
D = FiltUT(ThisWorkbook, "Техкарты", "Техкарты", vf, vc)

vf = Array("1")
vc = Array(fAkt.ComboBox_Lisnuctvo)
'D = FiltUT(ThisWorkbook, "Реестр_ЛК", "Реестр_ЛК", vf, vc)

End Sub

Private Sub ComboBox_№ДП_Change()
Me.ComboBox_№Акт.value = fAkt.TextBox_Cокр_ЛісВА & "\" & fAkt.ComboBox_Full_Dilyanka & "\" & VBA.Month(Me.TextBox_ДатаАкт)

End Sub

Private Sub ComboBox_Виділ_Change()

End Sub

Private Sub ComboBox_ВідШаблона_Change()

End Sub

Private Sub ComboBox_Квартал_Change()
'vf = Array("2", "8")
'If fAkt.ComboBox_Ділянка = 0 Then
'z = fAkt.ComboBox_Виділ
'Else
'z = fAkt.ComboBox_Виділ & "_" & fAkt.ComboBox_Ділянка
'End If
'
'vc = Array(fAkt.ComboBox_Lisnuctvo, fAkt.ComboBox_Квартал)
'D = FiltUT(ThisWorkbook, "Техкарты", "Техкарты", vf, vc)

'vf = Array("1", "4")
'vc = Array(fAkt.ComboBox_Lisnuctvo, fAkt.ComboBox_Full_Dilyanka)
'D = FiltUT(ThisWorkbook, "Реестр_ЛК", "Реестр_ЛК", vf, vc)
End Sub

Private Sub ComboBox_МастераВ_Change()
Worksheets("Форма").Range("fМастер_Вибір") = Me.ComboBox_МастераВ
Call CalculationA
Call CalculationM
End Sub

Private Sub ComboBox1_ДЛГ_Change()

End Sub

Private Sub CommandButton_ДатаАкт_Click()
 Call fDataShow
End Sub

Private Sub CommandButton_Мастер_Click()
Set twb = ThisWorkbook
Set shPriymannya = twb.Worksheets("Приймання")
Set shChoden = twb.Worksheets("Зведений щоденник")
nrow = 9
erow = LastRow("Приймання")
shPriymannya.Activate
Set dil = shPriymannya.Cells.Find("Ділянка")
Set lis = shPriymannya.Cells.Find("Лісництво")
Set mast = shPriymannya.Cells.Find("Майстер")
Set mast = shPriymannya.Cells.Find("Майстер", mast)
Set lk = shPriymannya.Cells.Find("Лісорубний квиток")
Set dp = shPriymannya.Cells.Find("Договір, Номер договору")
Set vik = shPriymannya.Cells.Find("Лісозаготівельник.Наименование")
'Me.ComboBox_Full_Dilyanka.Clear
For I = nrow To erow
If shPriymannya.Cells(I, lis.Column).value = Me.ComboBox_Lisnuctvo.value Then
 If shPriymannya.Cells(I, dil.Column).value = Me.ComboBox_Full_Dilyanka.value Then
   If shPriymannya.Cells(I, mast.Column).value <> "" Then
master = shPriymannya.Cells(I, mast.Column).value
Me.ListBox_Мастера.value = master
End If
z = Split(shPriymannya.Cells(I, lk.Column).value, " ")
Me.ComboBox_№ЛК.value = z(2)
Me.TextBox_ДатаЛК.Text = z(4)
'Range("f№ЛКДата").Value = z(4)
z = Split(shPriymannya.Cells(I, dp.Column).value, ",")
v = Split(Trim(z(0)), " ")
Me.ComboBox_№ДП.value = z(1)
Me.TextBox_ДатаДП.Text = v(3)

Me.ComboBox_Виконавець.value = shPriymannya.Cells(I, vik.Column).value
End If
End If
Next I

Me.Label_Log.Caption = "Вибирить мастера из списка ! После чего Нажмить кнопку Дата "

End Sub

Private Sub CommandButton1_Click()
If Me.CheckBox_Import_All Then
Else
Call ИмпортПоСтроке(Me.ComboBox_Report)
End If
Me.Label_Log.Caption = "Звіт успішно імпортовано"
End Sub

Private Sub CommandButton11_Click()
 Set ThisWorkbook.app = Application
End Sub

Private Sub CommandButton12_Click()
Call ИмпортИзТехкарти
End Sub

Private Sub CommandButton13_Click()
Dim shp As Worksheet
Dim lo As ListObject
Dim q As Range
Set wb = ThisWorkbook
Set sh = wb.Worksheets("Форма")
ИмяШаблона = sh.Range("fDot").value
ИмяФайлаШаблона = ИмяШаблона & ".dot"

РасширениеСоздаваемыхФайлов = ".doc"
ИмяТаблици = "Форма"
ИмяСтолбцаКлючи = "Заголовки"
ИмяСтолбцаПроверки = "Заголовки"
ИмяСтолбцаЗначений = "Значения"



Set ЗаголовокСтолбцаКлючи = FindAll(sh.Rows(1), ИмяСтолбцаКлючи)
Set ЗаголовокСтолбцаПроверки = FindAll(sh.Rows(1), ИмяСтолбцаПроверки)
Set ЗаголовокСтолбцаЗначения = FindAll(sh.Rows(1), ИмяСтолбцаЗначений)

  
    НоваяПапка = NewFolderName & Application.PathSeparator

 'sh.Activate
'  v = Split(Range("f№Акт").Value, "\")
'КВВИД = sh.Range("fПолнаяДіл").Value
'КВВИД = Replace(КВВИД, ".", "")
'            ИмяФайла = КВВИД & "-" & sh.Range("fDot").Value & "-" & sh.Range("f№ЛК").Value ' Trim$(MonthTK(Range("cMonth")) & "_" & СокрTK(Range("cLicVa")) & "_" & Range("сN"))
'            FileNameDoc = НоваяПапка & ИмяФайла & РасширениеСоздаваемыхФайлов
            
        ИмяФайла = NewFaleName
'            FileNameDoc = НоваяПапка & ИмяФайла & РасширениеСоздаваемыхФайлов
FileNameDoc = NewFileFullName(НоваяПапка, ИмяФайла, РасширениеСоздаваемыхФайлов)
            
            
If Me.ListBox_Печать.ListIndex = -1 Then
MsgBox "Не выбран вид документа для печати"
Exit Sub
End If
'If Me.ListBox_Печать Like "Додаток*" Then doc = Me.ListBox_Печать.ListIndex + 2 Else doc = Me.ListBox_Печать.ListIndex + 1
doc = Me.ListBox_Печать.ListIndex + 1
' Me.ListBox_Печать.Selected(doc-1) = True
Set shp = wb.Worksheets("Печать")
Set lo = shp.ListObjects("Печать")
Set q = lo.ListColumns("Закладка").DataBodyRange.Cells
Set r = ПоискСоСмещением(q, "D" & doc, 0, 4)
Me.TextBox_Копий = r.value
Call ПечатьWordизExcel(FileNameDoc, doc)

End Sub

Private Sub CommandButton14_Click()
Call ПечатьПоУмолчанию
End Sub

Private Sub CommandButton15_Click()
Dim twb As Workbook
Dim wb As Workbook
Dim NameTX As Range
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
Set twb = ThisWorkbook
Path = twb.Path & Application.PathSeparator
Set NameTX = ЗначениеУмнойТаблицы("Импорт", "Импорт", "Лист", "Файл для импорта", "Техкарты")
If fso.FileExists(Path & NameTX.value) Then
Set wb = Workbooks.Open(Path & NameTX.value)
Else
MsgBox "Файл-" & Path & NameTX.value & "не обнаружен"
End If
End Sub

Private Sub CommandButton_Импорт_Click()

Call ИмпортИзТехкарти
If Me.ComboBox_ВідШаблона = "Молодняки" Then
[fВитрати_хлисти_мех] = [fВитрати_сорт_ручне]
TextBox_ДилоВ = [fmЗнач2] + [fmЗнач3]
TextBox_ДровВ = [fmЗнач5] + [fmЗнач6]
TextBox_ХмизВ = [fmЗнач8] + [fmЗнач9]
[fСерОбХл] = ""
End If

Me.ComboBox_МастераВ = Worksheets("Форма").Range("fМастер_Вибір")
Me.ComboBox_Full_Dilyanka = Worksheets("Форма").Range("fПолнаяДіл")
If Me.ComboBox_ВідШаблона = "Молодняки" Then
'Me.TextBox_Розцінка.Value = Worksheets("Форма").Range("fВитрати_хлисти_мех")
Else
Sum = 0
'For i = 1 To 9 Step 4
'Sum = Sum + Worksheets("Форма").Range("Знач" & i)
'Next i
'For i = 2 To 10 Step 4
'Sum = Sum + Worksheets("Форма").Range("Знач" & i)
'Next i
For I = 1 To 6
Sum = Sum + Worksheets("Форма").Range("Знач" & I)
Next I

Me.TextBox_ДилоВ.value = Sum
Sum = 0
'For i = 3 To 11 Step 4
'Sum = Sum + Worksheets("Форма").Range("Знач" & i)
'Next i
'For i = 4 To 12 Step 4
'Sum = Sum + Worksheets("Форма").Range("Знач" & i)
'Next i
For I = 7 To 12
Sum = Sum + Worksheets("Форма").Range("Знач" & I)
Next I


Me.TextBox_ДровВ.value = Sum
If Worksheets("Форма").Range("fНелеквід") = "" Then
Me.TextBox_ХмизВ.value = 0
Else
Me.TextBox_ХмизВ.value = Worksheets("Форма").Range("fНелеквід")
End If


'Me.TextBox_ДилоВ.Value = Worksheets("Форма").Range("fДілової")
'Me.TextBox_ДровВ.Value = Worksheets("Форма").Range("fДровяної")
'Me.TextBox_ХмизВ.Value = Worksheets("Форма").Range("fХмизу")

'If Me.ComboBox_ВідШаблона = "Молодняк" Then
'Me.TextBox_Розцінка.Value = Worksheets("Форма").Range("fВитрати_хлисти_мех")
'Else
'Me.TextBox_Розцінка.Value = Worksheets("Форма").Range("fВитрати_сорт_ручне")
'End If


'Розцінка
Me.TextBox_ВартистьЗнебособленного = Worksheets("Форма").Range("fРозцінка")
Me.TextBox_ВартистьХлист = Worksheets("Форма").Range("fРозцінка_хл")
Me.TextBox_Ликвид = Worksheets("Форма").Range("fВсього__ліквідної_деревини")
If Val(Worksheets("Форма").Range("fНелеквід").value) = 0 Then
Worksheets("Форма").Range("fХмизу_Вибір").value = 0
Else
 Worksheets("Форма").Range("fХмизу_Вибір").value = Worksheets("Форма").Range("fНелеквід").value
 End If
 End If
Call UbdateSumm







End Sub

Private Sub CommandButton17_Click()
Call fff
End Sub

Private Sub CommandButton18_Click()
Worksheets("Техкарты").Activate

End Sub

Private Sub CommandButton19_Click()
Worksheets("Реест_ЛК").Activate

End Sub

Private Sub CommandButton2_Click()
ListBox_Разбивка.SetFocus
If ListBox_Разбивка.ListIndex = ListBox_Разбивка.ListCount - 1 Then
ListBox_Разбивка.ListIndex = 0
Else
ListBox_Разбивка.ListIndex = ListBox_Разбивка.ListIndex + 1
End If
End Sub

Private Sub CommandButton20_Click()
Me.TextBox_ДилоВ.value = 0
Me.TextBox_ДровВ.value = 0
Me.TextBox_ХмизВ.value = 0
End Sub

Private Sub CommandButton21_Click()
'ThisWorkbook.Path & "\Техкарта " & Me.TextBox_Год_Техкарта & "." & Me.ComboBox_Расширение_Техкарта

Set b = mywbBook("Техкарта " & Me.TextBox_Год_Техкарта & "." & Me.ComboBox_Расширение_Техкарта, ThisWorkbook.Path & "\")
If b Is Nothing Then MsgBox ("Файл " & twb.Path & "\" & EOBookName)
ThisWorkbook.Save
ThisWorkbook.Close
End Sub

Private Sub CommandButton22_Click()
Call Запрос_Техкарта
lr = LastRow("Техкарты")
Worksheets("Техкарты").Cells(lr, 10).Select
Me.MultiPage1.value = Me.MultiPage1.value + 1
Me.CommandButton_Импорт.value = True
CommandButton3.value = True
End Sub

Private Sub CommandButton23_Click()
D = FiltUT(ThisWorkbook, "Техкарты", "Техкарти", vf, vc)
End Sub

Private Sub CommandButton3_Click()
Dim shp As Worksheet
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject


 
 NameSheet = "Техкарты"
 
 
 
 Me.ListBox_Печать.value = "ДОГОВІР №"
 
'Call CalculationA
'Call CalculationM
'If ActiveSheet.Name <> NameSheet Then Exit Sub
Set sh = Worksheets(NameSheet)

If Selection.Cells.Count = 1 Then
'Call ИмпортИзТехкарти
Me.Label57.Caption = СформироватьДоговоры
If fso.FileExists(FileNameDoc) Then Call Добавим_Гиперсылку(Worksheets("Форма"), Worksheets("Форма").Range("fDot"), FileNameDoc, Worksheets("Форма").Range("fDot").value) Else Worksheets("Форма").Range("fDot").Hyperlinks.Delete
Worksheets(NameSheet).Activate
If fso.FileExists(FileNameDoc) Then Call Добавим_Гиперсылку(Worksheets(NameSheet), Worksheets(NameSheet).Cells(Selection.Row, 46), FileNameDoc, FileNameDoc) Else Worksheets(NameSheet).Cells(Selection.Row, 46).Hyperlinks.Delete
Worksheets("Форма").Activate
' FilenName = ThisWorkbook.Path & "\" & "Техкарты" & "\" & Range("m" & "NTK") & ".xls"
doc = Me.ListBox_Печать.ListIndex
If CheckBox_Печать Then Call ПечатьWordизExcel(FileNameDoc, doc)
MsgBox FileNameDoc
Else
Set r = Selection
Dim z As Range
For Each z In r.Cells
z.Activate
'Call ИмпортИзТехкарти


Call СформироватьДоговоры

If fso.FileExists(FileNameDoc) Then Call Добавим_Гиперсылку(Worksheets("Форма"), Worksheets("Форма").Range("fDot"), FileNameDoc, Worksheets("Форма").Range("fDot").value) Else Worksheets("Форма").Range("fDot").Hyperlinks.Delete
If fso.FileExists(FileNameDoc) Then Call Добавим_Гиперсылку(Worksheets(NameSheet), Worksheets(NameSheet).Cells(Selection.Row, 46), FileNameDoc, FileNameDoc) Else Worksheets(NameSheet).Cells(Selection.Row, 46).Hyperlinks.Delete
' FilenName = ThisWorkbook.Path & "\" & "Техкарты" & "\" & Range("m" & "NTK") & ".xls"
doc = Me.ListBox_Печать.ListIndex
If Me.ListBox_Печать.ListIndex = -1 Then
MsgBox "Не выбран вид документа для печати"
Exit For
End If
doc = Me.ListBox_Печать.ListIndex + 1
Set wb = ThisWorkbook
Set shp = wb.Worksheets("Печать")
Set lo = shp.ListObjects("Печать")
Dim q As Range
Set q = lo.ListColumns("Закладка").DataBodyRange.Cells
Set r = ПоискСоСмещением(q, "D" & doc, 0, 4)
Me.TextBox_Копий = r.value









If CheckBox_Печать Then Call ПечатьWordизExcel(FileNameDoc, doc)
MsgBox FileNameDoc
Set sh = wb.Worksheets(NameSheet)
sh.Activate
Next
End If
End Sub

Private Sub CommandButton4_Click()
Dim EOBookName As String
Dim pth
pth = ThisWorkbook.Path
EOBookName = Me.ComboBox_Lisnuctvo & " " & Me.ComboBox_Full_Dilyanka & ".xls"
'If WorkbookExist(pth, EOBookName) Then
Call ИмпортЩоденника(EOBookName)
Me.MultiPage1.value = 1
'Else
'Call MsgBox("Файл-" & vbLf & EOBookName & vbLf & "по пути " & vbLf & pth & vbLf & "не найден" & vbLf & "сформируйте его или переместите в папку с программой", vbCritical, "Файл не найден")
'
'End If
Me.Label_Log.Caption = "Импорт щоденника завершен!  Нажмить кнопку Найти "
End Sub

Private Sub CommandButton5_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub CommandButton6_Click()
Select Case Me.MultiPage1.value
Case 0
Me.MultiPage1.value = Me.MultiPage1.value + 1
Me.Label_Log.Caption = "Внесить початкові данні в поля віделеніе жолтім цветом или вібирите из списка в полях віделеніх оранжевім цветом та нажмить кнопку Видкрити "
ThisWorkbook.Worksheets("Форма").Activate
Case 1
Me.MultiPage1.value = Me.MultiPage1.value + 1
Me.Label_Log.Caption = "Внесить значення обьема для соответствующей позиции списка в поле віделеніе жовтимм цветом Для потверждения ввода нажимайта кнопку Ок нажмить кнопку Видкрити "
ThisWorkbook.Worksheets("Таблица Щоденник").Activate
Case 2
Me.MultiPage1.value = Me.MultiPage1.value + 1
Case 3
Me.MultiPage1.value = 0
End Select
End Sub

Private Sub CommandButton7_Click()
Me.Hide
End Sub

Private Sub CommandButton8_Click()
Call ЗаполнитьСуммами
Range("fТабл").Activate
Call InsertOrEditTableLink
End Sub

Private Sub CommandButton9_Click()
Call Запрос_Техкарта
End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label39_Click()

End Sub

Private Sub Label56_Click()

End Sub

Private Sub ListBox_Мастера_Change()
Range("fМастер").value = Me.ListBox_Мастера.value
Me.Label_Log.Caption = "Нажмить кнопку Дата "
End Sub

Private Sub ListBox_Мастера_Click()

End Sub

Private Sub ListBox_Печать_Click()

End Sub

Private Sub ListBox_Разбивка_Change()
Dim r As Range
Set r = Worksheets("Таблица Щоденник").Columns(10).Find(ListBox_Разбивка.value)
Me.TextBox_Разбивка_Обьем.ControlSource = r.Offset(columnoffset:=-3).address
Me.TextBox_Разбивка_Обьем.SetFocus

Me.TextBox_Разбивка_Обьем.SelStart = 0
Me.TextBox_Разбивка_Обьем.SelLength = Len(Me.TextBox_Разбивка_Обьем)
End Sub

Private Sub ListBox_Разбивка_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub



Private Sub TextBox_ВартистьЗнебособленного_Change()
'Розцінка
Worksheets("Форма").Range("fВитрати_сорт_ручне").value = Me.TextBox_ВартистьЗнебособленного.value
Call CalculationA
Call CalculationM
Call UbdateSumm
End Sub

Private Sub TextBox_ВартистьХлист_Change()
'Розцінка

Worksheets("Форма").Range("fВитрати_хлисти_мех").value = Me.TextBox_ВартистьХлист.value
Call CalculationA
Call CalculationM
End Sub

Private Sub TextBox_ДатаЗакДП_Change()
Range("f№ЛКДатаЗак").value = TextBox_ДатаЗакДП.value
End Sub

Private Sub TextBox_ДилоВ_Change()
'If TextBox_ДилоВ.Value <> "" And Me.TextBox_ДровВ.Value <> "" Then Me.TextBox_ВсьогоДоговірВибір = (CDbl(Me.TextBox_ДилоВ.Value) + CDbl(Me.TextBox_ДровВ.Value)) * Val(Me.TextBox_Розцінка.Value)

Call UbdateSumm

End Sub



Private Sub TextBox_ДилоФ_Change()

End Sub

Private Sub TextBox_ДровВ_Change()
'If Me.TextBox_ДилоВ.Value <> "" And Me.TextBox_ДровВ.Value <> "" Then Me.TextBox_ВсьогоДоговірВибір = (CDbl(Me.TextBox_ДилоВ.Value) + CDbl(Me.TextBox_ДровВ.Value)) * Val(Me.TextBox_Розцінка.Value)
Call UbdateSumm
End Sub

Private Sub TextBox_ДровФ_Change()

End Sub

Private Sub TextBox_ПДВ_Change()

End Sub

Private Sub TextBox_Разбивка_Обьем_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

'Private Sub TextBox_Розцінка_Change()
'
''If Me.TextBox_ДилоВ.Value <> "" And Me.TextBox_ДровВ.Value <> "" Then Me.TextBox_ВсьогоДоговірВибір = (CDbl(Me.TextBox_ДилоВ.Value) + CDbl(Me.TextBox_ДровВ.Value)) * Val(Me.TextBox_Розцінка.Value)
'End Sub

Private Sub TextBox_ХмизВ_Change()
Call UbdateSumm
'Call CalculationA
'Call CalculationM
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox_ДатаЗакДП_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call fDataShow
End Sub

Private Sub TextBox_ХмизФ_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
ThisWorkbook.Activate
'For i = 1 To 31
'Me.ComboBox_Day.AddItem i
'Next i
'M = VBA.Month(Now())
'For i = 1 To 12
'Me.ComboBox_Mouth.AddItem i
'Next i
'If M = 1 Then
'Me.ComboBox_Mouth.Value = 12
'k = 12
'Else
'Me.ComboBox_Mouth.Value = M - 1
'k = M - 1
'End If
'd = Year(Now())
'For i = d - 4 To d + 4
'Me.ComboBox_Year.AddItem i
'Next i
'If M = 1 Then
'Me.ComboBox_Year.Value = d - 1
'z = d - 1
'Else
'Me.ComboBox_Year.Value = d
'z = d
'End If
'dn = ДнейВМесяце(k, z)
'Me.ComboBox_Day.Value = dn
Worksheets("Форма").Range("fДілової_Вибір") = 0
Worksheets("Форма").Range("fДровяної_Вибір") = 0
Worksheets("Форма").Range("fХмизу_Вибір") = 0



КнигиИмпорта = СписокЗначенийСтолбцаУмнойТаблицы("Импорт", "Импорт", "Новое Имя Листа")
Me.ComboBox_Report.list = КнигиИмпорта
Me.ComboBox_Report.value = "Приймання"
z = Me.ComboBox_Lisnuctvo
For I = 1 To Me.ComboBox_Lisnuctvo.ListCount - 1
Me.ComboBox_Lisnuctvo.ListIndex = I
If Me.ComboBox_Lisnuctvo.value <> z Then Exit For
Next I
Me.ComboBox_Lisnuctvo.value = z
Me.MultiPage1.value = 0
'ThisWorkbook.RefreshAll
Me.ComboBox_Путь_Техкарта.value = ThisWorkbook.Path & "\Техкарта " & Me.TextBox_Год_Техкарта & "." & Me.ComboBox_Расширение_Техкарта
End Sub
Sub ЗаполнитьСуммами()
With fAkt
Set r = FindAll(Worksheets("Таблица Щоденник").Cells, "Всього без ПДВ")
.TextBox_Сума.value = r.Offset(columnoffset:=8).value

Set r = FindAll(Worksheets("Таблица Щоденник").Cells, "ПДВ")
.TextBox_ПДВ.value = r.Offset(columnoffset:=8).value

Set r = FindAll(Worksheets("Таблица Щоденник").Cells, "Всього з ПДВ")
.TextBox_Всього.value = r.Offset(columnoffset:=8).value
End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Call CalculationA
'Call CalculationM
End Sub

Sub UbdateSumm()
Call CalculationA
Worksheets("Форма").Range("fДілової_Вибір") = Me.TextBox_ДилоВ.value
Worksheets("Форма").Range("fДровяної_Вибір") = Me.TextBox_ДровВ.value
Worksheets("Форма").Range("fХмизу_Вибір") = Me.TextBox_ХмизВ.value

Me.TextBox_ДилоФ.value = Worksheets("Форма").Range("fДілової")
Me.TextBox_ДровФ.value = Worksheets("Форма").Range("fДровяної")
Me.TextBox_ХмизФ.value = Worksheets("Форма").Range("fХмизу")
s = Replace(Worksheets("Форма").Range("fРозцінка"), ",", ".")
Me.TextBox_ВартистьЗнебособленного = s

Me.TextBox_ВартистьХлист = s

Me.TextBox_СумаДоговір.Text = Worksheets("Форма").Range("fСумаДоговір").value
Me.TextBox_ПДВДоговір.Text = Worksheets("Форма").Range("fПДВДоговір").value
Me.TextBox_ВсьогоДоговірВибір.Text = Worksheets("Форма").Range("fСумаДоговірЗПДВ").value

Me.TextBox_СумаДоговір_хл.Text = Worksheets("Форма").Range("fСумаДоговірЗПДВ").value
Me.TextBox_ПДВДоговір_хл.Text = Worksheets("Форма").Range("fПДВ_Договір_хл").value
Me.TextBox_ВсьогоДоговірВибір_хл.Text = Worksheets("Форма").Range("fСумаДоговірЗПДВ_хл").value


Call CalculationM
End Sub
