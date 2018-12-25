Attribute VB_Name = "Module1"
Type МестоХранения
Кв As String
Вид As String
Діл As String
End Type
Dim MX As МестоХранения
Function МесяцьПрописом(Номер_месяца)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(Номер_месяца)) Then
    With dict
  .Add "1", "Січень"
  .Add "2", "Лютий"
  .Add "3", "Березень"
  .Add "4", "Квітень"
  .Add "5", "Травень"
  .Add "6", "Червень"
  .Add "7", "Липень"
  .Add "8", "Серпень"
  .Add "9", "Вересень"
  .Add "10", "Жовтень"
  .Add "11", "Листопад"
  .Add "12", "Грудень"
  
  
  
    End With
    
    
    
End If

МесяцьПрописом = dict(Trim(str(Номер_месяца)))
End Function
Function МесяцяПрописом(Номер_месяца)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(Номер_месяца)) Then
    With dict
  .Add "1", "Січня"
  .Add "2", "Лютого"
  .Add "3", "Березня"
  .Add "4", "Квітня"
  .Add "5", "Травня"
  .Add "6", "Червня"
  .Add "7", "Липня"
  .Add "8", "Серпня"
  .Add "9", "Вересня"
  .Add "10", "Жовтня"
  .Add "11", "Листопада"
  .Add "12", "Грудня"
  
  
  
    End With
    
    
    
End If

МесяцяПрописом = dict(Trim(str(Номер_месяца)))
End Function
Function МесяціПрописом(Номер_месяца)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(Номер_месяца)) Then
    With dict
  .Add "1", "Січні"
  .Add "2", "Лютому"
  .Add "3", "Березні"
  .Add "4", "Квітні"
  .Add "5", "Травні"
  .Add "6", "Червні"
  .Add "7", "Липні"
  .Add "8", "Серпні"
  .Add "9", "Вересні"
  .Add "10", "Жовтні"
  .Add "11", "Листопаді"
  .Add "12", "Грудні"
  
  
  
    End With
    
    
    
End If

МесяціПрописом = dict(Trim(str(Номер_месяца)))
End Function
Function ДнейВМесяце(Номер_месяца, Номер_года)
Dim dict
Set dict = CreateObject("Scripting.Dictionary")
If Not dict.Exists(str(Номер_месяца)) Then
    With dict
  .Add "1", "31"
  If IsDate("02/29/" & Номер_года) = True Then
  
  .Add "2", "29"
 Else
 .Add "2", "28"
 End If
  .Add "3", "31"
  .Add "4", "30"
  .Add "5", "31"
  .Add "6", "30"
  .Add "7", "31"
  .Add "8", "31"
  .Add "9", "30"
  .Add "10", "31"
  .Add "11", "30"
  .Add "12", "31"
  
  
  
    End With
    
    
    
End If

ДнейВМесяце = dict(Trim(str(Номер_месяца)))
End Function

Function МасивДилянок(Ділянка As String)


Call РазборДілянкиЕО(Ділянка)
v = Array(MX.Кв, MX.Вид, MX.Діл)
МасивДилянок = v
End Function
Sub РазборДілянкиЕО(Ділянка As String)
'66 кв (1, 2 вид)  діл.
'79 кв (12, 13, 15, 16 вид)  діл.
'27 кв (5 вид) 0 діл.



Dim v
If Ділянка = "" Then Exit Sub


v = Split(Ділянка, "(")
z = Split(v(1), ")")
MX.Кв = Val(v(0))
If z(0) Like "*,*" Then
MX.Вид = Left(z(0), Len(z(0)) - 4)
Else
MX.Вид = Val(z(0))
End If

MX.Діл = Val(z(1))
'РазборДілянкиЕО = MX
End Sub

Function РазборДоговораЕО(Договор As String)
'від заготівлі по договору (наряд-акту)  № 09\Шп\1_9 від 02.09.2013 року
'Договор As String


Dim v

v = Split(Договор, "№")
z = Split(VBA.Trim(v(1)), " ")
'0-номер договора
'1-від
'2-Дата договора
'3-року
РазборДоговораЕО = z
End Function


Function ПоискДилянкиПоДоговору(Строка_с_Договором)
Dim Строка_с_Договором As String
Dim fDogovor As Range
Dim fDilyanka As Range
Dim columnDogovor As Range
Dim columnDilyanka As Range
Dim shDogovor As Worksheet
'Проверка на наличие листа Реестр Договорів
Set shDogovor = ThisWorkbook.Worksheets("Реестр_Договора")
mDogovor = РазборДоговораЕО(Строка_с_Договором)
Договор = mDogovor(0)
With shDogovor
Set fDogovor = .Cells.Find("Номер договору")
Set columnDogovor = .Columns(fDogovor.Column).Cells
Set fDilyanka = .Cells.Find("Ділянка")
Set columnDilyanka = .Columns(fDilyanka.Column).Cells
'Ділянка
Set fDogovor = .Cells.Find(Договор)
Set fDilyanka = Application.Intersect(columnDilyanka, .Rows(fDogovor.Row))
ПоискДилянкиПоДоговору = fDilyanka.value
End With
End Function

Sub fff()
vf = Array("2") ', "8", "9"
If fAkt.ComboBox_Ділянка = 0 Then
z = fAkt.ComboBox_Виділ
Else
z = fAkt.ComboBox_Виділ & "_" & fAkt.ComboBox_Ділянка
End If

vc = Array(fAkt.ComboBox_Lisnuctvo) ', fAkt.ComboBox_Квартал, z
D = FiltUT(ThisWorkbook, "Техкарты", "Техкарти", vf, vc)

'vf = Array("1", "4")
'vc = Array(fAkt.ComboBox_Lisnuctvo, fAkt.ComboBox_Full_Dilyanka)
'D = FiltUT(ThisWorkbook, "Реестр_ЛК", "Реестр_ЛК", vf, vc)

End Sub

Function FiltUT(twb As Workbook, ShName, UtNAme, vf, vc)
Dim sh As Worksheet
Dim lo As ListObject
On Error Resume Next
Set lo = twb.Worksheets(ShName).ListObjects(UtNAme)

With lo
If IsArray(vf) Then
lo.Range.AutoFilter
lo.Range.AutoFilter
For I = 0 To UBound(vf)
.Range.AutoFilter Field:=Val(vf(I)), Criteria1:=vc(I)
Next I
Else
lo.Range.AutoFilter
lo.Range.AutoFilter
End If

End With


End Function
