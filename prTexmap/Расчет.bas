Attribute VB_Name = "Расчет"

Sub ОчиститьФормуРГК()
'
' ОчиститьФормуРГК Макрос
' Макрос записан 25.06.2012 (Пользователь)
'

'
With Worksheets("Расчет")
.Range("F8:G22").ClearContents
.Range("I8:P22").ClearContents
.Range("S8:U22").ClearContents
End With

'Call CalculationA
'Call CalculationM
End Sub
Sub ПеренестиРасчетвРГК()
'
' ПеренестиРасчетвРГК Макрос
' Макрос записан 25.06.2012 (Пользователь)
'

'

    Range("G29:G40").Select
    Selection.copy
    Sheets("РГК").Select
    Range("F40:F50").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
  Sheets("РГК").Range("gNel").value = Sheets("Расчет").Range("G42")
        
End Sub
