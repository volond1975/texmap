Attribute VB_Name = "РГК_ТЕХЗАВДАННЯ"
Sub ВзятьСТехзавдання()
Attribute ВзятьСТехзавдання.VB_Description = "Макрос записан 27.10.2012 (volond)"
Attribute ВзятьСТехзавдання.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос5 Макрос
' Макрос записан 27.10.2012 (volond)
Dim r1 As Range
Dim r2 As Range
Set r1 = Workbooks("Техзавдання.xls").Worksheets(1).Range("B20:C34")
Set r2 = Worksheets("Расчет").Range("F8")
Call Перенести(r1, r2)


Set r1 = Workbooks("Техзавдання.xls").Worksheets(1).Range("D20:F34")
Set r2 = Worksheets("Расчет").Range("I8")
Call Перенести(r1, r2)
    
Set r1 = Workbooks("Техзавдання.xls").Worksheets(1).Range("G20:H34")
Set r2 = Worksheets("Расчет").Range("L8")
Call Перенести(r1, r2)

    Set r1 = Workbooks("Техзавдання.xls").Worksheets(1).Range("I20:J34")
Set r2 = Worksheets("Расчет").Range("O8")
Call Перенести(r1, r2)
'
End Sub
Sub Перенести(r1 As Range, r2 As Range)
Attribute Перенести.VB_Description = "Макрос записан 27.10.2012 (volond)"
Attribute Перенести.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос6 Макрос
' Макрос записан 27.10.2012 (volond)
'

'
    Windows("Техзавдання.xls").Activate
    r1.Copy
'    Selection.Copy
    Windows("Техкарта Автоматизована уточнена 1.1.xls").Activate
    r2.Select
    Windows("Техзавдання.xls").Activate
    r1.Copy
    Windows("Техкарта Автоматизована уточнена 1.1.xls").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub ВзятьСРасчета()
'
' Макрос5 Макрос
' Макрос записан 27.10.2012 (volond)
Dim r1 As Range
Dim r2 As Range

'("Техкарта " & Me.TextBox_Год_Техкарта & "." & Me.ComboBox_Расширение_Техкарта, ThisWorkbook.Path & "\")

Set r1 = Workbooks("Техкарта " & Me.TextBox_Год_Техкарта & "." & Me.ComboBox_Расширение_Техкарта).Worksheets("Расчет").Range("F8:G22")
Set r2 = Worksheets("План заготівлі РГК").Range("B3")
Call Перенести(r1, r2)


Set r1 = Workbooks("Техзавдання.xls").Worksheets(1).Range("D20:F34")
Set r2 = Worksheets("Расчет").Range("I8")
Call Перенести(r1, r2)
    
Set r1 = Workbooks("Техзавдання.xls").Worksheets(1).Range("G20:H34")
Set r2 = Worksheets("Расчет").Range("L8")
Call Перенести(r1, r2)

    Set r1 = Workbooks("Техзавдання.xls").Worksheets(1).Range("I20:J34")
Set r2 = Worksheets("Расчет").Range("O8")
Call Перенести(r1, r2)
'
End Sub
Sub ПеренестиСРасчета(r1 As Range, r2 As Range)
'
' Макрос6 Макрос
' Макрос записан 27.10.2012 (volond)
'

'
    Windows("Техкарта " & Me.TextBox_Год_Техкарта & "." & Me.ComboBox_Расширение_Техкарта).Activate
    r1.Copy
'    Selection.Copy
    Windows("Техкарта Автоматизована уточнена 1.1.xls").Activate
    r2.Select
    Windows("Техзавдання.xls").Activate
    r1.Copy
    Windows("Техкарта Автоматизована уточнена 1.1.xls").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub Макрос13()
'
' Макрос13 Макрос
'

'Workbooks( _
'        "Техкарта " & fAkt.TextBox_Год_Техкарта & "." & fAkt.ComboBox_Расширение_Техкарта).Activate
'
'    Workbooks( _
'        "Техкарта " & fAkt.TextBox_Год_Техкарта & "." & fAkt.ComboBox_Расширение_Техкарта).Sheets("Расчет").Copy Before:=Workbooks( _
'        "Программа Акти 2 (Автосохраненный).xlsm").Sheets(1)
    Sheets("Расчет").Select
    Range("F8:G22").Select
    Selection.Copy
    Sheets("План заготівлі РГК").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Расчет").Select
    Range("I8:K22").Select
    Selection.Copy
    Sheets("План заготівлі РГК").Select
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Расчет").Select
    Range("L8:M22").Select
    Selection.Copy
    Sheets("План заготівлі РГК").Select
    Range("G3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Расчет").Select
    Range("O8:P22").Select
    Selection.Copy
    Sheets("План заготівлі РГК").Select
    Range("I3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Расчет").Select
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.Delete
End Sub
