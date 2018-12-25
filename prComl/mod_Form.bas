Attribute VB_Name = "mod_Form"
Sub ФормулаСуммыРегионНадЯчейкой(rend As Range, rhome As Range)
Attribute ФормулаСуммыРегионНадЯчейкой.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос1 Макрос
'

rk = rhome.Row - rend.Row
ck = rend.Column - rhome.Column
    rend.FormulaR1C1 = "=SUM(R[" & rk & "]C:R[-1]C)"
    Range("G30").Select
End Sub
Sub ОбьемНаВартисть(r As Range)
Attribute ОбьемНаВартисть.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос2 Макрос
'

'
    
    r.FormulaR1C1 = "=RC[-2]*RC[-1]"
    
End Sub
Sub Макрос7()
Attribute Макрос7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос7 Макрос
'

'
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub fAktShow()
Attribute fAktShow.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос8 Макрос
'

'
    fAkt.Show 0
End Sub
Sub fDataShow()
'
' Макрос8 Макрос
'

'
    fData.Show 0
End Sub
Sub Макрос9()
Attribute Макрос9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос9 Макрос
'

'
    With Selection.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub
Sub frmEDERPUorINNShow()
frmEDERPUorINN.Show 0
End Sub
Public Sub New_Commandbar()
Dim Cbr As CommandBar, Ctr As CommandBarControl
On Error Resume Next
Application.CommandBars("My_cell").Delete
Application.CommandBars.Add name:="My_cell", Position:=msoBarPopup, temporary:=True
For Each Ctr In Application.CommandBars("cell").Controls
    With Application.CommandBars("My_cell").Controls.Add(Ctr.Type, Ctr.ID, Ctr.Parameter, , 1)
        .Caption = Ctr.Caption
'       .OnAction = Ctr.OnAction
        .BeginGroup = Ctr.BeginGroup
' Копировать также используя CopyButton или CopyControl не помню как называется метод.
    End With
Next
End Sub
