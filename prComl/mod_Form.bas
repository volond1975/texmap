Attribute VB_Name = "mod_Form"
Sub ����������������������������(rend As Range, rhome As Range)
Attribute ����������������������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������1 ������
'

rk = rhome.Row - rend.Row
ck = rend.Column - rhome.Column
    rend.FormulaR1C1 = "=SUM(R[" & rk & "]C:R[-1]C)"
    Range("G30").Select
End Sub
Sub ���������������(r As Range)
Attribute ���������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������2 ������
'

'
    
    r.FormulaR1C1 = "=RC[-2]*RC[-1]"
    
End Sub
Sub ������7()
Attribute ������7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������7 ������
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
' ������8 ������
'

'
    fAkt.Show 0
End Sub
Sub fDataShow()
'
' ������8 ������
'

'
    fData.Show 0
End Sub
Sub ������9()
Attribute ������9.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������9 ������
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
' ���������� ����� ��������� CopyButton ��� CopyControl �� ����� ��� ���������� �����.
    End With
Next
End Sub
