Attribute VB_Name = "Module11"
Sub ����������������(r As Range)
Attribute ����������������.VB_Description = "������ ������� 09.11.2012 (volond)"
Attribute ����������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������17 ������
' ������ ������� 09.11.2012 (volond)
'

'
 
    With r
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
    End With
   ' r.Merge
    r.Font.Bold = True
'    With r
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlCenter
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = True
'    End With
End Sub
Sub ddd()
Dim r As Range
Set r = Selection
Call ����������������(r)
Call ������������(r, xlMedium)
End Sub

Sub ������������()
Dim r As Range
Set r = Selection
Call ������������(r, xlThin)
Call ������������(r, xlMedium)
End Sub
Sub ������������(r As Range, �������)
Attribute ������������.VB_Description = "������ ������� 09.11.2012 (volond)"
Attribute ������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������18 ������
' ������ ������� 09.11.2012 (volond)
'xlThin-������

'xlMedium-������
 r.Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = �������
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = �������
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = �������
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = �������
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    

End Sub
