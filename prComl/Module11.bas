Attribute VB_Name = "Module11"
Sub ФорматЗаголовков(r As Range)
Attribute ФорматЗаголовков.VB_Description = "Макрос записан 09.11.2012 (volond)"
Attribute ФорматЗаголовков.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос17 Макрос
' Макрос записан 09.11.2012 (volond)
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
Call ФорматЗаголовков(r)
Call СеткаТолщина(r, xlMedium)
End Sub

Sub СеткаСРамкой()
Dim r As Range
Set r = Selection
Call СеткаТолщина(r, xlThin)
Call СеткаТолщина(r, xlMedium)
End Sub
Sub СеткаТолщина(r As Range, Толщина)
Attribute СеткаТолщина.VB_Description = "Макрос записан 09.11.2012 (volond)"
Attribute СеткаТолщина.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос18 Макрос
' Макрос записан 09.11.2012 (volond)
'xlThin-тонкая

'xlMedium-жирная
 r.Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = Толщина
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = Толщина
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = Толщина
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = Толщина
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
