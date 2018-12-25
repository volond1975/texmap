Attribute VB_Name = "modHorizontalAlignmentMerge"
Sub Макрос3()
'
' Макрос3 Макрос
'

'
    Range("A2:B2").Select
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A3:B3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub
Sub ПоВыделению()
'
' ПоВыделению2 Макрос
'

'
   Dim r As Range
   v = Selection
   col = UBound(v)
 For Each r In Selection.Rows
    With r
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Next
End Sub
