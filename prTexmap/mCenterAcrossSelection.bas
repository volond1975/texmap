Attribute VB_Name = "mCenterAcrossSelection"
Sub ВыравниваниеПоВыделению(r As Range, Optional bFontBold = False)
Attribute ВыравниваниеПоВыделению.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ВыравниваниеПоВыделению Макрос
'

'
   
    With r
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = bFontBold
    End With
End Sub
Sub testGAS()
v = CAS1(Selection)
End Sub
'https://stackoverflow.com/questions/45113115/ms-excel-vba-to-find-range-formatted-with-center-across-selection-horizontal-a


Public Function CAS(r As Range) As String
    Dim i As Long, rng As Range
    CAS = ""
'For j = 1 To Columns.Count
    If r.HorizontalAlignment <> 7 Then Exit Function
    Set rng = r

    For i = 1 To Columns.Count
   ' If r.HorizontalAlignment <> 7 Or r.Offset(0, i + 1) <> "" Or r.Offset(0, i + 1) <> Empty Then
   ' Set rng = Union(rng, r.Offset(0, i))
    'Exit Function
    'End If
        If r.HorizontalAlignment <> 7 Or r.Offset(0, i) <> "" Or r.Offset(0, i) <> Empty Then
            CAS = rng.Address(0, 0)
            Exit Function
        Else
            Set rng = Union(rng, r.Offset(0, i))
        End If
    Next i
  '  Next i
End Function
Public Function CAS1(r As Range) As String
    Dim i As Long, rng As Range
    CAS1 = ""
For j = 1 To Columns.Count
    If r.HorizontalAlignment <> 7 Then Exit Function
    Set rng = r

    For i = 1 To Columns.Count
   ' If r.HorizontalAlignment <> 7 Or r.Offset(0, i + 1) <> "" Or r.Offset(0, i + 1) <> Empty Then
   ' Set rng = Union(rng, r.Offset(0, i))
    'Exit Function
    'End If
        If r.HorizontalAlignment <> 7 Or r.Offset(0, i) <> "" Or r.Offset(0, i) <> Empty Then
            CAS1 = rng.Address(0, 0)
            r = r.Offset(0, 0)
            Debug.Print CAS1
           ' Exit Function
        Else
            Set rng = Union(rng, r.Offset(0, i))
        End If
    Next i
   Next j
End Function
