Attribute VB_Name = "Module6"
Sub ������1()
'
' ������1 ������
'

'
    Range("E27").Select
    ActiveWindow.SmallScroll Down:=-27
    Range("E34").Select
    Sheets("��������").Select
    Range("C15").Select
    ActiveWorkbook.Save
    Columns("C:C").Select
    Selection.ColumnWidth = 20
End Sub
