Attribute VB_Name = "������"
Sub ������������()
ActiveWindow.RangeSelection.SpecialCells(xlCellTypeFormulas).Select


With Selection.Validation
.Delete
.Add Type:=xlValidateCustom, Formula1:="="""""
            .ErrorTitle = "������!"
            .ErrorMessage = "� ������ �������!" & vbCrLf & "���� ������ ��������!"
            .ShowError = True



End With
End Sub
