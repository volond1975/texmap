Attribute VB_Name = "Защита"
Sub ЗащитаФормул()
ActiveWindow.RangeSelection.SpecialCells(xlCellTypeFormulas).Select


With Selection.Validation
.Delete
.Add Type:=xlValidateCustom, Formula1:="="""""
            .ErrorTitle = "ОШИБКА!"
            .ErrorMessage = "В ячейке формула!" & vbCrLf & "Ввод данных запрещён!"
            .ShowError = True



End With
End Sub
