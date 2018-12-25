Attribute VB_Name = "mod_Form_and_ListObject"
 Sub ufNastroykaShow()
 ufNastroyka.Show 0
 End Sub

Sub uflistObjectManagerShow()
uflistObjectManager.Show 0
End Sub
Sub ufListObjectShow()
ufListObject.Show 0
'Call UFShow("ufListObject", 0)
End Sub

Sub ufSapeTranspositionShow()
ufSapeTransposition.Show 0
End Sub
Sub UFShow(UFName, modal)
Dim uf
Set uf = New ufListObject
'Set uf = GET_UFName(UFName)
uf.Show modal
End Sub

Function GET_UFName(UFName)
Attribute GET_UFName.VB_ProcData.VB_Invoke_Func = " \n14"
Dim frm
For Each frm In UserForms
If UCase(frm.name) = UCase(UFName) Then
Set GET_UFName = frm
Exit For
End If
Set GET_UFName = Nothing
Next

End Function


Function GET_UF_Control_Name(UFName, UF_ControlName)
Dim frm
Dim cntrl
For Each frm In UserForms
If UCase(frm.name) = UCase(UFName) Then
For Each cntrl In frm.Controls
If UCase(cntrl.name) = UCase(UF_ControlName) Then
Set GET_UF_Control_Name = cntrl

Exit For
End If
Set GET_UF_Control_Name = Nothing
Next
End If

Next

End Function
Function UF_Loaded(UFName)
Dim bLoaded As Boolean
Dim frm
'Load UserForm1
For Each frm In UserForms
If UCase(frm.name) = UCase(UFName) Then
'MsgBox "Userform1 is loaded"
bLoaded = True
Exit For
End If
Next
If Not bLoaded Then
UF_Loaded = False
Else
UF_Loaded = True
End If

End Function
Sub HeaderRowRangeColumnName_Listbox_or_combobox(wb, lo, UFName, UF_ControlName)
Dim r
'Заголовки столбцов Умной Таблицы в список ListBox or Combobox
Set UFContrlol = GET_UF_Control_Name(UFName, UF_ControlName)

Set r = lo.HeaderRowRange 'ActiveListObjectHeaderRowRangeRange(WB, lo.Parent.name, lo.Range.Cells(1))
v = Application.WorksheetFunction.Transpose(r)
If r.Cells.Count = 1 Then

UFContrlol.Clear
UFContrlol.AddItem v
Else
UFContrlol.list = v
End If

End Sub
Sub ListObject_to_Listbox_or_combobox(wb, lo, UFName, UF_ControlName, Optional ColumnName)

'Заголовки столбцов Умной Таблицы в список ListBox or Combobox
Set UFContrlol = GET_UF_Control_Name(UFName, UF_ControlName)
If IsMissing(ColumnName) Then
Set r = lo.DataBodyRange
Else
Set r = lo.ListColumns(ColumnName).DataBodyRange
End If
v = r
If r.Cells.Count = 1 Then

UFContrlol.Clear
UFContrlol.AddItem v
Else
UFContrlol.list = v
End If

End Sub


Sub Value_Control(UFName, UF_ControlName, value)

'Заголовки столбцов Умной Таблицы в список ListBox or Combobox
Set UFContrlol = GET_UF_Control_Name(UFName, UF_ControlName)


UFContrlol.value = value


End Sub
