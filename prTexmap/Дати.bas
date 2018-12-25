Attribute VB_Name = "Дати"
Public SelectedDate As String, DefaultDate As String
Public ssst_1 As Date, ssst_2 As Date
'Private Sub КнопкаВыбораДаты_Click()
   ' Me.txt_Дата = Get_Date(Me.txt_Дата, Now)    ' выбор даты из календаря
'End Sub
Function Get_Date1(Optional ByVal StartDate As String, Optional ByVal Default_date As String = "") As String
    SelectedDate = StartDate
    DefaultDate = Default_date
    Form_SelectDate.Show
    Get_Date1 = CStr(SelectedDate)
End Function



Sub Макрос1()
Attribute Макрос1.VB_Description = "Макрос записан 26.09.2012 (volond)"
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос1 Макрос
' Макрос записан 26.09.2012 (volond)
'

'
    Selection.NumberFormat = "[$-FC22]d mmmm yyyy"" р."";@"
End Sub
Sub Макрос3()
Attribute Макрос3.VB_Description = "Макрос записан 26.09.2012 (volond)"
Attribute Макрос3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос3 Макрос
' Макрос записан 26.09.2012 (volond)
'

'
    Selection.NumberFormat = "m/d/yyyy"
End Sub
