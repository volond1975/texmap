Attribute VB_Name = "����"
Public SelectedDate As String, DefaultDate As String
Public ssst_1 As Date, ssst_2 As Date
'Private Sub ����������������_Click()
   ' Me.txt_���� = Get_Date(Me.txt_����, Now)    ' ����� ���� �� ���������
'End Sub
Function Get_Date1(Optional ByVal StartDate As String, Optional ByVal Default_date As String = "") As String
    SelectedDate = StartDate
    DefaultDate = Default_date
    Form_SelectDate.Show
    Get_Date1 = CStr(SelectedDate)
End Function



Sub ������1()
Attribute ������1.VB_Description = "������ ������� 26.09.2012 (volond)"
Attribute ������1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������1 ������
' ������ ������� 26.09.2012 (volond)
'

'
    Selection.NumberFormat = "[$-FC22]d mmmm yyyy"" �."";@"
End Sub
Sub ������3()
Attribute ������3.VB_Description = "������ ������� 26.09.2012 (volond)"
Attribute ������3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������3 ������
' ������ ������� 26.09.2012 (volond)
'

'
    Selection.NumberFormat = "m/d/yyyy"
End Sub
