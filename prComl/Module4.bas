Attribute VB_Name = "Module4"
Sub ������1()
Attribute ������1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������1 ������
'

'
    Columns("G:G").ColumnWidth = 16.22
End Sub
Sub Example()
    Dim ADO As New ADO
    
    ADO.Query ("SELECT F1 FROM [����1$];")
    Range("E1").CopyFromRecordset ADO.Recordset
    
    ADO.Query ("SELECT F2 FROM [����1$];")
    Range("F1").CopyFromRecordset ADO.Recordset
    
    ' ��������� ����������, ����� �� ������ : )
    ADO.Disconnect
    
    ADO.Query ("SELECT F1 FROM [����1$] UNION SELECT F2 FROM [����1$];")
    Range("G1").CopyFromRecordset ADO.Recordset
    
    ' ��� ������������� ��������� ����������
    ' � ������������ ������� Recordset � Connection
End Sub
