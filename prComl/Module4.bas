Attribute VB_Name = "Module4"
Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос1 Макрос
'

'
    Columns("G:G").ColumnWidth = 16.22
End Sub
Sub Example()
    Dim ADO As New ADO
    
    ADO.Query ("SELECT F1 FROM [Лист1$];")
    Range("E1").CopyFromRecordset ADO.Recordset
    
    ADO.Query ("SELECT F2 FROM [Лист1$];")
    Range("F1").CopyFromRecordset ADO.Recordset
    
    ' Закрываем соединение, чтобы не висело : )
    ADO.Disconnect
    
    ADO.Query ("SELECT F1 FROM [Лист1$] UNION SELECT F2 FROM [Лист1$];")
    Range("G1").CopyFromRecordset ADO.Recordset
    
    ' Тут автоматически закроется соединение
    ' и уничтожиться объекты Recordset и Connection
End Sub
