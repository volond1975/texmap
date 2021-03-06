VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADO"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'--------------------------------
' ADO (�� ����. ActiveX Data Objects � �������� ������ ActiveX�) � ��������� ���������������� ���������� ��� ������� � ������,
' ������������� ��������� Microsoft (MS Access, MS SQL Server) � ���������� �� ���������� ����������� ActiveX.
' ADO ��������� ������������ ������ �� ������������� ���������� (����������� ��� ������, ��������� ������ � �. �.)
' � ��������-��������������� ����.
'
' ����� ��� �������� ������� ����������� (Connection) � ����������� SQL �������� � ������.
'
' reference - http://msdn.microsoft.com/ru-ru/library/windows/desktop/ms678086(v=vs.85).aspx
' wiki - http://ru.wikipedia.org/wiki/ADO
'
' http://support.microsoft.com/kb/257819/ru
' http://support.microsoft.com/kb/316934/ru
' http://support.microsoft.com/kb/246335/ru
'
' http://www.w3schools.com/ado/default.asp
'
' @author nerv
' @version 01/05/2013, 0.5
'--------------------------------

' ������:
'   Create     - ������� ����� �����������. ������������� ���������� ��� �������������.
'   Connect    - ��������� ����������. ������������� ���������� ��� �������.
'   Destroy    - ���������� ������ ����������� � ������ �������. ���������� ������������� ��� ������ �� ���������.
'   Disconnect - ��������� �������� ������ � �����������. ���������� ������������� ��� ������ �� ���������.
'   Query      - ��������� SQL ������. ��������� ������� �������� � ������ Recordset. ���������� �����, � ������� ��� �������� ������.
'   ToArray    - ���������� ��������� ������� � ���� �������

' ��������:
'   Connection - ������ ����������.
'   Recordset  - ��������� ��������� �������.
'   DataSoure  - �������� ������. ������ ��� ����� ������. �� ��������� ������� �����.
'   Header     - ��������� ��������� (��/���). �� ��������� ���. � ���� ������ ����� ����� ����������� ������������� F1 ... Fn.
'                ���� ��, ������ ������ ��������� ��������� ���������� ����.

' � ������ �������� ��������� (ConnectionString) � ����� Connect, �������� ������� DataSoure � Header �� �����������,
' � ������������ ������ ���������� �������� ��������� �� ������������.


Option Explicit


Public Connection As Object
Public Recordset As Object
Public DataSource As String
Public Header As Boolean


Public Sub Create()
    Set Me.Connection = CreateObject("ADODB.Connection")
End Sub


Public Sub Destroy()
    Call Me.Disconnect
    
    Set Me.Recordset = Nothing
    Set Me.Connection = Nothing
End Sub


Public Sub Connect(Optional ByVal ConnectionString As String)
    If Me.Connection Is Nothing Then
        Call Me.Create
    End If

    If ConnectionString = "" Then
        Me.Connection.Open GetExcelConnectionString()
    Else
        Me.Connection.Open ConnectionString
    End If
End Sub


Public Sub Disconnect()
    If Not Me.Recordset Is Nothing Then
        If Me.Recordset.State = 1 Then
            Me.Recordset.Close
            Me.Connection.Close
        End If
    End If
End Sub


Public Function Query(ParamArray QueryString() As Variant) As Date
    If Me.Connection Is Nothing Then
        Call Me.Create
    End If
    
    If Me.Connection.State = 0 Then
        Call Me.Connect
    End If
    
    Set Me.Recordset = Me.Connection.Execute(Join(QueryString, " "))
    Query = Now
End Function


Public Function ToArray() As Variant
    Dim Data() As Variant
    Dim Arr() As Variant
    Dim Column As Long
    Dim Row As Long
    
    If Me.Recordset Is Nothing Then
        Exit Function
    End If
    
    If Me.Recordset.BOF And Me.Recordset.EOF Then
        Exit Function
    End If
    
    Data = Me.Recordset.getRows()
    
    ReDim Arr(1 To UBound(Data, 2) + 1, 1 To UBound(Data, 1) + 1)
    
    For Column = 0 To UBound(Data, 2)
        For Row = 0 To UBound(Data, 1)
            Arr(Column + 1, Row + 1) = Data(Row, Column)
        Next
    Next
    
    ToArray = Arr
End Function


Private Sub Class_Initialize()
    Call Me.Create
End Sub


Private Sub Class_Terminate()
    Call Me.Disconnect
    Call Me.Destroy
End Sub


Private Function GetExcelConnectionString() As String
    Dim Params As String
    Params = IIf(Val(Application.Version) < 12, _
        "Provider='Microsoft.Jet.OLEDB.4.0';Data Source=':1';Extended Properties='Excel 4.0;HDR=:2;IMEX=1';", _
        "Provider='Microsoft.ACE.OLEDB.12.0';Data Source=':1';Extended Properties='Excel 12.0;HDR=:2;IMEX=1';")
    Params = Replace(Params, ":1", IIf(Me.DataSource = "", ThisWorkbook.FullName, Me.DataSource), , 1)
    Params = Replace(Params, ":2", IIf(Me.Header, "Yes", "No"), , 1)
    GetExcelConnectionString = Params
End Function
