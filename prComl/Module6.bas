Attribute VB_Name = "Module6"
Sub ������11()
Attribute ������11.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������11 ������
'

'
    Range("AK4").Select
    ActiveWorkbook.RefreshAll
End Sub
Sub ������12()
Attribute ������12.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������12 ������
'

'
    Cells.Select
    Range("P1").Activate
    Selection.Delete Shift:=xlUp
    Range("T11").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("S2").Select
End Sub
Sub ������_��������()
Attribute ������_��������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������13 ������
'
Call ��������_��������
mDBQ = fAkt.ComboBox_����_��������
mDefaultDir = ThisWorkbook.Path
ThisWorkbook.Activate
msourse = "ODBC;DSN=����� Excel;DBQ=" & mDBQ & ";DefaultDir=" & mDefaultDir & ";DriverId=790;MaxBufferSize=2048;PageTimeout=5;"
'    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
'        "ODBC;DSN=����� Excel;DBQ=G:\Dropbox\���\�������� 2015.xls;DefaultDir=G:\Dropbox\���;DriverId=790;MaxBufferSize=2048;PageTimeout=5;" _
'        , Destination:=Range("$A$1")).QueryTable
        Worksheets("��������").Activate
         With Worksheets("��������").ListObjects.Add(SourceType:=0, Source:= _
        msourse _
        , destination:=Range("$A$1")).QueryTable
        .CommandText = Array( _
        "SELECT `��������$`.������, `��������$`.˳�������, `��������$`.`��� �����`, `��������$`.`������������� ����`, `��������$`.`���#���#`, `��������$`.`������ �����`, `��������$`.`����� ������`, `��������$`" _
        , _
        ".`��-�`, `��������$`.����, `��������$`.`�����,��`, `��������$`.`ʳ���-��� �����`, `��������$`.����1, `��������$`.����2, `��������$`.����3, `��������$`.����4, `��������$`.����5, `��������$`.����6, `��" _
        , _
        "������$`.����7, `��������$`.����8, `��������$`.����9, `��������$`.����10, `��������$`.����11, `��������$`.����12, `��������$`.`���� ������`, `��������$`.�����, `��������$`.`����� ��1 ��`, `��������$`." _
        , _
        "`������� ���`, `��������$`.`������� ������`, `��������$`.��������, `��������$`.`��������� ��������`, `��������$`.`����� �����`, `��������$`.`�#�����`, `��������$`.`���� �������� ������`, `��������$`.`" _
        , _
        "ID ����� ����������˳��\³� �����\��\���`, `��������$`.`���� ��������`, `��������$`.`���� ������ ���� ��������`, `��������$`.`���� ��������� ���� ��������`, `��������$`.³�����������, `��������$`." _
        , _
        "`���� ��������`, `��������$`.`���� �������� � ���`, `��������$`.`����� �������� � ���`, `��������$`.��, `��������$`.`���� ��`, `��������$`.`���������� ���������`, `��������$`.`������ ����� ���������`" _
        , _
        ", `��������$`.`���� � ��������`, `��������$`.`���� � ��������`, `��������$`.`���� � ���������`, `��������$`.`������� �� ��`, `��������$`.�������1, `��������$`.����������, `��������$`.�����������, `���" _
        , "�����$`.`ϳ�������� `, `��������$`.��� FROM `��������$`")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "�������_������_��_�����_Excel"
        .Refresh 'BackgroundQuery:=False
    End With
    ActiveSheet.ListObjects(1).name = "��������"
End Sub
Sub ��������_��������()
Attribute ��������_��������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������14 ������
'

On Error Resume Next
ThisWorkbook.Activate
 Worksheets("��������").Activate
    ActiveSheet.ListObjects("��������").Unlist
    Cells.Select
    Range("P1").Activate
    Selection.ClearContents
End Sub

Sub ������15()
'
' ������14 ������
'

'
    Range("Q4").Select
   Worksheets("��������").ListObjects("��������").Unlist
    Cells.Select
    Range("P1").Activate
    Selection.ClearContents
End Sub
