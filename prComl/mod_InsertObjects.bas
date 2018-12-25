Attribute VB_Name = "mod_InsertObjects"


'---------------------------------------------------------------------------------------
' Module        : mod_InsertObjects
' Author        : �����                     Date: 07.07.2013
' Professional application development for Microsoft Excel
' http://ExcelVBA.ru/          ICQ: 5836318           Skype: ExcelVBA.ru
'---------------------------------------------------------------------------------------

Option Compare Text
Option Private Module

Public Const LINK_HEADER_TABLE$ = "<ExcelTable>", LINK_HEADER_IMAGE$ = "<Image>"
Public CellWithLink As Range, ExcelTablesToBeClosed As New Collection

Sub CtrlShiftT(): InsertOrEditTableLink: End Sub
Sub CtrlShiftI(): MsgBox "���� �� �����������, - ����� � ��������� ������� ���������", vbInformation, "������� ������ �� �����������": End Sub


Sub InsertOrEditTableLink()
    On Error Resume Next
    Set CellWithLink = Nothing
    Set CellWithLink = ActiveCell
    If CellWithLink Is Nothing Then Exit Sub
    F_SelectTable.Show 0
End Sub

Function HasLinkToObject(ByVal txt$, Optional ByVal Key$) As Boolean
' ���������� TRUE, ���� � ������ ������ txt$ ���������� ������ �� ������
    HasLinkToObject = (txt$ Like LINK_HEADER_TABLE$ & "*/*/*") Or (txt$ Like LINK_HEADER_IMAGE$ & "*")
End Function

Sub InsertObjectIntoDOC(ByRef doc As Object, ByVal txt$, ByVal Key$)
' ��������� ������, ��������� � ������ txt$, � Word-�������� doc �� ����� ���� key$

    Dim Msg$, InsertMode$
    If Not CopyExcelTable(txt$, Msg$, InsertMode$) Then
        Msg$ = "�� ������� ����������� �������" & vbNewLine & txt & vbNewLine & vbNewLine & Msg$
        MsgBox Msg$, vbExclamation, "������ ��� ����������� ������� � ������"
        Exit Sub
    End If

    doc.Range.Select
    doc.ActiveWindow.View.ShowRevisionsAndComments = False

    With doc.Parent.Selection.Find
        .Text = Key$
        While .Execute
            Select Case InsertMode$
                Case "Excel", ""
                    doc.Parent.Selection.PasteExcelTable False, False, False    ' �������� ��������������
                Case "Word"
                    doc.Parent.Selection.PasteExcelTable False, True, False    ' ����� Word
                Case "PlainText"
                    doc.Parent.Selection.PasteAndFormat 22    ' (wdFormatPlainText) ' ������� �����
                Case "Picture"
                    doc.Parent.Selection.PasteAndFormat 13    '(wdChartPicture) ' ��������
            End Select
        Wend
    End With
    
    Application.CutCopyMode = False
End Sub

Function InsertTableStylesArray() As Variant
    ReDim Arr(0 To 3, 0 To 1)
    Arr(0, 0) = "Excel": Arr(0, 1) = "�������� �������������� Excel"
    Arr(1, 0) = "Word": Arr(1, 1) = "������������ ����� Word"
    Arr(2, 0) = "PlainText": Arr(2, 1) = "������ �������� �����"
    Arr(3, 0) = "Picture": Arr(3, 1) = "��������� ��� �����������"
    InsertTableStylesArray = Arr
End Function

Function CopyExcelTable(ByVal link$, Optional ByRef Msg$, Optional ByRef InsertMode$) As Boolean
    On Error Resume Next
    Dim ra As Range, sh As Worksheet, wb As Workbook
    If link$ Like LINK_HEADER_TABLE$ & "*/*/*" Then
        link$ = Split(link$, LINK_HEADER_TABLE$)(1)
        Filename$ = Split(link$, "/")(0)
        If Filename$ = "" Then
            Filename$ = ActiveWorkbook.FullName
        Else
            If (Not Filename$ Like "[A-Z]:\*") And (Not Filename$ Like "\\*") Then
                Filename$ = TABLES_FOLDER$ & Filename$
            End If
        End If
        shortFilename$ = Dir(Filename$, vbNormal)
        If Len(shortFilename$) Then
            Err.Clear: Set wb = Workbooks(CStr(shortFilename$))
            If wb Is Nothing Then
                Application.DisplayAlerts = False
                ExcelTablesToBeClosed.Add shortFilename$, shortFilename$
                Set wb = Workbooks.Open(Filename$, , True)
                Application.DisplayAlerts = True
            End If
            'If Err Then Debug.Print Err.Number, Err.Description, filename$
            If wb Is Nothing Then Msg$ = "�� ������� ������� ���� �" & shortFilename$ & "�": Exit Function

            SheetName$ = Split(link$, "/")(1)
            Set sh = wb.Worksheets(CStr(SheetName$))
            If sh Is Nothing Then Msg$ = "� ����� �" & shortFilename$ & "� �� ������ ���� �" & SheetName$ & "�": Exit Function

            RangeAddress$ = Split(link$, "/")(2)
            If RangeAddress$ = "UsedRange" Then
                Set ra = sh.UsedRange
            Else
                Set ra = sh.Range(RangeAddress$)
            End If
            If sh Is Nothing Then Msg$ = "������ � ������ ��������� �����: �" & RangeAddress$ & "�": Exit Function

            InsertMode$ = Split(link$, "/")(3)

            ra.Copy
            CopyExcelTable = True
        Else
            Msg$ = "���� �" & Filename$ & "� �� ������": Exit Function
        End If
    End If
End Function
