VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_SelectTable 
   OleObjectBlob   =   "F_SelectTable.frx":0000
   Caption         =   "Добавление \ изменение ссылки на таблицу"
   ClientHeight    =   3720
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11175
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   25
End
Attribute VB_Name = "F_SelectTable"
Attribute VB_Base = "0{811A8216-6112-4E7D-AA51-B2AAB0B1099D}{16E79F10-C3F9-45A7-9318-BACF8871407D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

'Public Const LINK_HEADER_TABLE$ = "<ExcelTable>", LINK_HEADER_IMAGE$ = "<Image>"
'Public CellWithLink As Range, ExcelTablesToBeClosed As New Collection

 Private Sub CheckBox_UsedRange_Click()
     On Error Resume Next
     Me.TextBox_Range.Enabled = Not Me.CheckBox_UsedRange.value
     Me.CommandButton_SelectRange.Enabled = Not Me.CheckBox_UsedRange.value

     If Me.CheckBox_UsedRange.value Then
         Me.TextBox_Range.Tag = Me.TextBox_Range.Text
         Me.TextBox_Range = ActiveSheet.UsedRange.address(0, 0, xlA1)
         ActiveSheet.UsedRange.Select
     Else
         Me.TextBox_Range.Text = Me.TextBox_Range.Tag
         Application.ReferenceStyle = xlA1
         ActiveSheet.Range(Me.TextBox_Range).Select
     End If
 End Sub

 Private Sub ComboBox_Filename_Change()
     On Error Resume Next
     If Me.ComboBox_Filename = "" Then Exit Sub
     Workbooks(CStr(Me.ComboBox_Filename)).Activate
     Me.ComboBox_SheetName.Clear

     Dim sh As Worksheet
     For Each sh In ActiveWorkbook.Worksheets
         If sh.Visible = xlSheetVisible Then
             Me.ComboBox_SheetName.AddItem sh.name
         End If
     Next sh
     Me.ComboBox_SheetName = ActiveSheet.name
 End Sub

 Private Sub ComboBox_SheetName_Change()
     On Error Resume Next
     ActiveWorkbook.Worksheets(CStr(Me.ComboBox_SheetName)).Activate
     CheckFields
 End Sub

 Private Sub CommandButton_AddLink_Click()
     On Error Resume Next: Err.Clear
     Dim wb As Workbook, sh As Worksheet
     Set wb = Workbooks(CStr(Me.ComboBox_Filename))
     Set sh = wb.Worksheets(CStr(Me.ComboBox_SheetName))
     If wb Is Nothing Then Exit Sub
     If wb.Path = "" Then Exit Sub
     If sh Is Nothing Then Exit Sub

     RangeAddress$ = Me.TextBox_Range
     If Me.CheckBox_UsedRange Then RangeAddress$ = "UsedRange"

     Filename$ = wb.FullName
     If Filename$ Like TABLES_FOLDER$ & "*" Then Filename$ = Split(Filename$, TABLES_FOLDER$)(1)
     If Filename$ = CellWithLink.Worksheet.Parent.FullName Then Filename$ = ""
     link$ = "<ExcelTable>" & Filename$ & "/" & sh.name & "/" & RangeAddress$ & "/" & Me.ComboBox_InsertTableMode

     CellWithLink.value = link$
     CellWithLink.Worksheet.Parent.Activate
     CellWithLink.Worksheet.Activate
     'If Err = 0 Then
     Unload Me
 End Sub

 Private Sub CommandButton_OpenWokbook_Click()
     On Error Resume Next
     ChDrive Left(TABLES_FOLDER, 2)
     ChDir TABLES_FOLDER
     Application.Dialogs(xlDialogOpen).Show
     UpdateFilenamesList
 End Sub

 Private Sub CommandButton_Quit_Click()
     Unload Me
 End Sub

 Private Sub CommandButton_SelectRange_Click()
     On Error Resume Next
     Dim ra As Range, n As name
     Set ra = Application.InputBox("Выделите диапазон ячеек", , , , , , , 8)
     Me.TextBox_Range = ra.address(0, 0, xlA1)
     For Each n In ra.Worksheet.Parent.Names
         If n.RefersToRange.address = ra.address Then Me.TextBox_Range = n.name
     Next

     Me.ComboBox_Filename = ra.Worksheet.Parent.name
     Me.ComboBox_SheetName = ra.Worksheet.name
     ra.Select
 End Sub

 Private Sub TextBox_Range_Change()
     CheckFields
 End Sub



 Private Sub UserForm_Initialize()
     On Error Resume Next
     If CellWithLink Is Nothing Then Exit Sub
     UpdateFilenamesList
     Me.ComboBox_InsertTableMode.list = InsertTableStylesArray()

     link$ = CellWithLink.value
     If link$ Like "<ExcelTable>" & "*/*/*" Then
         link$ = Split(link$, "<ExcelTable>")(1)

         InsertTableMode$ = Trim(Split(link$, "/")(3))
         If InsertTableMode$ = "" Then InsertTableMode$ = "Excel"
         Me.ComboBox_InsertTableMode = InsertTableMode$


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
             If Not IsObject(Workbooks(CStr(shortFilename$))) Then
                 Application.DisplayAlerts = False
                 Workbooks.Open Filename$
                 Application.DisplayAlerts = True
             End If

             Err.Clear: Me.ComboBox_Filename = shortFilename$
             If Err Then Me.ComboBox_Filename.AddItem shortFilename$: Me.ComboBox_Filename = shortFilename$

             SheetName$ = Split(link$, "/")(1)
             Err.Clear: Me.ComboBox_SheetName = SheetName$
             If Err Then Exit Sub


             RangeAddress$ = Split(link$, "/")(2)
             If RangeAddress$ = "UsedRange" Then
                 Me.CheckBox_UsedRange = True
             Else
                 Me.TextBox_Range = RangeAddress$
             End If
             ActiveSheet.Range(Me.TextBox_Range).Select
         End If
Else
Me.ComboBox_SheetName = "Таблица Щоденник"
lr = LastRow("Таблица Щоденник")
 Me.TextBox_Range = Worksheets("Таблица Щоденник").Range(Cells(1, 1), Cells(lr, 9)).address
     End If
 End Sub

 Sub UpdateFilenamesList()
     On Error Resume Next: Me.ComboBox_Filename.Clear
     Dim wb As Workbook
     For Each wb In Application.Workbooks
         If wb.Windows(1).Visible = True And wb.Path <> "" Then
             Me.ComboBox_Filename.AddItem wb.name
         End If
     Next wb
     Me.ComboBox_Filename = ActiveWorkbook.name
 End Sub

 Sub CheckFields()
     On Error Resume Next
     Me.CommandButton_AddLink.Enabled = Me.ComboBox_SheetName.Text <> "" And Me.ComboBox_Filename.value <> "" And Me.TextBox_Range.Text <> ""
 End Sub

 Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     On Error Resume Next
     If CellWithLink Is Nothing Then Exit Sub
     CellWithLink.Worksheet.Parent.Activate
     CellWithLink.Worksheet.Activate
     CellWithLink.Select
     Set ThisWorkbook.app = Application
 End Sub
