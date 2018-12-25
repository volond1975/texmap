VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} nerv_FDropDownList 
   OleObjectBlob   =   "nerv_FDropDownList.frx":0000
   ClientHeight    =   645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4440
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   1
End
Attribute VB_Name = "nerv_FDropDownList"
Attribute VB_Base = "0{45275436-6689-40D3-B7F0-0FEA443D6928}{D0B16936-422D-4D8E-82DA-C46CDFFF7AB6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


'=========================================================
 ' Author: nerv            | E-mail: nerv-net@yandex.ru
 ' Last Update: 28/03/2012 | Yandex.Money: 41001156540584
 '=========================================================
 Option Explicit

 Private NotUse As Boolean   ' switch
 Private Arr()               ' array


 Private Sub UserForm_Initialize()
     Dim elem
     On Error Resume Next
     With CreateObject("Scripting.Dictionary")
         .compaode = vbTextCompare
         For Each elem In IIf(DDLRange.Count = 1, Array(DDLRange.value), DDLRange.value)
             If VarType(elem) = vbString Then elem = Trim(elem)
             If Not IsError(elem) Then
                 If Len(elem) > 0 Then .Add CStr(elem), elem
             End If
         Next
         Arr = .items
     End With
     With Me
         Call QuickSortNonRecursive(Arr)
         .Width = getFormWidth(DLLSheetSettings.Range("F59").value)
         .ComboBox1.Width = .Width - 10
         .Caption = getFormCaprion(DLLSheetSettings.Range("F41").value) & UBound(Arr) + 1
         If UBound(Arr) Then .ComboBox1.DropDown
         Call SetFormPosition(Me, DDLCell)
         .ComboBox1.List = Arr
     End With
 End Sub


 Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
     NotUse = False
     Select Case KeyCode
         Case vbKeyEscape: Unload Me: Exit Sub
         Case vbKeyReturn: Call ActionOfChoice: Exit Sub
         Case vbKeyPageDown: NotUse = True
         Case vbKeyPageUp: NotUse = True
         Case vbKeyDown: NotUse = True
         Case vbKeyUp: NotUse = True
     End Select
 End Sub


 Private Sub ComboBox1_Change()
     If NotUse Then Exit Sub
     If Me.ComboBox1.Text = "" Then
         Me.Caption = getFormCaprion(DLLSheetSettings.Range("F41").value) & UBound(Arr) + 1
         Me.ComboBox1.List = Arr
         Exit Sub
     End If
     Dim elem, pattern As String, register As Boolean
     pattern = getPattern(DLLSheetSettings.Range("F2").value)
     register = getRegister(DLLSheetSettings.Range("F29").value)
     pattern = getCase(pattern, register)
     With CreateObject("Scripting.Dictionary")
         For Each elem In Arr
             If getCase(elem, register) Like pattern Then .Add CStr(elem), elem
         Next
         Me.Caption = getSearchCaption(DLLSheetSettings.Range("F50").value) & .Count
         Me.ComboBox1.List = .items
     End With
 End Sub


 Private Sub UserForm_Activate() ' Search for the value entered in cell
     If searchEnteredValue(DLLSheetSettings.Range("F17").value) Then
         Dim EnteredValue
         EnteredValue = DDLCell.value
         If Not IsEmpty(EnteredValue) Then
             If Not IsError(EnteredValue) Then Me.ComboBox1.value = CStr(EnteredValue)
         End If
     End If
 End Sub


 Private Sub ActionOfChoice() ' Action when item is selected
     On Error Resume Next
     DDLCell.value = Me.ComboBox1.value
     If Err = 1004 Then
         MsgBox "Sorry, cell is protected from changes", vbExclamation
         Unload Me
         Exit Sub
     End If
     With Application
         If .MoveAfterReturn Then
             Select Case .MoveAfterReturnDirection
                 Case xlDown: DDLCell.Offset(1).Select
                 Case xlToLeft: DDLCell.Offset(, -1).Select
                 Case xlToRight: DDLCell.Offset(, 1).Select
                 Case xlUp: DDLCell.Offset(-1).Select
             End Select
         End If
     End With
     Unload Me
 End Sub


 Private Sub UserForm_Terminate()
     Set DLLSheetSettings = Nothing
     Set DDLRange = Nothing
     Set DDLCell = Nothing
     Erase Arr()
     Unload Me
 End Sub


 Private Sub ComboBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
     Call ActionOfChoice
 End Sub


 '-------------------------------------------
 Private Function getFormWidth(ByVal config As Variant) As Single
     If IsNumeric(config) Then
         If Not IsEmpty(config) Then
             If config >= 100 And config <= Application.Width / 2 Then
                 getFormWidth = CSng(config)
                 Exit Function
             End If
         End If
     End If
     getFormWidth = 210
 End Function


 Private Function getFormCaprion(ByVal config As Variant) As String
     If IsEmpty(config) Then
         getFormCaprion = "Unique records: "
         Exit Function
     End If
     getFormCaprion = config
 End Function


 Private Function getPattern(ByVal config As Variant) As String
     If VarType(config) = vbString Then
         If InStr(1, config, "request", vbTextCompare) Then
             getPattern = Replace(config, "request", Me.ComboBox1.Text)
             Exit Function
         End If
     End If
     getPattern = "*" & Me.ComboBox1.Text & "*"
 End Function


 Private Function getRegister(ByVal config As Variant) As Boolean
     If VarType(config) = vbBoolean Then
         getRegister = config
         Exit Function
     End If
     getRegister = False
 End Function


 Private Function getCase(ByVal Text As String, ByVal register As Boolean) As String
     If register Then getCase = Text Else getCase = LCase(Text)
 End Function


 Private Function getSearchCaption(ByVal config As Variant) As String
     If IsEmpty(config) Then
         getSearchCaption = "Search result: "
         Exit Function
     End If
     getSearchCaption = config
 End Function


 Private Function searchEnteredValue(ByVal config As Variant) As Boolean
     If VarType(config) = vbBoolean Then
         searchEnteredValue = config
         Exit Function
     End If
     searchEnteredValue = True
 End Function
