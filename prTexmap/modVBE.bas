Attribute VB_Name = "modVBE"
'---------------------------------------------------------------------------------------
' Module    : modVBE
' Author    : ВАЛЕРА
' Date      : 22.04.2016
' Purpose   :
'---------------------------------------------------------------------------------------
Public Enum ProcScope1
        ScopePrivate = 1
        ScopePublic = 2
        ScopeFriend = 3
        ScopeDefault = 4
    End Enum
    
    Public Enum LineSplits1
        LineSplitRemove = 0
        LineSplitKeep = 1
        LineSplitConvert = 2
    End Enum
    
    Public Type ProcInfo
        ProcName As String
        ProcKind As VBIDE.vbext_ProcKind
        ProcStartLine As Long
        ProcBodyLine As Long
        ProcCountLines As Long
        ProcScope1 As ProcScope1
        ProcDeclaration As String
    End Type








'http://www.cpearson.com/excel/vbe.aspx
'http://www.rondebruin.nl/win/s9/win002.htm
'https://isc.sans.edu/diary/OfficeMalScanner+helps+identify+the+source+of+a+compromise/18291
'Добавление пунктов меню в редактор VBA (VBE)
'http://www.cpearson.com/excel/vbemenus.aspx
'http://stackoverflow.com/questions/15457262/vba-msforms-vs-controls-whats-the-difference
'http://www.mrexcel.com/forum/excel-questions/658136-each-frame-userform.html
Option Explicit
Dim wb As Workbook, VBCompName As String, tCodeModText As String
Sub ааа()
Call ImporttCodeModule
End Sub

'---------------------------------------------------------------------------------
Sub ImporttCodeModule(Optional fileName$)
    Dim Filt$, Title$, Message As VbMsgBoxResult
    Do Until Message = vbNo
         'type of file to browse for
        Filt = "VB Files (*.bas; *.frm; *.cls)(*.bas; *.frm; *.cls)," & _
        "*.bas;*.frm;*.cls"
         'caption for browser
        Title = "SELECT A FOLDER - CLICK OPEN TO IMPORT - " & _
        "CANCEL TO QUIT"
         'browser
       If IsMissing(fileName$) Then fileName = Application.GetOpenFilename(FileFilter:=Filt, _
        FilterIndex:=5, Title:=Title)
        On Error GoTo Finish '< cancelled
        Application.VBE.ActiveVBProject.VBComponents.Import _
        (fileName)
         'finished?
        Message = MsgBox(fileName & vbCrLf & " has been imported " & _
        "- more imports?", vbYesNo, "More Imports?")
    Loop
Finish:
End Sub
 
Sub ImportActiveCeltCodeModule()
If Not IsVBcomponent(FSO_GetBaseName(ActiveCell.value)) Then ImporttCodeModule (ActiveCell.value)
End Sub

Function IsVBcomponent(VbcName) As Boolean
Dim VBc As Object
On Error Resume Next
Set VBc = Application.VBE.ActiveVBProject.VBComponents(VbcName)
If Err = 0 Then IsVBcomponent = True Else IsVBcomponent = False
End Function
'------------------------------------------------------------------------------------------------------

Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    
'ThisWorkbook.VBProject.References.AddFromGuid _
'        GUID:="{0002E157-0000-0000-C000-000000000046}", _
'        Major:=5, Minor:=3
    ''' The tCode modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The tCode below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
'    References dialog, scroll down to Microsoft Visual Basic for Applications Extensibility 5.3
''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
  
    If FolderWithVBAProjectFiles(wkbSource) & "\*.*" = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles(wkbSource) & "\*.*"
    On Error GoTo 0

    
    
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the tCode"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles(wkbSource) & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub
Sub ExportComponents()
On Error Resume Next
Dim lr
Dim B
Dim vbp
Dim i
Dim sh As Worksheet
Dim NumComponents
Dim v, z, k, s, p, VBComponentsFileName
'    Dim VBP As VBProject
    Set vbp = ActiveWorkbook.VBProject
    NumComponents = vbp.VBComponents.Count
   Set sh = SheetExistBookCreate(ActiveWorkbook, "VBComponents", True)
   v = Array("????????????", "????", "???.????? ????", "??????", "?????? ???")
   sh.Range(sh.Cells(1, 1), sh.Cells(1, 5)) = v
   
    z = 2
    For i = 1 To NumComponents
'
        
'
        Select Case vbp.VBComponents(i).Type
            Case 1
                k = "bas"
            Case 2
                k = "cls"
            Case 3
                k = "frm"
            Case 100
                k = "cls"
        End Select
        s = Split(ActiveWorkbook.name, ".")
     p = ActiveWorkbook.Path & "\" & s(0) & "\"
     
  VBComponentsFileName = vbp.VBComponents(i).name & "." & k
  If vbp.VBComponents(i).tCodeModule.CountOfLines <> 0 Then
  sh.Cells(z, 1) = vbp.VBComponents(i).name
  
  ChDir (p)
  If Err <> 0 Then
  MkDir (p)
  ChDir (p)
  End If
vbp.VBComponents(i).Export p & VBComponentsFileName
sh.Cells(z, 2) = "\" & s(0) & "\" & VBComponentsFileName
        sh.Cells(z, 3) = _
         vbp.VBComponents(i).tCodeModule.CountOfLines
    z = z + 1
End If

    Next i
End Sub

Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Object
    Dim objFile As Object
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents
Set objFSO = CreateObject("Scripting.FileSystemObject")
    If ActiveWorkbook.name = ThisWorkbook.name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the tCode"
    Exit Sub
    End If

    ''' NOTE: Path where the tCode modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the tCode modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Sub ImportComponents()
Dim Path As String
Dim sh As Worksheet
Dim fd As FileDialog
Dim lr
Dim B
Dim vbp
Dim i
Dim shc As Worksheet
'Требуется
 'LastRow
' SheetExistBookCreate
On Error Resume Next
Set sh = SheetExistBookCreate(ThisWorkbook, "VBComponents", False)

lr = LastRow(sh.name)
Set fd = Application.FileDialog(msoFileDialogOpen)
With fd
.AllowMultiSelect = False
.Filters.Clear
.Filters.Add "xls", "*.xls;,*.xlsm;*.xlsb;*.xlsx", 1
fd.InitialView = msoFileDialogViewList
If fd.Show = -1 Then fd.Execute
End With

'    Dim VBP As VBProject
Set B = ActiveWorkbook
    Set vbp = B.VBProject
  
   
    
    For i = 2 To lr
With vbp.VBComponents
If sh.Cells(i, 4) = 1 Then
.Remove vbp.VBComponents(sh.Cells(i, 1))
.Import ThisWorkbook.Path & sh.Cells(i, 2)
sh.Cells(i, 5) = Date
End If
End With
       
    Next i
    
  Set sh = SheetExistBookCreate(ThisWorkbook, "Оглавление", False)
  lr = LastRow(sh.name)
  For i = 2 To lr
  If sh.Cells(i, 5) = 1 Then
  Set shc = Worksheets(sh.Cells(i, 1))
  ThisWorkbook.Worksheets(sh.Cells(i, 1)).copy After:=B.Worksheets(B.Worksheets.Count)
  End If
  Next i
End Sub
Function FolderWithVBAProjectFiles(wkbSource) As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String
Dim myNameProject As String
    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = wkbSource.Path

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    myNameProject = FSO_GetBaseName(wkbSource.FullName)
    If fso.FolderExists(SpecialPath & myNameProject) = False Then
        On Error Resume Next
        MkDir SpecialPath & myNameProject
        On Error GoTo 0
    End If
    
    If fso.FolderExists(SpecialPath & myNameProject) = True Then
        FolderWithVBAProjectFiles = SpecialPath & myNameProject
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function

Function DeleteVBAModulesAndUserFormsName(name As String)


        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
            If VBComp.name = name Then
                VBProj.VBComponents.Remove VBComp
                End If
            End If
        Next VBComp
End Function

Sub AddModuleToProject(wb As Workbook, bext_ct_StdModule As String, myVBCompName As String)
'Добавление модуль в проект

        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent

        Set VBProj = wb.VBProject
        Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
        VBComp.name = myVBCompName
    End Sub

Sub AddModuleToProjectActiveWorkbook(bext_ct_StdModule As String, VBCompName As String)
        Dim wb As Workbook
    Set wb = ActiveWorkbook
    Call AddModuleToProject(wb, bext_ct_StdModule, VBCompName)
    End Sub

Sub tCodeModToModule()

        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim tCodeMod As VBIDE.tCodeModule
        Dim LineNum As Long
        Dim VBCompName, sttCodeModToModule
        Const DQUOTE = """" ' one " character
VBCompName = "UDFs"
        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(VBCompName)
        Set tCodeMod = VBComp.tCodeModule

        With tCodeMod
  
 
'      .InsertLines 1
     .InsertLines 1, "'dateshtamp=" & VBA.Now
     sttCodeModToModule = .Lines(2, 1)
     If sttCodeModToModule Like "*dateshtamp*" Then .DeleteLines 2
'            LineNum = .CountOfLines + 1
'            .InsertLines LineNum, "Public Sub SayHello()"
'            LineNum = LineNum + 1
'            .InsertLines LineNum, "    MsgBox " & DQUOTE & "Hello World" & DQUOTE
'            LineNum = LineNum + 1
'            .InsertLines LineNum, "End Sub"


        End With
    End Sub




Sub AddProcedureToModule(wb As Workbook, VBCompName As String, tCodeModText As String)
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim tCodeMod As VBIDE.tCodeModule
        Dim LineNum As Long
        Const DQUOTE = """" ' one " character

        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(VBCompName)
        Set tCodeMod = VBComp.tCodeModule

        With tCodeMod
'            LineNum = .CountOfLines + 1
'            .InsertLines LineNum, "Public Sub SayHello()"
'            LineNum = LineNum + 1
'            .InsertLines LineNum, "    MsgBox " & DQUOTE & "Hello World" & DQUOTE
'            LineNum = LineNum + 1
'            .InsertLines LineNum, "End Sub"


        End With
    End Sub
Sub AddProcedureToModuleActiveWorkbook(bext_ct_StdModule As String, VBCompName As String)
        Dim wb As Workbook
     Set wb = ActiveWorkbook
    Call AddProcedureToModule(wb, bext_ct_StdModule, VBCompName)
    End Sub
Sub AddReferensVBE()
ThisWorkbook.VBProject.references.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", _
        major:=5, minor:=3
End Sub

Sub ExportNameControlsForm()
On Error Resume Next
Dim lr
Dim B
Dim vbp
Dim i, j, ContrCount
Dim sh As Worksheet
Dim NumComponents
Dim contrl As Object
Dim muf As Object
Dim v, z, k, s, p, VBComponentsFileName
'    Dim VBP As VBProject
    Set vbp = ActiveWorkbook.VBProject
    NumComponents = vbp.VBComponents.Count
   Set sh = SheetExistBookCreate(ActiveWorkbook, "NameControlsForm", True)
   v = Array("FormName", "ContrName", "ContrType", "Comment", "?????? ???")
   sh.Range(sh.Cells(1, 1), sh.Cells(1, 5)) = v
   sh.Activate
    z = 2
    For j = 1 To NumComponents
'
        
'
        
  If vbp.VBComponents(j).Type = 3 Then
Set muf = OpenUserForms(vbp.VBComponents(j).name)
' ContrCount = vbp.VBComponents(i).Controls.Count
For Each contrl In muf.Controls
Select Case TypeName(contrl)
            Case 1
                k = "bas"
            Case 2
                k = "cls"
            Case 3
                k = "frm"
            Case 100
                k = "cls"
        End Select

  sh.Cells(z, 1) = vbp.VBComponents(j).name
  sh.Cells(z, 2) = contrl.name
  sh.Cells(z, 3) = TypeName(contrl)
  sh.Cells(z, 4) = contrl.ControlTipText

    z = z + 1

Next
muf.Hide
End If

    Next j
End Sub
Sub SyncVBAEditor()
'=======================================================================
' SyncVBAEditor
' This syncs the editor with respect to the ActiveVBProject and the
' VBProject containing the ActiveCodePane. This makes the project
' that conrains the ActiveCodePane the ActiveVBProject.
'=======================================================================
With Application.VBE
If Not .ActiveCodePane Is Nothing Then
    Set .ActiveVBProject = .ActiveCodePane.CodeModule.parent.Collection.parent
End If
End With
End Sub
Sub OpenVBE()
' Open the Visual Basic Editor Programmatically

Application.VBE.MainWindow.Visible = True
'The next line of code goes to a specified module
ThisWorkbook.VBProject.VBComponents(ActiveCell.value).Activate

End Sub
 
Sub OpenVBEProcStart()
Dim lStartLine As Long
Application.VBE.MainWindow.Visible = True
ThisWorkbook.VBProject.VBComponents(ActiveCell.value).Activate

With Application.VBE.ActiveCodePane.CodeModule
lStartLine = .ProcStartLine(ActiveCell.Offset(columnoffset:=1).value, 0)
.CodePane.SetSelection lStartLine, 1, lStartLine, 1
End With
End Sub


Sub OpenVBEProcSelect()
        Dim VBProj ' В VBIDE.VBProject
       Dim VBComp 'As VBIDE.VBComponent
     Dim CodeMod ' As VBIDE.CodeModule
        Dim startLine As Long
        Dim NumLines As Long
        Dim ProcName As String
      Application.VBE.MainWindow.Visible = True
ThisWorkbook.VBProject.VBComponents(ActiveSheet.name).Activate
        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(ActiveSheet.name)
        Set CodeMod = VBComp.CodeModule
    
        ProcName = ActiveCell.value
        With CodeMod
            startLine = .ProcStartLine(ProcName, vbext_pk_Proc)
            NumLines = .ProcCountLines(ProcName, vbext_pk_Proc)
.CodePane.SetSelection startLine, 1, startLine + NumLines, 1

'  .DeleteLines StartLine:=StartLine, Count:=NumLines
        End With
    End Sub

Sub ListProcedures()

        Dim VBProj ' As VBIDE.VBProject
        Dim VBComp 'As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim LineNum As Long
        Dim NumLines As Long
        Dim ws As Worksheet
        Dim rng As Range
        Dim ProcName As String
        Dim ProcKind As VBIDE.vbext_ProcKind
        
        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(ActiveCell.value)
        Set CodeMod = VBComp.CodeModule
        Set ws = SheetExistBookCreate(ActiveWorkbook, ActiveCell.value, True)
'   v = Array("FormName", "ContrName", "ContrType", "Comment", "?????? ???")
'   Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, 5)) = v
'        Set WS = ActiveWorkbook.Worksheets("Лист7")
        Set rng = ws.Range("A1")
        With CodeMod
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                ProcName = .ProcOfLine(LineNum, ProcKind)
                rng.value = ProcName
                rng(1, 2).value = ProcKindString(ProcKind)
                LineNum = .ProcStartLine(ProcName, ProcKind) + _
                        .ProcCountLines(ProcName, ProcKind) + 1
                Set rng = rng(2, 1)
            Loop
        End With

    End Sub
    
    
    Function ProcKindString(ProcKind) As String
        Select Case ProcKind
            Case vbext_pk_Get
                ProcKindString = "Property Get"
            Case vbext_pk_Let
                ProcKindString = "Property Let"
            Case vbext_pk_Set
                ProcKindString = "Property Set"
            Case vbext_pk_Proc
                ProcKindString = "Sub Or Function"
            Case Else
                ProcKindString = "Unknown Type: " & CStr(ProcKind)
        End Select
    End Function


 

    Function ProcedureInfo(ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
        CodeMod As VBIDE.CodeModule) As ProcInfo
    
        Dim PInfo As ProcInfo
        Dim BodyLine As Long
        Dim Declaration As String
        Dim FirstLine As String
        
        
        BodyLine = CodeMod.ProcStartLine(ProcName, ProcKind)
        If BodyLine > 0 Then
            With CodeMod
                PInfo.ProcName = ProcName
                PInfo.ProcKind = ProcKind
                PInfo.ProcBodyLine = .ProcBodyLine(ProcName, ProcKind)
                PInfo.ProcCountLines = .ProcCountLines(ProcName, ProcKind)
                PInfo.ProcStartLine = .ProcStartLine(ProcName, ProcKind)
                
                FirstLine = .Lines(PInfo.ProcBodyLine, 1)
                If StrComp(Left(FirstLine, Len("Public")), "Public", vbBinaryCompare) = 0 Then
                    PInfo.ProcScope1 = ScopePublic
                ElseIf StrComp(Left(FirstLine, Len("Private")), "Private", vbBinaryCompare) = 0 Then
                    PInfo.ProcScope1 = ScopePrivate
                ElseIf StrComp(Left(FirstLine, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
                    PInfo.ProcScope1 = ScopeFriend
                Else
                    PInfo.ProcScope1 = ScopeDefault
                End If
                PInfo.ProcDeclaration = GetProcedureDeclaration(CodeMod, ProcName, ProcKind, LineSplitKeep)
            End With
        End If
        
        ProcedureInfo = PInfo
    
    End Function
    
    
    Public Function GetProcedureDeclaration(CodeMod As VBIDE.CodeModule, _
        ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
        Optional LineSplitBehavior As LineSplits1 = LineSplitRemove)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetProcedureDeclaration
    ' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
    ' determines what to do with procedure declaration that span more than one line using
    ' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
    ' entire procedure declaration is converted to a single line of text. If
    ' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
    ' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
    ' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
    ' The function returns vbNullString if the procedure could not be found.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim LineNum As Long
        Dim s As String
        Dim Declaration As String
        
        On Error Resume Next
        LineNum = CodeMod.ProcBodyLine(ProcName, ProcKind)
        If Err.Number <> 0 Then
            Exit Function
        End If
        s = CodeMod.Lines(LineNum, 1)
        Do While Right(s, 1) = "_"
            Select Case True
                Case LineSplitBehavior = LineSplitConvert
                    s = Left(s, Len(s) - 1) & vbNewLine
                Case LineSplitBehavior = LineSplitKeep
                    s = s & vbNewLine
                Case LineSplitBehavior = LineSplitRemove
                    s = Left(s, Len(s) - 1) & " "
            End Select
            Declaration = Declaration & s
            LineNum = LineNum + 1
            s = CodeMod.Lines(LineNum, 1)
        Loop
        Declaration = SingleSpace(Declaration & s)
        GetProcedureDeclaration = Declaration
        
    
    End Function
    
    Private Function SingleSpace(ByVal Text As String) As String
        Dim pos As String
        pos = InStr(1, Text, Space(2), vbBinaryCompare)
        Do Until pos = 0
            Text = Replace(Text, Space(2), Space(1))
            pos = InStr(1, Text, Space(2), vbBinaryCompare)
        Loop
        SingleSpace = Text
    End Function

'You can call the ProcedureInfo function using code like the following:
    Sub ShowProcedureInfo()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim CompName As String
        Dim ProcName As String
        Dim ProcKind As VBIDE.vbext_ProcKind
        Dim PInfo As ProcInfo
        
        CompName = "modVBE"
        ProcName = "ProcedureInfo"
        ProcKind = vbext_pk_Proc
        
        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(CompName)
        Set CodeMod = VBComp.CodeModule
        
        PInfo = ProcedureInfo(ProcName, ProcKind, CodeMod)
        
        Debug.Print "ProcName: " & PInfo.ProcName
        Debug.Print "ProcKind: " & CStr(PInfo.ProcKind)
        Debug.Print "ProcStartLine: " & CStr(PInfo.ProcStartLine)
        Debug.Print "ProcBodyLine: " & CStr(PInfo.ProcBodyLine)
        Debug.Print "ProcCountLines: " & CStr(PInfo.ProcCountLines)
        Debug.Print "ProcScope1: " & CStr(PInfo.ProcScope1)
        Debug.Print "ProcDeclaration: " & PInfo.ProcDeclaration
    End Sub
Sub SearchCodeModule()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim FindWhat As String
        Dim sl As Long ' start line
        Dim EL As Long ' end line
        Dim SC As Long ' start column
        Dim EC As Long ' end column
        Dim Found As Boolean
        
        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents("modVBE")
        Set CodeMod = VBComp.CodeModule
        
        FindWhat = "ProcedureInfo"
        
        With CodeMod
            sl = 1
            EL = .CountOfLines
            SC = 1
            EC = 255
            Found = .Find(Target:=FindWhat, startLine:=sl, startColumn:=SC, _
                endLine:=EL, endColumn:=EC, _
                wholeword:=True, MatchCase:=False, patternsearch:=False)
            Do Until Found = False
                Debug.Print "Found at: Line: " & CStr(sl) & " Column: " & CStr(SC)
                EL = .CountOfLines
                SC = EC + 1
                EC = 255
                Found = .Find(Target:=FindWhat, startLine:=sl, startColumn:=SC, _
                    endLine:=EL, endColumn:=EC, _
                    wholeword:=True, MatchCase:=False, patternsearch:=False)
            Loop
        End With
    End Sub

Sub Remove_MISSING_VBProject()
'http://www.excel-vba.ru/chto-umeet-excel/oshibka-cant-find-project-or-library/
'для автоматического поиска и отключения ошибочных ссылок на  библиотеки можно делать и макросом
    Dim oReferences As Object, oRef As Object
    Set oReferences = ThisWorkbook.VBProject.references
    For Each oRef In oReferences
        If (oRef.IsBroken) Then oReferences.Remove Reference:=oRef
    Next
End Sub

Sub Check_VBOM()
'http://www.excel-vba.ru/chto-umeet-excel/chto-neobxodimo-dlya-vneseniya-izmenenij-v-proekt-vbamakrosy-programmno/
    Dim oVBProj As Object
    On Error Resume Next
    Set oVBProj = ActiveWorkbook.VBProject
    If Not oVBProj Is Nothing Then
        MsgBox "Доступ к проектной модели VBA разрешен", vbInformation
    Else
        MsgBox "Доступ к проектной модели VBA запрещен", vbInformation
    End If
End Sub
Sub Change_VBOM()
'http://www.excel-vba.ru/chto-umeet-excel/chto-neobxodimo-dlya-vneseniya-izmenenij-v-proekt-vbamakrosy-programmno/
    Dim objExcelApp As Object, objShell As Object, sExVersion As String, lLevel As Long
 
    'Определяем версию Excel и в зависимости от этого определяем ветку реестра
    Set objExcelApp = CreateObject("Excel.Application")
    sExVersion = objExcelApp.version: objExcelApp.Quit
 
    Set objShell = CreateObject("WScript.Shell")
    lLevel = objShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & sExVersion & "\Excel\Security\AccessVBOM")
    'Разрешаем доступ к объектной модели VBA
    'AccessVBOM - 0 - запрещен доступ; 1 - разрешен
    If lLevel = 0 Then
        objShell.RegWrite _
                "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
                sExVersion & "\Excel\Security\AccessVBOM", 1, "REG_DWORD"
    End If
    Set objExcelApp = Nothing: Set objShell = Nothing
End Sub
Sub REG_VBOM()
'http://www.excel-vba.ru/chto-umeet-excel/chto-neobxodimo-dlya-vneseniya-izmenenij-v-proekt-vbamakrosy-programmno/
    Dim objShell As Object, sExVersion As String, lLevel As Long
 
    'Определяем версию Excel и в зависимости от этого определяем ветку реестра
    sExVersion = Application.version
 
    Set objShell = CreateObject("WScript.Shell")
    lLevel = objShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & sExVersion & "\Excel\Security\AccessVBOM")
    'Проверяем доступ к объектной модели VBA
    If lLevel = 0 Then
        MsgBox "Доступ к проектной модели VBA запрещен", vbInformation
    Else
        MsgBox "Доступ к проектной модели VBA разрешен", vbInformation
    End If
    Set objShell = Nothing
End Sub
Sub Unprotect_VBA()
    Dim objVBProject As Object, objVBComponent As Object, objWindow As Object
 
    Workbooks.Open "C:\1.xls"
    Set objVBProject = ActiveWorkbook.VBProject
    'просматриваем все окна проекта в поисках окна снятия защиты
    For Each objWindow In objVBProject.VBE.Windows
        ' Type = 6 - это нужное нам окно
        If objWindow.Type = 6 Then
            objWindow.Visible = True
            objWindow.SetFocus: Exit For
        End If
    Next
    'вводим пароль и подтверждаем ввод
'    Тильды нужны, но они не являются частью кода. Т.е. сам код это - 1234
    SendKeys "~1234~", True: SendKeys "{ENTER}", True
    'здесь Ваш код по внесению изменений в проект
    Set objVBProject = Nothing: Set objVBComponent = Nothing: Set objWindow = Nothing
    ActiveWorkbook.Close True
End Sub
Sub Copy_Module()
    Dim objVBProjFrom As Object, objVBProjTo As Object, objVBComp As Object
    Dim sModuleName As String, sFullName As String
    'расширение стандартного модуля
    Const sExt As String = ".bas"
 
    'имя модуля для копирования
    sModuleName = "Module1"
    On Error Resume Next
    'проект книги, из которой копируем модуль
    Set objVBProjFrom = ThisWorkbook.VBProject
    'необходимый компонент
    Set objVBComp = objVBProjFrom.VBComponents(sModuleName)
    'если указанного модуля не существует
    If objVBComp Is Nothing Then
        MsgBox "Модуль с именем '" & sModuleName & "' отсутствует в книге.", vbCritical, "Error"
        Exit Sub
    End If
    'проект книги для добавления модуля
    Set objVBProjTo = ActiveWorkbook.VBProject
    'полный путь для экспорта/импорта модуля. К папке должен быть доступ на запись/чтение
    sFullName = "C:\" & sModuleName & sExt
    objVBComp.Export fileName:=sFullName
    objVBProjTo.VBComponents.Import fileName:=sFullName
    'удаляем временный файл для импорта
    Kill sFullName
End Sub



'---------------------------------------------------------------------------------------
' Procedure : CopyVBComponent
' DateTime  : 02.08.2013 23:10
' Author    : The_Prist(Щербаков Дмитрий)
'             http://www.excel-vba.ru
' Purpose   : Функция копирует компонент из одной книги в другую.
'             Возвращает True, если копирование прошло удачно
'             False - если компонент не удалось скопировать
'
' wbFromFrom             Книга, компонент из VBA-проекта которой необходимо копировать
'
' wbFromTo               Книга, в VBA-проект которой необходимо копировать компонент
'
' sModuleName            Имя модуля, который необходимо копировать.
'
' bOverwriteExistModule  Если True или 1, то при наличии в конечной книге
'                        компонента с именем sModuleName - он будет удален,
'                        а вместо него импортирован копируемый.
'                        Если False, то при наличии в конечной книге
'                        компонента с именем sModuleName функция вернет False,
'                        а сам компонент не будет скопирован.
'---------------------------------------------------------------------------------------
'
Function CopyVBComponent(sModuleName As String, _
    wbFromFrom As Workbook, wbFromTo As Workbook, _
    bOverwriteExistModule As Boolean) As Boolean
    
    Dim objVBProjFrom As Object, objVBProjTo As Object
    Dim objVBComp As Object, objTmpVBComp As Object
    Dim sTmpFolderPath As String, sVBCompName As String, sModuleCode As String
    Dim lSlashPos As Long, lExtPos As Long
    
    'Проверяем корректность указанных параметров
    On Error Resume Next
    Set objVBProjFrom = wbFromFrom.VBProject
    Set objVBProjTo = wbFromTo.VBProject
    
    If objVBProjFrom Is Nothing Then
        CopyVBComponent = False: Exit Function
    End If
    If objVBProjTo Is Nothing Then
        CopyVBComponent = False: Exit Function
    End If
    
    If Trim(sModuleName) = "" Then
        CopyVBComponent = False: Exit Function
    End If
    
    If objVBProjFrom.Protection = 1 Then
        CopyVBComponent = False: Exit Function
    End If
    
    If objVBProjTo.Protection = 1 Then
        CopyVBComponent = False: Exit Function
    End If
    
    Set objVBComp = objVBProjFrom.VBComponents(sModuleName)
    If objVBComp Is Nothing Then
        CopyVBComponent = False: Exit Function
    End If
    
    '====================================================
    'полный путь для экспорта/импорта модуля. К папке должен быть доступ на запись/чтение
    sTmpFolderPath = Environ("Temp") & "\" & sModuleName & ".bas" '"
    If bOverwriteExistModule = True Then
        ' Если bOverwriteExistModule = True
        ' удаляем из временной папки и из конечного проекта
        ' модуль с указанным именем
        If Dir(sTmpFolderPath, 6) <> "" Then
            Err.Clear
            Kill sTmpFolderPath
            If Err.Number <> 0 Then
                CopyVBComponent = False: Exit Function
            End If
        End If
        With objVBProjTo.VBComponents
            .Remove .Item(sModuleName)
        End With
    Else
        Err.Clear
        Set objVBComp = objVBProjTo.VBComponents(sModuleName)
        If Err.Number <> 0 Then
            'Err.Number 9 - отсутствие указанного компонента, что нам не мешает.
            'Если ошибка другая - выход из функции
            If Err.Number <> 9 Then
                CopyVBComponent = False: Exit Function
            End If
        End If
    End If
    
    '====================================================
    'Экспорт/Импорт компонента во временную директорию
    objVBProjFrom.VBComponents(sModuleName).Export sTmpFolderPath
    'Получаем имя компонента из экспортированного файла
    lSlashPos = InStrRev(sTmpFolderPath, "\")
    lExtPos = InStrRev(sTmpFolderPath, ".")
    sVBCompName = Mid(sTmpFolderPath, lSlashPos + 1, lExtPos - lSlashPos - 1)
    
    '====================================================
    'копируем
    Set objVBComp = Nothing
    Set objVBComp = objVBProjTo.VBComponents(sVBCompName)
    If objVBComp Is Nothing Then
        objVBProjTo.VBComponents.Import sTmpFolderPath
    Else
        'Если компонент - модуль листа или книги -
        'его нельзя удалить. Поэтому удаляем из него весь код
        'и добавляем код из копируемого компонента
        If objVBComp.Type = 100 Then
            'создаем временный компонент
            Set objTmpVBComp = objVBProjTo.VBComponents.Import(sTmpFolderPath)
            'копируем из него код
            With objVBComp.CodeModule
                .DeleteLines 1, .CountOfLines
                sModuleCode = objTmpVBComp.CodeModule.Lines(1, objTmpVBComp.CodeModule.CountOfLines)
                .InsertLines 1, sModuleCode
            End With
            On Error GoTo 0
            'удаляем временный компонент
            objVBProjTo.VBComponents.Remove objTmpVBComp
        End If
    End If
    'удаляем временный файл компонента
    Kill sTmpFolderPath
    CopyVBComponent = True
End Function
Sub CopyComponent()
'Пример вызова функции CopyVBComponent:
    Workbooks.Add
    If CopyVBComponent("ЭтаКнига", ThisWorkbook, ActiveWorkbook, True) Then
        MsgBox "Указанный компонент успешно скопирован в новую книгу", vbInformation
    Else
        MsgBox "Компонент не был скопирован", vbInformation
    End If
End Sub
Sub CreateEventProcedure()
'CОЗДАНИЕ СОБЫТИЙНОЙ ПРОЦЕДУРЫ Workbook_Open
'Важно: для русской версии используется ссылка на ЭтаКнига. Для английской ThisWorkbook
    Dim objVBProj As Object, objVBComp As Object, objCodeMod As Object
    Dim lLineNum As Long
    'добавляем новую книгу
    Workbooks.Add
    'получаем ссылку на проект и модуль книги
    Set objVBProj = ActiveWorkbook.VBProject
    Set objVBComp = objVBProj.VBComponents("ЭтаКнига")
    Set objCodeMod = objVBComp.CodeModule
    'вставляем код
    With objCodeMod
        lLineNum = .CreateEventProc("Open", "Workbook")
        lLineNum = lLineNum + 1
        .InsertLines lLineNum, "    MsgBox ""Hello World"""
    End With
End Sub
Sub CreateEventProcedure_WorkSheetChange()
'CОЗДАНИЕ СОБЫТИЙНОЙ ПРОЦЕДУРЫ Worksheet_Change в Лист1
'Важно: для русской версии используется ссылка на Лист1. Для английской как правило Sheet1

    Dim objVBProj As Object, objVBComp As Object, objCodeMod As Object
    Dim lLineNum As Long
    'добавляем новую книгу
    Workbooks.Add
    'получаем ссылку на проект и модуль листа
    Set objVBProj = ActiveWorkbook.VBProject
    Set objVBComp = objVBProj.VBComponents("Лист1")
    Set objCodeMod = objVBComp.CodeModule
    'вставляем код
    With objCodeMod
        lLineNum = .CreateEventProc("Change", "Worksheet")
        lLineNum = lLineNum + 1
        .InsertLines lLineNum, "    MsgBox ""Hello World"""
    End With
End Sub
Sub Create_NewModule()
    Dim objVBProj As Object, objVBComp As Object, objCodeMod As Object
    Dim sModuleName As String, sFullName As String
    Dim sProcLines As String
    Dim lLineNum As Long
    
    'добавляем новый стандартный модуль в активную книгу
    Set objVBComp = ActiveWorkbook.VBProject.VBComponents.Add(1)
    'получаем ссылку на коды модуля
    Set objCodeMod = objVBComp.CodeModule
    'узнаем количество строк в модуле
    '(т.к. VBA в зависимости от настроек может добавлять строки деклараций)
    lLineNum = objCodeMod.CountOfLines + 1
    'текст всставляемой процедуры
    sProcLines = "Sub Test()" & vbCrLf & _
        "    MsgBox ""Hello, World""" & vbCrLf & _
        "End Sub"
    'вставляем текст процедуры в тело нового модуля
    objCodeMod.InsertLines lLineNum, sProcLines
End Sub


 
Sub AddControlsDesigner()
     
     
    Dim Frm             As Object
    Dim Btn             As MSForms.CommandButton
    Dim cntr            As Object
    Dim X               As Long
    Dim n               As Long
    Dim BtnName         As String
     Dim cntrtype As String
     Dim ws As Worksheet
Dim rf As Range
Dim crf As Range
Dim lis As Range
Dim loj As clsmListObjs
Dim wb As Workbook
Dim v As Range
Dim lop As ListObject
Set wb = ActiveWorkbook
Dim mufcntr As Object
Set loj = New clsmListObjs
'Set muf = OpenUserForms(UFName)
Set mufcntr = UserFormControl(muf, contrName)
mufcntr.Clear
With loj
.Initialize wb
If ListObjName Like "*[!_0-9]*" Then
Set lop = .items(ListObjName)
     
    For X = 1 To ThisWorkbook.VBProject.VBComponents.Count
        If ThisWorkbook.VBProject.VBComponents(X).Type = 3 Then
            Set Frm = ThisWorkbook.VBProject.VBComponents(X)
            If Frm.name = "UserForm1" Then
            cntrtype = "forms.CommandButton.1"
'            Set Btn = Frm.Designer.Controls.Add()
Set cntr = Frm.Designer.Controls.Add(cntrtype)
            With cntr
'                .Caption = "Caption"
                .Height = 25
                .Width = 60
                .Left = 12
                .Top = 6
            End With
'            With Btn
'                .Caption = "Caption"
'                .Height = 25
'                .Width = 60
'                .Left = 12
'                .Top = 6
'            End With
'            With ThisWorkbook.VBProject.VBComponents(x).CodeModule
'                n = .CountOfLines
'                .InsertLines n + 1, "Sub CommandButton1_Click()"
'                .InsertLines n + 2, vbNewLine
'                .InsertLines n + 3, vbTab & "MsgBox " & """" & "Hi" & """"
'                .InsertLines n + 4, vbNewLine
'                .InsertLines n + 5, "End Sub"
'            End With
            End If
        End If
    Next X
   End With
End Sub
 
Sub AddUserFormName()
'UserFormControl(muf, contrName)
End Sub
