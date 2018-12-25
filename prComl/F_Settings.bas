Attribute VB_Name = "F_Settings"
Public UseTempSettings As Boolean, TempSettingsCollection As New Collection
Sub F_Settings()
 Private Sub CheckBox_AddHyperlinks_Click()
     On Error Resume Next: Err.Clear
     Me.Label_HLink_Text.Enabled = Me.CheckBox_AddHyperlinks
     Me.TextBox_HyperlinkText.Enabled = Me.CheckBox_AddHyperlinks
 End Sub

 Private Sub CheckBox_Mail_AttachCreatedFiles_Click()
     On Error Resume Next: Err.Clear
     Me.CheckBox_Mail_AttachCreatedFiles.Font.Bold = Me.CheckBox_Mail_AttachCreatedFiles
     Me.Label_AttachCreatedFiles.Enabled = Me.CheckBox_Mail_AttachCreatedFiles
     Me.TextBox_AttachCreatedFilesMask.Enabled = Me.CheckBox_Mail_AttachCreatedFiles
 End Sub

 Private Sub CheckBox_Mail_AttachStaticFiles_Click()
     On Error Resume Next: Err.Clear
     Me.CheckBox_Mail_AttachStaticFiles.Font.Bold = Me.CheckBox_Mail_AttachStaticFiles
     Me.TextBox_Mail_AttachStaticFolder.Enabled = Me.CheckBox_Mail_AttachStaticFiles
     Me.CommandButton_Change_AttachStaticFolder.Enabled = Me.CheckBox_Mail_AttachStaticFiles

     Me.TextBox_Mail_AttachStaticFolder.BackColor = IIf(Me.CheckBox_Mail_AttachStaticFiles, vbWindowBackground, vbButtonFace)
 End Sub

 Private Sub CheckBox_MultiRow_Click()
     On Error Resume Next
     Me.CheckBox_Multirow_GroupRows.Enabled = Me.CheckBox_MultiRow
     Me.ComboBox_Multirow_GroupColumn.Visible = Me.CheckBox_MultiRow
     Me.Label_Multirow_GroupColumn.Visible = Me.CheckBox_MultiRow
 End Sub

 Private Sub CheckBox_Multirow_GroupRows_Click()
     On Error Resume Next
     Me.ComboBox_Multirow_GroupColumn.Enabled = Me.CheckBox_Multirow_GroupRows
     Me.Label_Multirow_GroupColumn.Enabled = Me.CheckBox_Multirow_GroupRows
 End Sub

 Private Sub CheckBox_SendEmail_Click()
     On Error Resume Next: Err.Clear
     Me.MultiPage_Options.Pages("Page_SendMail").Visible = Me.CheckBox_SendEmail
     Me.Label_SendEmail.Visible = Me.CheckBox_SendEmail
 End Sub


 Sub CommandButton_Change_AttachStaticFolder_Click()
     On Error Resume Next: Err.Clear
     AttachFolder$ = CreateObject("WScript.Shell").SpecialFolders("mydocuments") & "\"    'Вложения\"
     InitialPath$ = IIf(Me.TextBox_Mail_AttachStaticFolder <> "" And Not Me.TextBox_Mail_AttachStaticFolder Like "{*}", Me.TextBox_Mail_AttachStaticFolder, AttachFolder$)
     folder$ = GetFolderPath("Выберите папку, все файлы из которой будут прикреплены к письмам", InitialPath$)
     If folder$ = "" Then Exit Sub
     Me.TextBox_Mail_AttachStaticFolder = folder$
     Me.TextBox_Mail_AttachStaticFolder.ForeColor = vbBlack
 End Sub

 Private Sub CommandButton_Change_TheBAT_Path_Click()
     On Error Resume Next: Err.Clear
     New_TheBAT_Path$ = GetFilePath("Укажите путь к исполняемому файлу программы TheBAT!", TheBAT_PATH, "Приложение TheBAT!", "*.exe")
     If New_TheBAT_Path$ = "" Then Exit Sub
     Me.TextBox_TheBAT_Path = New_TheBAT_Path$
 End Sub

 Private Sub CommandButton_Quit_Click()
     Unload Me
 End Sub

 Private Sub CommandButton_ResetAllSettings_Click()
     On Error Resume Next
     Msg = "Вы уверены, что хотите сбросить все настройки программы к значениям по-умолчанию?" & vbNewLine & _
           "Отменить это действие невозможно." & vbNewLine & vbNewLine & _
           "Привести все настройки программы к исходным значениям?"
     If MsgBox(Msg, vbQuestion + vbOKCancel + vbDefaultButton2, "Сброс всех настроек программы") = vbCancel Then Exit Sub
     DeleteSetting PROJECT_NAME$, "Settings"

     ЗадержкаВЧасах$ = Replace(Format(CDbl(TimeSerial(0, 0, 1)) * 0.3, "0.000000000"), ",", ".")
     ExecuteExcel4Macro "ON.TIME(NOW()+" & ЗадержкаВЧасах$ & ", ""'" & ThisWorkbook.name & "'!ShowSettingsPage"")"
     Unload Me
 End Sub


 Private Sub Image_ExportSettings_Click()
     ExportSettings
 End Sub
 Private Sub Image_ImportSettings_Click()
     If Not ImportSettings Then Exit Sub

     ЗадержкаВЧасах$ = Replace(Format(CDbl(TimeSerial(0, 0, 1)) * 0.3, "0.000000000"), ",", ".")
     ExecuteExcel4Macro "ON.TIME(NOW()+" & ЗадержкаВЧасах$ & ", ""'" & ThisWorkbook.name & "'!ShowSettingsPage"")"
     Unload Me
 End Sub

 Private Sub Label_Help_FieldCodes_Click()
     On Error Resume Next
     URL$ = DEVELOPER_WEBSITE$ & "programmes/" & PROJECT_NAME$ & "/FieldCodes?ref=" & HID$
     CreateObject("wscript.Shell").Run URL$
 End Sub

 Private Sub Label_Help_FilenamesMask_Click()
     On Error Resume Next
     URL$ = DEVELOPER_WEBSITE$ & "programmes/" & PROJECT_NAME$ & "/FilenamesMask?ref=" & HID$
     CreateObject("wscript.Shell").Run URL$
 End Sub

 Private Sub Label_Help_InsertFormulasForSeparateLetters_Click()
     On Error Resume Next
     URL$ = DEVELOPER_WEBSITE$ & "programmes/" & PROJECT_NAME$ & "/SeparateLetters?ref=" & HID$
     CreateObject("wscript.Shell").Run URL$
 End Sub

 Private Sub Label_Help_InsertObjects_Click()
     On Error Resume Next
     URL$ = DEVELOPER_WEBSITE$ & "programmes/" & PROJECT_NAME$ & "/InsertObjects?ref=" & HID$
     CreateObject("wscript.Shell").Run URL$
 End Sub

 Private Sub Label_HelpMultiRow_Click()
     On Error Resume Next
     URL$ = DEVELOPER_WEBSITE$ & "programmes/" & PROJECT_NAME$ & "/MultiRow?ref=" & HID$
     CreateObject("wscript.Shell").Run URL$
 End Sub

 Private Sub Label_HelpMultiRowGroup_Click()
     On Error Resume Next
     URL$ = DEVELOPER_WEBSITE$ & "programmes/" & PROJECT_NAME$ & "/MultiRow/Group?ref=" & HID$
     CreateObject("wscript.Shell").Run URL$
 End Sub

 Private Sub Label_SendEmail_Click()
     On Error Resume Next: Err.Clear
     Me.MultiPage_Options.value = Me.MultiPage_Options.Pages("Page_SendMail").Index
 End Sub

 Private Sub Label_SendMail_Help_Click()
     On Error Resume Next
     URL$ = DEVELOPER_WEBSITE$ & "programmes/" & PROJECT_NAME$ & "/SendEmail?ref=" & HID$
     CreateObject("wscript.Shell").Run URL$
 End Sub


 Private Sub UserForm_Initialize()

     On Error Resume Next
     Set ThisWorkbook.app = Application
     For I = 1 To 50
         Me.ComboBox_BaseColumn.AddItem ColunmNameByColumnNumber(I)
         Me.ComboBox_Multirow_GroupColumn.AddItem ColunmNameByColumnNumber(I)
     Next

     Me.ComboBox_TheBAT_Account.Clear
     ' макрос выводит список всех потовых ящиков, настроенных в программе TheBAT!
     Err.Clear
     With CreateObject("WScript.Shell")
         For I = 1 To 100
             Key$ = "HKEY_CURRENT_USER\Software\RIT\The Bat!\Users depot\User #" & I
             Err.Clear: mailBox$ = .RegRead(Key$)
             If Err = 0 Then Me.ComboBox_TheBAT_Account.AddItem mailBox$
         Next
         Key$ = "HKEY_CURRENT_USER\Software\RIT\The Bat!\Users depot\Default"
         DefaultAccount$ = .RegRead(Key$)
         If Len(DefaultAccount$) Then Me.ComboBox_TheBAT_Account = DefaultAccount$
     End With

     ' код заполнения полей на вкладке "НАСТРОЙКИ ПРОГРАММЫ"
     Me.TextBox_TemplatesFolder = TEMPLATES_FOLDER$
     Me.TextBox_OutputFolder = OUTPUT_FOLDER$
     Me.CheckBox_PDF.value = PRINT_TO_PDF
     Me.CheckBox_ImmediatePrintOut = IMMEDIATE_PRINTOUT
     Me.TextBox_OutputMask = OUTPUT_MASK$
     Me.TextBox_TheBAT_Path = TheBAT_PATH


     Me.CheckBox_UseCurrentFolder = USE_CURRENT_FOLDER

     For I = 1 To 20: Me.ComboBox_FirstRow.AddItem I: Next I
     Me.ComboBox_FirstRow = HEADER_ROW


     Me.ComboBox_LineFeed.list = LineFeedOptions
     Me.ComboBox_LineFeed = LINEFEED_CHAR
     Me.CheckBox_USE_TEMPLATES_WITH_NAMES_LIKE_WORKSHEET_NAME = USE_TEMPLATES_WITH_NAMES_LIKE_WORKSHEET_NAME

     'Me.CheckBox_PDF.Enabled = Val(Application.Version) > 11
     LoadProgramSettings

     Me.MultiPage_Options.value = 0
     Me.MultiPage_Options.Pages("Page_AdditionalOptions").ScrollTop = 0
     Me.MultiPage_Options.Pages("Page_SendMail").ScrollTop = 0

     BaseCol& = Val(Settings("ComboBox_BaseColumn", 2))
     If BaseCol& = 0 Then BaseCol& = 2
     Me.ComboBox_BaseColumn = ColunmNameByColumnNumber(BaseCol&)

     If Settings("TextBox_AttachCreatedFilesMask", "") = "" Then
         TextBox_AttachCreatedFilesMask = "*"
         SaveSetting PROJECT_NAME$, "Settings", "TextBox_AttachCreatedFilesMask", "*"
     End If

     If Settings("TextBox_AttachStaticFilesMask", "") = "" Then
         TextBox_AttachStaticFilesMask = "*"
         SaveSetting PROJECT_NAME$, "Settings", "TextBox_AttachStaticFilesMask", "*"
     End If

     Me.MultiPage_Options.Pages("Page_SendMail").Visible = Me.CheckBox_SendEmail

     If Me.TextBox_HyperlinkText = "" Then Me.TextBox_HyperlinkText = "открыть файл"
 End Sub

 Sub LoadProgramSettings()
     On Error Resume Next: Dim ctrl As Control: NoSetting$ = "not found"
     For Each ctrl In Me.Controls
         sett = GetSetting(PROJECT_NAME$, "Settings", ctrl.name, NoSetting$)
         If ctrl.name Like "CheckBox*" Then
             If sett <> NoSetting$ Then sett = CBool(sett) Else sett = False
         End If
         If ctrl.Tag = "" And sett <> NoSetting$ Then ctrl.value = sett
     Next: Err.Clear
     UpdateFoldersFieldsAndButtons
 End Sub
 Sub SaveProgramSettings()
     On Error Resume Next: Dim ctrl As Control
     For Each ctrl In Me.Controls
         If ctrl.Enabled Then
             If Not ctrl.name Like "CommandButton_*" Then
                 SaveSetting PROJECT_NAME$, "Settings", ctrl.name, ctrl.value
             End If
         End If
     Next: Err.Clear
 End Sub

 Private Sub CommandButton_SaveSettings_Click()
     On Error Resume Next
     Dim RebuildMenu As Boolean
     If SettingsBoolean("CheckBox_ShowAdditionalMenu") <> Me.CheckBox_ShowAdditionalMenu Then RebuildMenu = True
     SaveProgramSettings
     If RebuildMenu Then CreateProgramCommandBar
     Enable_HotKeys
     Set ThisWorkbook.app = Application
     Unload Me
 End Sub

 Private Sub CheckBox_UseCurrentFolder_Click()
     On Error Resume Next
     SaveSetting PROJECT_NAME$, "Settings", "CheckBox_UseCurrentFolder", Me.CheckBox_UseCurrentFolder
     UpdateFoldersFieldsAndButtons
 End Sub

 Sub UpdateFoldersFieldsAndButtons()
     On Error Resume Next
     Me.TextBox_TemplatesFolder = TEMPLATES_FOLDER$(True)
     Me.TextBox_OutputFolder = OUTPUT_FOLDER$(True)
     Dim UseCurrentFolder As Boolean: UseCurrentFolder = USE_CURRENT_FOLDER

     Me.CommandButton_ChangeOutputFolder.Enabled = Not UseCurrentFolder
     Me.CommandButton_ChangeTemplatesFolder.Enabled = Not UseCurrentFolder

     Me.TextBox_OutputFolder.Enabled = Not UseCurrentFolder
     Me.TextBox_TemplatesFolder.Enabled = Not UseCurrentFolder
 End Sub

 Private Sub CommandButton_ChangeOutputFolder_Click()
     InitialPath$ = IIf(Dir(OUTPUT_FOLDER$, vbDirectory) <> "", OUTPUT_FOLDER$, ThisWorkbook.Path)
     folder$ = GetFolderPath("Выберите папку, куда будут помещаться созданные файлы", InitialPath$)
     If folder$ = "" Then Exit Sub
     Me.TextBox_OutputFolder = folder$
 End Sub

 Private Sub CommandButton_ChangeTemplatesFolder_Click()
     InitialPath$ = IIf(Dir(TEMPLATES_FOLDER$, vbDirectory) <> "", TEMPLATES_FOLDER$, ThisWorkbook.Path)
     folder$ = GetFolderPath("Выберите папку, содержащую шаблоны документов", InitialPath$)
     If folder$ = "" Then Exit Sub
     Me.TextBox_TemplatesFolder = folder$
 End Sub

 Private Sub Label18_Click()
     On Error Resume Next: OpenFolder TEMPLATES_FOLDER$
 End Sub
 Private Sub Label19_Click()
     On Error Resume Next: OpenFolder OUTPUT_FOLDER$
 End Sub



End Sub

 Function Settings(ByVal SettingName, Optional ByVal DefValue As Variant) As Variant
     On Error Resume Next
     Settings = GetSetting(PROJECT_NAME$, "Settings", SettingName, DefValue)
     If UseTempSettings Then
         Err.Clear: res = TempSettingsCollection(CStr(SettingName))
         If Err = 0 Then Settings = res
     End If
 End Function

 Function SettingsBoolean(ByVal SettingName, Optional ByVal DefValue As Boolean = False) As Boolean
     On Error Resume Next
     SettingsBoolean = CBool(GetSetting(PROJECT_NAME$, "Settings", SettingName, DefValue))
     If UseTempSettings Then
         Err.Clear: res = TempSettingsCollection(CStr(SettingName))
         If Err = 0 Then SettingsBoolean = CBool(res)
     End If
 End Function
 
