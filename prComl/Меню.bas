Attribute VB_Name = "Меню"
Option Explicit

Dim Рабочая_книга As String
    
Sub CreateMenu()
'   This sub should be executed when the workbook is opened.
'   NOTE: There is no error handling in this subroutine
Dim konkurent As Range

    Dim MenuSheet As Worksheet
    Dim MenuObject As CommandBarPopup

    Dim MenuItem As Object
    Dim SubMenuItem As CommandBarButton
    Dim Row As Integer
    Dim MenuLevel, NextLevel, PositionOrMacro, Caption, Divider, FaceId, visible_on, may_booton
    
''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Location for menu data
    Set MenuSheet = ThisWorkbook.Sheets("ЛистМеню")
   
   'Set konkurent =[thisworkbook]Конкурент!currentregion
''''''''''''''''''''''''''''''''''''''''''''''''''''

'   Make sure the menus aren't duplicated
    Call DeleteMenu
    
'   Initialize the row counter
    Row = 2

'   Add the menus, menu items and submenu items using
'   data stored on MenuSheet
    
    Do Until IsEmpty(MenuSheet.Cells(Row, 1))
        With MenuSheet
            MenuLevel = .Cells(Row, 1)
            Caption = .Cells(Row, 2)
            PositionOrMacro = .Cells(Row, 3)
            Divider = .Cells(Row, 4)
            FaceId = .Cells(Row, 5)
            visible_on = .Cells(Row, 6)
            may_booton = .Cells(Row, 7)
            NextLevel = .Cells(Row + 1, 1)
        End With
        
        Select Case MenuLevel
            Case 1 ' A Menu
'              Add the top-level menu to the Worksheet CommandBar
                Set MenuObject = Application.CommandBars(1). _
                    Controls.Add(Type:=msoControlPopup, _
                    before:=PositionOrMacro, _
                    temporary:=True)
                           With MenuObject
               .Caption = Caption
               '.Enabled = visible_on
              '.ShortcutText = may_booton
                End With
            
            
            Case 2 ' A Menu Item
                If NextLevel = 3 Then
                    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlPopup)
                Else
                    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
                    MenuItem.OnAction = PositionOrMacro
                End If
                'If FaceId <> "" Then MenuItem.FaceId = FaceId
                If Divider Then MenuItem.BeginGroup = True
           
                 With MenuItem
               .Caption = Caption
               '.Enabled = visible_on
               '.ShortcutText = may_booton
                End With
            
                
                
                
                MenuItem.Enabled = visible_on
              
            Case 3 ' A SubMenu Item
                Set SubMenuItem = MenuItem.Controls.Add(Type:=msoControlButton)
                SubMenuItem.Caption = Caption
                SubMenuItem.OnAction = PositionOrMacro
                'If FaceId <> "" Then SubMenuItem.FaceId = FaceId
                If Divider Then SubMenuItem.BeginGroup = True
                With SubMenuItem
               .Enabled = visible_on
               .ShortcutText = may_booton
                End With
                
        End Select
        Row = Row + 1
    Loop
End Sub

Sub DeleteMenu()
'   This sub should be executed when the workbook is closed
'   Deletes the Menus
    Dim MenuSheet As Worksheet
    Dim Row As Integer
    Dim Caption As String
    
    On Error Resume Next
    Set MenuSheet = ThisWorkbook.Sheets("ЛистМеню")
    Row = 2
    Do Until IsEmpty(MenuSheet.Cells(Row, 1))
        If MenuSheet.Cells(Row, 1) = 1 Then
            Caption = MenuSheet.Cells(Row, 2)
            Application.CommandBars(1).Controls(Caption).Delete
        End If
        Row = Row + 1
    Loop
    On Error GoTo 0
End Sub
