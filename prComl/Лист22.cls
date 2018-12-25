VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист22"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Call ИмпортИзТехкарти
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Dim icbc As Object
    For Each icbc In Application.CommandBars("cell").Controls
        If icbc.Tag = "printtexkart" Then icbc.Delete
    Next icbc
 
    If Not Application.Intersect(Target, Worksheets("Техкарта").Range("b1:b100")) Is Nothing Then
       
                    With Application.CommandBars("cell").Controls _
                .Add(Type:=msoControlButton, before:=2, _
                temporary:=False)
           .Caption = "Печать текарт"
           .OnAction = "PrintSelectionTexkart"
           .Tag = "printtexkart"
        End With
   End If
Application.ScreenUpdating = True

End Sub
