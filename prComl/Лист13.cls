VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ëèñò13"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Dim icbc As Object
    For Each icbc In Application.CommandBars("cell").Controls
        If icbc.Tag = "brccm" Then icbc.Delete
    Next icbc
 
    If Not Application.Intersect(Target, Range("b1:b100")) _
         Is Nothing Then
                    With Application.CommandBars("cell").Controls _
                .Add(Type:=msoControlButton, before:=2, _
                temporary:=False)
           .Caption = "ÅÄÐÏÓ"
           .OnAction = "frmEDERPUorINNShow"
           .Tag = "brccm"
        End With
   End If


End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
