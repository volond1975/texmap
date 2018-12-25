VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressIndicator"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public FP As New F_Progress

Public CurActionIndex As Single, CurActionCount As Single
Public ShowPercents As Boolean
Const ProgressBar_Default_Width = 300
Private PrS As Integer, PrE As Integer

Private Sub Class_Initialize()
    Set FP = New F_Progress: ShowPercents = False
    PrS = 0: PrE = 100: Set_ProgressBar 0: FP.PrBar.Caption = ""
End Sub

Sub Show(ByVal Caption As String)
    FP.Caption = Caption: 'ShowCloseButton FP, True
    FP.Tag = Caption: FP.Show: FP.Repaint: DoEvents
End Sub

Sub Hide()
    Unload FP
End Sub

Sub Move(ByVal Position As Integer)
    If Abs(Position) > 3 Then Exit Sub
    FP.Top = FP.Top + FP.Height * Position
End Sub

Public Property Let line1(ByVal NewValue As String): FP.L1.Caption = NewValue: End Property
Public Property Let line2(ByVal NewValue As String): FP.L2.Caption = NewValue: End Property
Public Property Let line3(ByVal NewValue As String): FP.L3.Caption = NewValue: End Property
Public Property Let Caption(ByVal NewValue As String): FP.Caption = NewValue: End Property




' функции для Progress Bar
Sub CurAction(ByVal Action_Ind As Single, ByVal Actions_Count As Single, _
              Optional ByVal L1_txt As String = "", Optional ByVal L2_txt As String = "", _
              Optional ByVal L3_txt As String = "")


    If L1_txt <> "" Then FP.L1.Caption = L1_txt
    If L2_txt <> "" Then FP.L2.Caption = L2_txt
    If L3_txt <> "" Then FP.L3.Caption = L3_txt

    If CurActionCount > 0 Then
        If CurActionIndex > CurActionCount Then CurActionIndex = CurActionCount
        Percent = PrS + (PrE - PrS) * (CurActionIndex / CurActionCount)
    Else
        If Actions_Count <= 0 Then Exit Sub
        If Action_Ind > Actions_Count Then Action_Ind = Actions_Count
        Percent = PrS + (PrE - PrS) * (Action_Ind / Actions_Count)
    End If
    CurActionIndex = CurActionIndex + 1
    Set_ProgressBar Percent: DoEvents
End Sub
Sub StartNewAction(Optional ByVal Pr_Start As Integer = 0, Optional ByVal Pr_End As Integer = 100, _
                   Optional ByVal L1_txt As String = "", Optional ByVal L2_txt As String = "", Optional ByVal L3_txt As String = "")

    PrS = Pr_Start: PrE = Pr_End: CurActionIndex = 1: CurActionCount = 0
    If L1_txt <> "" Then FP.L1.Caption = L1_txt
    If L2_txt <> "" Then FP.L2.Caption = L2_txt
    If L3_txt <> "" Then FP.L3.Caption = L3_txt
    Set_ProgressBar PrS
End Sub
Sub Set_ProgressBar(ByVal Percent As Double)
    If Percent > 100 Then Percent = 100
    If Percent < 0 Then Percent = 0
    FP.PrBar.Width = Int(Percent * ProgressBar_Default_Width / 100)
    If ShowPercents Then FP.Caption = Fix(Percent) & " %    " & FP.Tag
    FP.Repaint
End Sub
