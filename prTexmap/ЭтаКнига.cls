VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
  
Public Function lCountWorkbooks() As Long
Dim lCount As Long, wbBook As Workbook
For Each wbBook In Application.Workbooks
If wbBook.Windows(1).Visible Then lCount = lCount + 1
Next wbBook
lCountWorkbooks = lCount
End Function


Private Sub Workbook_Open()
Set App = Application
 Application.OnKey "^~", "nerv_DropDownList.DropDownListShow"
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
'Call reestrshow
End Sub
