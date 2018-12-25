VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVikonavec 
   OleObjectBlob   =   "frmVikonavec.frx":0000
   Caption         =   "UserForm2"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9270
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   12
End
Attribute VB_Name = "frmVikonavec"
Attribute VB_Base = "0{D8F4D880-0303-426D-971B-8E6E2C0225B2}{8A9EE54C-0BCD-464A-AA9A-4B4FC4388A3A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub lbParametr_Click()
 Me.tbValue = ЗначениеУмнойТаблицы("ВиконавецФорма", "ВиконавецФорма", "Наименование", "Значение", lbParametr.value)
Me.tbComment = ЗначениеУмнойТаблицы("ВиконавецФорма", "ВиконавецФорма", "Наименование", "Пример", lbParametr.value)
End Sub

Private Sub tbValue_Change()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Set lo = myListObject(ThisWorkbook, "ВиконавецФорма", "ВиконавецФорма")
Call ListObject_to_Listbox_or_combobox(ThisWorkbook, lo, "frmVikonavec", "lbParametr", "Наименование")
End Sub
