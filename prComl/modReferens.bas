Attribute VB_Name = "modReferens"
Sub RemoveReference()
For Each ref In Application.VBE.ActiveVBProject.References
If ref.name = "MyProject" Then Application.VBE.ActiveVBProject.References.Remove ref
Next
End Sub

Sub AddReference()
found = False
For Each ref In Application.VBE.ActiveVBProject.References
If ref.name = "MyProject" Then found = True
Next
If Not found Then Application.VBE.ActiveVBProject.References.AddFrom File = "c:\Program Files\000.xla"
End Sub
