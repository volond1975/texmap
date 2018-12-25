Attribute VB_Name = "cJsonExpl"
Public Sub Question1()
 
 Dim JSONText As String
 JSONText = "{""JSON"":{""JSON"":{""JSON"":{""JSON"":{""JSON"":{""JSON"":{""JSON"":{""JSON"":{""JSON"":{""JSON"":""VBA""}}}}}}}}}}"
 
 Dim JSON
 Set JSON = New cJSON
 
 Dim D As Dictionary
 Set D = JSON.Deserialize(JSONText)
 If (JSON.IsOk()) Then
 MsgBox D.Item("JSON").Item("JSON").Item("JSON").Item("JSON").Item("JSON").Item("JSON").Item("JSON").Item("JSON").Item("JSON").Item("JSON")  '          <span class="cVBAResult">'shows Value</span>
 Else
 MsgBox JSON.ShowWhyNotOk()
 End If
 Set D = Nothing
 End Sub
