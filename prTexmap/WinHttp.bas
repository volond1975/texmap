Attribute VB_Name = "WinHttp"
'#Region "@HTTP Request"
'Func HttpPost($sURL, $sData = "")
'   Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
'   $oHTTP.Open("POST", $sURL, False)
'   If (@error) Then Return SetError(1, 0, 0)
'   $oHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
'   $oHTTP.Send($sData)
'   If (@error) Then Return SetError(2, 0, 0)
'   If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)
'   Return SetError(0, 0, $oHTTP.ResponseText)
'EndFunc
Function HTTPPost(sURL, sRequest)
  Set oHttp = CreateObject("Microsoft.XMLHTTP")
  oHttp.Open "POST", sURL, False
  oHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  oHttp.setRequestHeader "Content-Length", Len(sRequest)
  oHttp.Send sRequest
  HTTPPost = oHttp.ResponseText
 End Function
Function HttpGet(sURL, Optional sData = "", Optional bResponse As Boolean = False)

Debug.Print sURL
Dim oHttp As Object
On Error Resume Next
  Set oHttp = CreateObject("MSXML2.XMLHTTP")
If Err.Number <> 0 Then
Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
End If
On Error GoTo 0
If oHttp Is Nothing Then Exit Function
oHttp.Open "GET", sURL, False
oHttp.Send
If bResponse Then MsgBox oHttp.ResponseText, vbInformation, "ответ"
HttpGet = oHttp.ResponseText
Set oHttp = Nothing
   
End Function
'#EndRegion

'; #FUNCTION# ;===============================================================================
'; Name...........: _WinHttpOpen
'; Description ...: Initializes the use of WinHttp functions and returns a WinHttp-session handle.
'; Syntax.........: _WinHttpOpen([$sUserAgent = Default [, $iAccessType = Default [, $sProxyName = Default [, $sProxyBypass = Default [, $iFlag = Default ]]]]])
'; Parameters ....: $sUserAgent - [optional] The name of the application or entity calling the WinHttp functions.
';                  $iAccessType - [optional] Type of access required. Default is $WINHTTP_ACCESS_TYPE_NO_PROXY.
';                  $sProxyName - [optional] The name of the proxy server to use when proxy access is specified by setting $iAccessType to $WINHTTP_ACCESS_TYPE_NAMED_PROXY. Default is $WINHTTP_NO_PROXY_NAME.
';                  $sProxyBypass - [optional] An optional list of host names or IP addresses, or both, that should not be routed through the proxy when $iAccessType is set to $WINHTTP_ACCESS_TYPE_NAMED_PROXY. Default is $WINHTTP_NO_PROXY_BYPASS.
';                  $iFlag - [optional] Integer containing the flags that indicate various options affecting the behavior of this function. Default is 0.
'; Return values .: Success - Returns valid session handle.
';                  Failure - Returns 0 and sets @error:
';                  |1 - DllCall failed
'; Author ........: trancexx
'; Remarks .......: <b>You are strongly discouraged to use WinHTTP in asynchronous mode with AutoIt. AutoIt's callback implementation can't handle reentrancy properly.</b>
';                  +For asynchronous mode set [[$iFlag]] to [[$WINHTTP_FLAG_ASYNC]]. In that case [[$WINHTTP_OPTION_CONTEXT_VALUE]] for the handle will inernally be set to [[$WINHTTP_FLAG_ASYNC]] also.
'; Related .......: _WinHttpCloseHandle, _WinHttpConnect
'; Link ..........: http://msdn.microsoft.com/en-us/library/aa384098.aspx
';============================================================================================
'Func _WinHttpOpen($sUserAgent = Default, $iAccessType = Default, $sProxyName = Default, $sProxyBypass = Default, $iFlag = Default)
'    __WinHttpDefault($sUserAgent, __WinHttpUA())
'    __WinHttpDefault($iAccessType, $WINHTTP_ACCESS_TYPE_NO_PROXY)
'    __WinHttpDefault($sProxyName, $WINHTTP_NO_PROXY_NAME)
'    __WinHttpDefault($sProxyBypass, $WINHTTP_NO_PROXY_BYPASS)
'    __WinHttpDefault($iFlag, 0)
'    Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpOpen", _
'            "wstr", $sUserAgent, _
'            "dword", $iAccessType, _
'            "wstr", $sProxyName, _
'            "wstr", $sProxyBypass, _
'            "dword", $iFlag)
'    If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
'    If $iFlag = $WINHTTP_FLAG_ASYNC Then _WinHttpSetOption($aCall[0], $WINHTTP_OPTION_CONTEXT_VALUE, $WINHTTP_FLAG_ASYNC)
'    Return $aCall[0]
'EndFunc
