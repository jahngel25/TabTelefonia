Attribute VB_Name = "Service"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function InvokeWebService(strSoap, strSOAPAction, strURL, ByRef xmlResponse, strPuerto) As Boolean

Dim xmlhttp As MSXML2.XMLHTTP30
Dim blnSuccess As Boolean
Dim TimeOutSec As Long
TimeOutSec = 4

Set xmlhttp = New MSXML2.XMLHTTP30
xmlhttp.Open "POST", strURL, False
xmlhttp.SetRequestHeader "Man", "POST " & strURL & " HTTP/1.1"
xmlhttp.SetRequestHeader "Accept-Encoding", "gzip,deflate"
xmlhttp.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
xmlhttp.SetRequestHeader "SOAPAction", strSOAPAction
xmlhttp.SetRequestHeader "Content-Length", "2124"
xmlhttp.SetRequestHeader "Host", "172.24.42.211:" & strPuerto
xmlhttp.SetRequestHeader "Connection", "Keep-Alive"
xmlhttp.SetRequestHeader "User-Agent", "Apache-HttpClient/4.1.1 (java 1.5)"
Call xmlhttp.Send(strSoap)

If xmlhttp.status = 200 Then
blnSuccess = True
Else
blnSuccess = False
End If

Set xmlResponse = xmlhttp.responseXML
InvokeWebService = blnSuccess
Set xmlhttp = Nothing
End Function

