VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claCreateObjectWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function InformacionWS(strSoap, strSOAPAction, strWsdl, strPuerto) As Object
    Dim xmlResponse As MSXML2.DOMDocument30
    Dim ObjectWs As Object
    Dim StringWs As String
    
    If (strSOAPAction = "getNumbers") Then
        StringWs = crearJsonGetNumbers(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto)
    Else
        StringWs = crearJsonReserveNumbers(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto)
    End If
    
    Set ObjectWs = JSON.parse(StringWs)
    
    Set xmlResponse = Nothing
        
    Set InformacionWS = ObjectWs

End Function


Public Function crearJsonGetNumbers(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto) As String
    
    countWs = 1
    dataWS = "{" & Chr(34) & "item" & countWs & Chr(34) & ": [{"
    
    If InvokeWebService(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto) Then
        MuestraNodosGetNumbers xmlResponse.childNodes
    Else
        
    End If

    dataWS = Left(dataWS, Len(dataWS) - 14) & "}]}"
    
    crearJsonGetNumbers = dataWS
    
End Function

Public Function crearJsonReserveNumbers(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto) As String

    dataWS = "{"

    If InvokeWebService(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto) Then
        MuestraNodosReserveNumbers xmlResponse.childNodes
    Else
        resultWS.Text = "Error"
    End If
    
    dataWS = Left(dataWS, Len(dataWS) - 1) & "}"
    
    crearJsonAssignNumbers = dataWS
    
End Function
Public Sub MuestraNodosGetNumbers(ByRef Nodos As MSXML2.IXMLDOMNodeList)
    Dim oNodo As MSXML2.IXMLDOMNode
    
        For Each oNodo In Nodos
        If oNodo.nodeType = NODE_TEXT Then
            If (Mid(oNodo.parentNode.nodeName, 5) = "transaction_id") Then
               
            Else
                If (Mid(oNodo.parentNode.nodeName, 5) = "core_network_element") Then
                    countWs = countWs + 1
                    dataWS = dataWS & Chr(34) & Mid(oNodo.parentNode.nodeName, 5) & Chr(34) & ":" & Chr(34) & oNodo.nodeValue & Chr(34) & ","
                    dataWS = Left(dataWS, Len(dataWS) - 1) & "}]," & Chr(34) & "item" & countWs & Chr(34) & ": [{"
                Else
                    dataWS = dataWS & Chr(34) & Mid(oNodo.parentNode.nodeName, 5) & Chr(34) & ":" & Chr(34) & oNodo.nodeValue & Chr(34) & ","
                End If
            End If
            
        End If
        If oNodo.hasChildNodes Then
            MuestraNodosGetNumbers oNodo.childNodes
        End If
        Next oNodo
End Sub
Public Sub MuestraNodosReserveNumbers(ByRef Nodos As MSXML2.IXMLDOMNodeList)
    Dim oNodo As MSXML2.IXMLDOMNode
    For Each oNodo In Nodos
      If oNodo.nodeType = NODE_TEXT Then
       dataWS = dataWS & Chr(34) & Mid(oNodo.parentNode.nodeName, 5) & Chr(34) & ":" & Chr(34) & oNodo.nodeValue & Chr(34) & ","
      End If
      If oNodo.hasChildNodes Then
        MuestraNodosReserveNumbers oNodo.childNodes
      End If
    Next oNodo
End Sub








