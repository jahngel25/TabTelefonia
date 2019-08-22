Attribute VB_Name = "CreateObjectWS"
Public Function InformacionWS(strSoap, strSOAPAction, strWsdl, strPuerto) As Object
    Dim xmlResponse As MSXML2.DOMDocument30
    Dim ObjectWs As Object
    Dim StringWs As String
    
    If (strSOAPAction = "getNumbers") Then
        StringWs = crearJsonGetNumbers(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto)
    Else
        StringWs = crearJson(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto)
    End If
    
    Set ObjectWs = JSON.parse(StringWs)
    
    Set xmlResponse = Nothing
        
    Set InformacionWS = ObjectWs

End Function


Public Function crearJsonGetNumbers(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto)

    Dim validacion As String
    validacion = "{" & Chr(34) & "item" & Count & Chr(34) & ": [{"
    dataWS = "{" & Chr(34) & "item" & Count & Chr(34) & ": [{"
    Count = 1
    
    If InvokeWebService(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto) Then
        MuestraNodosGetNumbers xmlResponse.childNodes
    Else
        resultWS.Text = "Error"
    End If
    
     If (dataWS = validacion) Then
        dataWS = ""
        MsgBox "Los filtros selecionados no traen ningun numero", vbInformation, App.Title
        Exit Sub
    Else
        dataWS = Left(dataWS, Len(dataWS) - 14) & "}]}"
    End If
    
End Function

Public Function crearJson(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto)

    dataWS = "{"

    If InvokeWebService(strSoap, strSOAPAction, strWsdl, xmlResponse, strPuerto) Then
        MuestraNodos xmlResponse.childNodes
    Else
        resultWS.Text = "Error"
    End If
    
    dataWS = Left(dataWS, Len(dataWS) - 1) & "}"
    
End Function
Public Sub MuestraNodosGetNumbers(ByRef Nodos As MSXML2.IXMLDOMNodeList)
    Dim oNodo As MSXML2.IXMLDOMNode
    
        For Each oNodo In Nodos
        If oNodo.nodeType = NODE_TEXT Then
            If (Mid(oNodo.parentNode.nodeName, 5) = "transaction_id") Then
               
            Else
                If (Mid(oNodo.parentNode.nodeName, 5) = "core_network_element") Then
                    dataWS = Left(dataWS, Len(dataWS) - 1) & "}]," & Chr(34) & "item" & countWs & Chr(34) & ": [{"
                    countWs = countWs + 1
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
Public Sub MuestraNodos(ByRef Nodos As MSXML2.IXMLDOMNodeList)
    Dim oNodo As MSXML2.IXMLDOMNode
    For Each oNodo In Nodos
      If oNodo.nodeType = NODE_TEXT Then
       dataWS = dataWS & Chr(34) & Mid(oNodo.parentNode.nodeName, 5) & Chr(34) & ":" & Chr(34) & oNodo.nodeValue & Chr(34) & ","
      End If
      If oNodo.hasChildNodes Then
        MuestraNodos oNodo.childNodes
      End If
    Next oNodo
End Sub







