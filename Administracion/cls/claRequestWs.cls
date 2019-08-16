VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claRequestWs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection
Public crm_in_use As String
Public request_id As String
Public transaction_id As String
Public area_code As String
Public country_code As String
Public customer_name As String
Public customer_id As String
Public customer_id_type As String
Public account_id As String
Public customer_location As String
Public number As String
Public city_code As String
Public consecutive_number As String
Public quantity_numbers As String
Public number_mask As String
Public initial_number As String
Public final_number As String
Public category As String
Public status As String
Public coleccionPrueba As Collection



Public Function RequestPeticionWs(tipo As String) As Object
    
    Dim textSOAP As String
    Dim varComando As String
    Dim varResultados As ADODB.Recordset
    Dim ObjectDataWs As Object
    Dim Script As String
    Dim classWS As claCreateObjectWS
    Dim strWsdl As String
    
    On Error GoTo ErrorManager
    
    Set varResultados = New ADODB.Recordset
    
    Script = " SELECT * " & _
                 "FROM AtributosSoapWebService " & _
                 "WHERE vchMetodo = " & Chr(39) & tipo & Chr(39) & " "
    
    varResultados.Open Script, Me.proConexion
    
    textSOAP = "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34)
    textSOAP = textSOAP & " xmlns:v1=" & Chr(34) & "http://www.amx.com/Schema/Operation/" & tipo & "/V1.0" & Chr(34)
    textSOAP = textSOAP & " xmlns:v11=" & Chr(34) & "http://www.amx.com/CO/Schema/ClaroHeaders/v1" & Chr(34) & ">"
    textSOAP = textSOAP & "<soapenv:Header/>"
    textSOAP = textSOAP & "<soapenv:Body>"
    textSOAP = textSOAP & "<v1:" & tipo & ">"
    textSOAP = textSOAP & "<v1:headerRequest>"
    textSOAP = textSOAP & "<v11:transactionId>string</v11:transactionId>"
    textSOAP = textSOAP & "<v11:system>string</v11:system>"
    textSOAP = textSOAP & "<v11:target>string</v11:target>"
    textSOAP = textSOAP & "<v11:user>string</v11:user>"
    textSOAP = textSOAP & "<v11:password>string</v11:password>"
    textSOAP = textSOAP & "<v11:requestDate>2008-09-28T20:49:45</v11:requestDate>"
    textSOAP = textSOAP & "<v11:ipApplication>string</v11:ipApplication>"
    textSOAP = textSOAP & "<v11:traceabilityId>string</v11:traceabilityId>"
    textSOAP = textSOAP & "</v1:headerRequest>"
    
    'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
    While varResultados.EOF = False
        If (varResultados("vchMetododAtributo") = "number") Then
            textSOAP = textSOAP & "<v1:list_of_numbers>"
            textSOAP = textSOAP & "<v1:number_element>"
            If (coleccionPrueba(varResultados("vchMetododAtributo")) <> "") Then
                textSOAP = textSOAP & "<v1:" & varResultados("vchMetododAtributo") & ">" & coleccionPrueba(varResultados("vchMetododAtributo")) & "</v1:" & varResultados("vchMetododAtributo") & ">"
            End If
            textSOAP = textSOAP & "</v1:number_element>"
            textSOAP = textSOAP & "</v1:list_of_numbers>"
        Else
            If (coleccionPrueba(varResultados("vchMetododAtributo")) <> "") Then
                textSOAP = textSOAP & "<v1:" & varResultados("vchMetododAtributo") & ">" & coleccionPrueba(varResultados("vchMetododAtributo")) & "</v1:" & varResultados("vchMetododAtributo") & ">"
            End If
        End If
        
        'Siguiente registro del cursor
        varResultados.MoveNext
    Wend
    
    textSOAP = textSOAP & "</v1:" & tipo & ">"
    textSOAP = textSOAP & "</soapenv:Body>"
    textSOAP = textSOAP & "</soapenv:Envelope>"

    Set classWS = New claCreateObjectWS
    strWsdl = "http://172.24.42.211:8002/NetCracker/V2.0"
    Set ObjectDataWs = classWS.InformacionWS(textSOAP, tipo, strWsdl, 1563)
        
    Set RequestPeticionWs = ObjectDataWs
    
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

