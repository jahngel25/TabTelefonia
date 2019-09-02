VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claPeticionNetcracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proConexion As ADODB.Connection
Public proNumeros As colNumero
Dim varCommand As ADODB.Command

Public Function ParametrosPeticionWs(evento As String, city_name As String, estado_name As String, crm_in_use As String, request_id As String, transaction_id As String, area_code As String, country_code As String, customer_name As String, customer_id As String, customer_id_type As String, account_id As String, customer_location As String, number As String, city_code As String, consecutive_number As String, quantity_numbers As String, number_mask As String, initial_number As String, final_number As String, category As String, status As String)
    
    Select Case evento
        Case "getNumbers"
            Set ParametrosPeticionWs = getNumbers(evento, city_name, estado_name, crm_in_use, request_id, transaction_id, area_code, country_code, customer_name, customer_id, customer_id_type, account_id, customer_location, number, city_code, consecutive_number, quantity_numbers, number_mask, initial_number, final_number, category, status)
            
            TipoLog = "Response"
            CodigoLog = "300"
            Set varCommand = New ADODB.Command
            varCommand.CommandType = adCmdStoredProc
            varCommand.ActiveConnection = Me.proConexion
            varCommand.CommandText = "SP_InsertarLogNetCracker"
            varCommand.Parameters.Append varCommand.CreateParameter("@vchTransaccionId", adVarChar, adParamInput, 50, transactionIdNet)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchEvento", adVarChar, adParamInput, 50, EventoLog)
            varCommand.Parameters.Append varCommand.CreateParameter("@iIncidentId", adInteger, adParamInput, , IncidentId)
            varCommand.Parameters.Append varCommand.CreateParameter("@chUseriD", adChar, adParamInput, 10, ClientId)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchNombreMaquina", adVarChar, adParamInput, 50, NombreMaquinaLog)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchTrama", adLongVarChar, adParamInput, 2147483647, TramaResponse)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchTipo", adVarChar, adParamInput, 50, TipoLog)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchCodigo", adVarChar, adParamInput, 8, CodigoLog)
            varCommand.Execute
            
        Case "reserveNumbers"
            Set ParametrosPeticionWs = reserveNumbers(evento, city_name, estado_name, crm_in_use, request_id, transaction_id, area_code, country_code, customer_name, customer_id, customer_id_type, account_id, customer_location, number, city_code, consecutive_number, quantity_numbers, number_mask, initial_number, final_number, category, status)
            
            TipoLog = "Response"
            CodigoLog = ParametrosPeticionWs.Item("code")
            Set varCommand = New ADODB.Command
            varCommand.CommandType = adCmdStoredProc
            varCommand.ActiveConnection = Me.proConexion
            varCommand.CommandText = "SP_InsertarLogNetCracker"
            varCommand.Parameters.Append varCommand.CreateParameter("@vchTransaccionId", adVarChar, adParamInput, 50, transactionIdNet)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchEvento", adVarChar, adParamInput, 50, EventoLog)
            varCommand.Parameters.Append varCommand.CreateParameter("@iIncidentId", adInteger, adParamInput, , IncidentId)
            varCommand.Parameters.Append varCommand.CreateParameter("@chUseriD", adChar, adParamInput, 10, ClientId)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchNombreMaquina", adVarChar, adParamInput, 50, NombreMaquinaLog)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchTrama", adLongVarChar, adParamInput, 2147483647, TramaResponse)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchTipo", adVarChar, adParamInput, 50, TipoLog)
            varCommand.Parameters.Append varCommand.CreateParameter("@vchCodigo", adVarChar, adParamInput, 8, CodigoLog)
            varCommand.Execute
            
    End Select
       
End Function

Public Function getNumbers(evento As String, city_name As String, estado_name As String, crm_in_use As String, request_id As String, transaction_id As String, area_code As String, country_code As String, customer_name As String, customer_id As String, customer_id_type As String, account_id As String, customer_location As String, number As String, city_code As String, consecutive_number As String, quantity_numbers As String, number_mask As String, initial_number As String, final_number As String, category As String, status As String)

    Dim ScriptCity As String
    Dim ScriptNombre As String
    Dim ScriptNumeros As String
    Dim cityCode As String
    Dim areaCode As String
    Dim strEstado As String
    Dim strproClasificacionId As String
    Dim proClasificacionDescripcion As String
    Dim strClasificacionNet As String
    Dim varResultadosCity As ADODB.Recordset
    Dim varResultadosEstado As ADODB.Recordset
    Dim varResultadosNumeros As ADODB.Recordset
    Dim resultWS As Object
    Dim classWS As claRequestWs
    Dim objetoPrueba As claRequestWs
    Dim varContadorObject As Integer
    Dim consecutiveCheck As String
    
    'consulta para traer el codigo de la cuidad
    Set varResultadosCity = New ADODB.Recordset
    ScriptCity = "SELECT " & _
                 "Ind.vchCodRegion, " & _
                 "Ind.vchIndicativo, " & _
                 "SUBSTRING (Ciu.vchCodigoCiudad , 1, 5) As vchCodigoCiudad " & _
                 "FROM " & _
                 "CT_IndicativoCiudadROC Ind " & _
                 "INNER JOIN ct_CiudadDANE Ciu " & _
                 "ON Ind.vchCityName = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(Ciu.vchCiudad, 'Á', 'A'), 'É','E'), 'Í', 'I'), 'Ó', 'O'), 'Ú','U') " & _
                 "WHERE Ind.vchCodRegion = " & Chr(39) & city_code & Chr(39) & " "
    varResultadosCity.Open ScriptCity, Me.proConexion
    
    'consulta para conparar la clasificacion de netcraker y onyx
    If (category = "NUMEROS DORADOS") Then
        strproClasificacionId = ""
        proClasificacionDescripcion = ""
        strClasificacionNet = ""
    Else
        Set varResultadosNumeros = New ADODB.Recordset
        ScriptNumeros = "SELECT cla.iClasificacionId, " & _
                    "cla.vchClasificacionNet, " & _
                    "ct.vchClasificacion " & _
                    "FROM ClasificacionNetCracker cla " & _
                    "INNER JOIN CT_clasificacion ct " & _
                    "ON cla.iClasificacionId = ct.iClasificacionId " & _
                    "WHERE cla.iClasificacionId = " & category
        varResultadosNumeros.Open ScriptNumeros, Me.proConexion
        
        strproClasificacionId = varResultadosNumeros("iClasificacionId")
        proClasificacionDescripcion = varResultadosNumeros("vchClasificacion")
        strClasificacionNet = varResultadosNumeros("vchClasificacionNet")
    End If
    
    
    'consulta para conparar el estado de netcraker y onyx
    Set varResultadosEstado = New ADODB.Recordset
    ScriptNombre = "SELECT vchEstadoNetCracker FROM EstadoNetCracker WHERE vchEstadoOnyx = " & Chr(39) & status & Chr(39)
    varResultadosEstado.Open ScriptNombre, Me.proConexion
    
    'variables necesarias para el consumo del servicio web
    cityCode = varResultadosCity("vchCodigoCiudad")
    areaCode = varResultadosCity("vchIndicativo")
    strEstado = varResultadosEstado("vchEstadoNetCracker")
    
    Set classWS = New claRequestWs
    Set classWS.proConexion = Me.proConexion
    Set objetoPrueba = New claRequestWs
    Set classWS.coleccionPrueba = New Collection
    
    If (consecutive_number = 1) Then
        consecutiveCheck = "true"
    Else
        consecutiveCheck = ""
    End If
            
    transactionIdNet = GenGUID()
    objetoPrueba.crm_in_use = crm_in_use
    classWS.coleccionPrueba.Add objetoPrueba.crm_in_use, "crm_in_use"
    objetoPrueba.request_id = transactionIdNet
    classWS.coleccionPrueba.Add objetoPrueba.request_id, "request_id"
    objetoPrueba.transaction_id = transactionIdNet
    classWS.coleccionPrueba.Add objetoPrueba.transaction_id, "transaction_id"
    objetoPrueba.city_code = cityCode
    classWS.coleccionPrueba.Add objetoPrueba.city_code, "city_code"
    objetoPrueba.country_code = country_code
    classWS.coleccionPrueba.Add objetoPrueba.country_code, "country_code"
    objetoPrueba.area_code = areaCode
    classWS.coleccionPrueba.Add objetoPrueba.area_code, "area_code"
    objetoPrueba.consecutive_number = consecutiveCheck
    classWS.coleccionPrueba.Add objetoPrueba.consecutive_number, "consecutive_number"
    If (quantity_numbers = "") Then
        objetoPrueba.quantity_numbers = "1"
    Else
        objetoPrueba.quantity_numbers = quantity_numbers
    End If
    classWS.coleccionPrueba.Add objetoPrueba.quantity_numbers, "quantity_numbers"
    objetoPrueba.number_mask = number_mask
    classWS.coleccionPrueba.Add objetoPrueba.number_mask, "number_mask"
    objetoPrueba.initial_number = initial_number
    classWS.coleccionPrueba.Add objetoPrueba.initial_number, "initial_number"
    objetoPrueba.final_number = final_number
    classWS.coleccionPrueba.Add objetoPrueba.final_number, "final_number"
    objetoPrueba.category = strClasificacionNet
    classWS.coleccionPrueba.Add objetoPrueba.category, "category"
    objetoPrueba.status = strEstado
    classWS.coleccionPrueba.Add objetoPrueba.status, "status"
    classWS.coleccionPrueba.Add objetoPrueba
    Set resultWS = classWS.RequestPeticionWs(evento)
    
    If (dataWS = "") Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    Screen.MousePointer = 11
    Set proNumeros = New colNumero

    For varContadorObject = 1 To resultWS.Count
                        
        Dim getNumero As String
        Dim fechaActual As String
        Dim varCtNumeros As ADODB.Recordset
        Dim ScriptCtNumeros As String
        Dim chUpdateBy As String
        Dim dtUpdateDate As String
                            
        fechaActual = Now
        getNumero = resultWS.Item("item" & varContadorObject).Item(1).Item("number")
        
        Set varCtNumeros = New ADODB.Recordset
        ScriptCtNumeros = "SELECT chUpdateBy, dtUpdateDate FROM CT_Numeros WHERE vchNumero = " & Chr(39) & Mid(getNumero, 4) & Chr(39)
        varCtNumeros.Open ScriptCtNumeros, Me.proConexion
        
        If Not varCtNumeros.EOF Then
            chUpdateBy = varCtNumeros("chUpdateBy")
            dtUpdateDate = varCtNumeros("dtUpdateDate")
        Else
            chUpdateBy = fechaActual
            dtUpdateDate = UserName
        End If
        
        proNumeros.Add Me.proConexion, _
                        1, _
                        dtUpdateDate, _
                        chUpdateBy, _
                        strproClasificacionId, _
                        proClasificacionDescripcion, _
                        estado_name, _
                        status, _
                        Mid(getNumero, 4), _
                        city_name, _
                        city_code, _
                        "", _
                        "", _
                        "", _
                        "", _
                        ""
    Next varContadorObject
    
    Set getNumbers = proNumeros
        
      
End Function

Public Function reserveNumbers(evento As String, city_name As String, estado_name As String, crm_in_use As String, request_id As String, transaction_id As String, area_code As String, country_code As String, customer_name As String, customer_id As String, customer_id_type As String, account_id As String, customer_location As String, number As String, city_code As String, consecutive_number As String, quantity_numbers As String, number_mask As String, initial_number As String, final_number As String, category As String, status As String)

    Dim classWS As claRequestWs
    Dim objetoPrueba As claRequestWs
    Dim varContadorObject As Integer
    Dim resultWS As Object
    Dim strEstado As String
    Dim ScriptNombre As String
    Dim varResultadosEstado As ADODB.Recordset
    
    Set classWS = New claRequestWs
    Set classWS.proConexion = Me.proConexion
    Set objetoPrueba = New claRequestWs
    Set classWS.coleccionPrueba = New Collection
    
    Set varResultadosEstado = New ADODB.Recordset
    ScriptNombre = "SELECT vchEstadoNetCracker FROM EstadoNetCracker WHERE vchEstadoOnyx = " & Chr(39) & status & Chr(39)
    varResultadosEstado.Open ScriptNombre, Me.proConexion
    strEstado = varResultadosEstado("vchEstadoNetCracker")
        
    objetoPrueba.crm_in_use = crm_in_use
    classWS.coleccionPrueba.Add objetoPrueba.crm_in_use, "crm_in_use"
    objetoPrueba.request_id = transactionIdNet
    classWS.coleccionPrueba.Add objetoPrueba.request_id, "request_id"
    objetoPrueba.transaction_id = transactionIdNet
    classWS.coleccionPrueba.Add objetoPrueba.transaction_id, "transaction_id"
    objetoPrueba.country_code = country_code
    classWS.coleccionPrueba.Add objetoPrueba.country_code, "country_code"
    objetoPrueba.area_code = area_code
    classWS.coleccionPrueba.Add objetoPrueba.area_code, "area_code"
    objetoPrueba.status = strEstado
    classWS.coleccionPrueba.Add objetoPrueba.status, "status"
    objetoPrueba.customer_name = customer_name
    classWS.coleccionPrueba.Add objetoPrueba.customer_name, "customer_name"
    objetoPrueba.customer_id = customer_id
    classWS.coleccionPrueba.Add objetoPrueba.customer_id, "customer_id"
    objetoPrueba.customer_id_type = customer_id_type
    classWS.coleccionPrueba.Add objetoPrueba.customer_id_type, "customer_id_type"
    objetoPrueba.account_id = account_id
    classWS.coleccionPrueba.Add objetoPrueba.account_id, "account_id"
    objetoPrueba.customer_location = customer_location
    classWS.coleccionPrueba.Add objetoPrueba.customer_location, "customer_location"
    objetoPrueba.number = number
    classWS.coleccionPrueba.Add objetoPrueba.number, "number"
    classWS.coleccionPrueba.Add objetoPrueba
    Set resultWS = classWS.RequestPeticionWs(evento)
    
    If (dataWS = "") Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
   
    Set reserveNumbers = resultWS

    
End Function

Function GenGUID() As String
    Dim MyGUID(35) As Byte
    Dim DashNum As Byte
    Dim FourNum As Byte
    Dim i As Integer
    
    
    For i = 0 To 7
        MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i

    For i = 9 To 13
        MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i

    For i = 15 To 17
        MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i

    For i = 19 To 22
        MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i

    For i = 24 To 35
        MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i

    MyGUID(8) = Asc("-")
    MyGUID(14) = Asc("-")
    MyGUID(18) = Asc("-")
    MyGUID(23) = Asc("-")

    GenGUID = StrConv(MyGUID, vbUnicode)
    
End Function

