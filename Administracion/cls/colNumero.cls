VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claNumero"
Attribute VB_Ext_KEY = "Member0" ,"claNumero"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Se agrega la funcion MetVerificarAsignacionOnyx
'       VERSION:       1.0.000
'       REQUERIMIENTO: 5290
'       FECHA:       2009/09/02
'*******************************************************************
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Se agregan los metodos
'       MetCambiarEstadoNumero, FunGAgregarACollection y SubSumarEnIndex
'       REQUERIMIENTO:          5322
'       VERSION:       1.0.100
'       FECHA:       2009/10/10
'*******************************************************************
Option Explicit

Public proRegionCode As String
Public proNumeroInicial As String
Public proNumeroFinal As String
Public proCantidadNumeros As String
Public proReclasificarTodos As String
Public proUsuario As String
Public proEstado As String

Public proUsarConjuntoClasificaciones As String
Public proClasificacion As colClasificacion
Public proConexion As ADODB.Connection
Public proLogNumero As colLogNumero

Private varConexion As ADODB.Connection

Public proSeleccionados As Integer

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proRecordStatus As String, _
                    proUpdateDate As String, proUpdateBy As String, _
                    proClasificacionId As String, proClasificacionDescripcion As String, _
                    proEstadoNumeroDescripcion As String, proEstadoNumero As String, _
                    proNumero As String, proRegionCodeDescripcion As String, _
                    proRegionCode As String, proUserIdAprobador As String, proFechaAprobacion As String, _
                    proIncidentId As String, proCompanyId As String, proCompanyName As String, Optional sKey As String) As claNumero
    'create a new object
    Dim objNewMember As claNumero
    Set objNewMember = New claNumero


    
    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
        
    objNewMember.proRecordStatus = proRecordStatus
    objNewMember.proUpdateDate = proUpdateDate
    objNewMember.proUpdateBy = proUpdateBy
    objNewMember.proClasificacionDescripcion = proClasificacionDescripcion
    objNewMember.proClasificacionId = proClasificacionId
    objNewMember.proEstadoNumeroDescripcion = proEstadoNumeroDescripcion
    objNewMember.proEstadoNumero = proEstadoNumero
    objNewMember.proNumero = proNumero
    objNewMember.proRegionCodeDescripcion = proRegionCodeDescripcion
    objNewMember.proRegionCode = proRegionCode
    objNewMember.proUserIdAprobador = proUserIdAprobador
    objNewMember.proFechaAprobacion = proFechaAprobacion
    objNewMember.proIncidentId = proIncidentId
    objNewMember.proCompanyId = proCompanyId
    objNewMember.proCompanyName = proCompanyName
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claNumero
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Me.proSeleccionados = 0
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Function MetGenerarNumeros() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    If Not MetAbrirConexion Then
        MetGenerarNumeros = False
        Exit Function
    End If
    
    varComando = "EXEC CP_GenerarNumerosPublicos '" & _
                    Me.proRegionCode & "', '" & _
                    Me.proNumeroInicial & "', '" & _
                    Me.proNumeroFinal & "', '" & _
                    Me.proUsuario & "'"
        
    varConexion.Execute varComando, , adAsyncExecute
    
    MetGenerarNumeros = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Private Function MetAbrirConexion() As Boolean
    On Error GoTo ErrManager
    
    Set varConexion = New ADODB.Connection
    varConexion.ConnectionString = Me.proConexion.ConnectionString
    varConexion.CommandTimeout = 0
    varConexion.ConnectionTimeout = 0
    
    varConexion.Open
    
    MetAbrirConexion = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
    
Public Function MetValidarReclasificacion() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    
        varComando = "Select    Count(1)    " & _
                     "From      CT_Numeros  " & _
                     "Where     tiRecordStatus = 1 "
                     
    If Me.proReclasificarTodos = "N" Then
        varComando = varComando & "And chRegionCode = '" & Me.proRegionCode & "' " _
                                & "And convert(int, vchNumero) between " & Me.proNumeroInicial & " and " & Me.proNumeroFinal
    End If
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields(0)) Then
            Me.proCantidadNumeros = 0
        Else
            Me.proCantidadNumeros = varResultado.Fields(0)
        End If
    Else
        Me.proCantidadNumeros = 0
    End If
    
    MetValidarReclasificacion = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetReclasificarNumeros() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    If Not MetAbrirConexion Then
        MetReclasificarNumeros = False
        Exit Function
    End If
    
    varComando = "EXEC CP_ClasificacionNumeros '" & _
                    Me.proRegionCode & "', " & _
                    Val(Me.proNumeroInicial) & ", " & _
                    Val(Me.proNumeroFinal) & " "
        
    varConexion.Execute varComando, , adAsyncExecute
    
    MetReclasificarNumeros = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarNumeros() As Boolean
    Dim varComando As ADODB.Command
    Dim varComandoAux As String
    Dim varEntro As Boolean
    Dim varContador As Integer
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claNumero
    Dim varCantidadClasificaciones As Integer
    On Error GoTo ErrManager
    
    Set varComando = New ADODB.Command
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_ConsultaNumeros"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@vchRegionCode").Value = Me.proRegionCode
    varComando.Parameters.Item("@chEstadoNumero").Value = Me.proEstado
    varComando.Parameters.Item("@vchNumeroInicial").Value = Me.proNumeroInicial
    varComando.Parameters.Item("@vchNumeroFinal").Value = Me.proNumeroFinal
    varComando.Parameters.Item("@iCantidadNumeros").Value = Val(Me.proCantidadNumeros)
    
    If Me.proClasificacion.Count <> 0 Then
        varCantidadClasificaciones = 0
        For varContador = 1 To Me.proClasificacion.Count
            If Me.proClasificacion.Item(varContador).proSeleccionado = "S" Then
                varComandoAux = varComandoAux & "and sum (case when iClasificacionId = " & Me.proClasificacion.Item(varContador).proClasificacionId & " then 1 else 0 end) > = 1 "
                varEntro = True
                varCantidadClasificaciones = varCantidadClasificaciones + 1
            End If
        Next varContador
    End If
    
    If varComandoAux <> "" Then
        varComando.Parameters.Item("@vchCodigosClasificacion").Value = varComandoAux
    Else
        varComando.Parameters.Item("@vchCodigosClasificacion").Value = ""
    End If
    varComando.Parameters.Item("@iCantidadClasificaciones").Value = Val(varCantidadClasificaciones)
    varComando.Parameters.Item("@cUsarClasificacionConjunto").Value = Me.proUsarConjuntoClasificaciones
    
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute
    
    Set varNumero = New claNumero
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varNumero.proRegionCode = ""
        Else
            varNumero.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("chRegionName")) Then
            varNumero.proRegionCodeDescripcion = ""
        Else
            varNumero.proRegionCodeDescripcion = Trim(varResultado.Fields("chRegionName"))
        End If
        
        If IsNull(varResultado.Fields("vchNumero")) Then
            varNumero.proNumero = ""
        Else
            varNumero.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        If IsNull(varResultado.Fields("chEstadoNumero")) Then
            varNumero.proEstadoNumero = ""
        Else
            varNumero.proEstadoNumero = Trim(varResultado.Fields("chEstadoNumero"))
        End If
        
        If IsNull(varResultado.Fields("vchDescripcionEstado")) Then
            varNumero.proEstadoNumeroDescripcion = ""
        Else
            varNumero.proEstadoNumeroDescripcion = Trim(varResultado.Fields("vchDescripcionEstado"))
        End If
        
        If varCantidadClasificaciones <> 0 Then
            If IsNull(varResultado.Fields(5)) Then
                varNumero.proClasificacionDescripcion = ""
            Else
                varNumero.proClasificacionDescripcion = Trim(varResultado.Fields(5))
            End If
        Else
            varNumero.proClasificacionDescripcion = ""
        End If
        
        If IsNull(varResultado.Fields("chUpdateBy")) Then
            varNumero.proUpdateBy = ""
        Else
            varNumero.proUpdateBy = Trim(varResultado.Fields("chUpdateBy"))
        End If
        
        If IsNull(varResultado.Fields("dtUpdateDate")) Then
            varNumero.proUpdateDate = ""
        Else
            varNumero.proUpdateDate = Trim(varResultado.Fields("dtUpdateDate"))
        End If
        If IsNull(varResultado.Fields("chUserIdAprobador")) Then
            varNumero.proUserIdAprobador = ""
        Else
            varNumero.proUserIdAprobador = Trim(varResultado.Fields("chUserIdAprobador"))
        End If
        If IsNull(varResultado.Fields("dtFechaAprobacion")) Then
            varNumero.proFechaAprobacion = ""
        Else
            varNumero.proFechaAprobacion = Trim(varResultado.Fields("dtFechaAprobacion"))
        End If
        
        varNumero.proIncidentId = ""
        varNumero.proCompanyId = "0"
        varNumero.proCompanyName = ""
        
        Me.Add Me.proConexion, _
               1, _
               varNumero.proUpdateDate, _
               varNumero.proUpdateBy, _
               varNumero.proClasificacionId, _
               varNumero.proClasificacionDescripcion, _
               varNumero.proEstadoNumeroDescripcion, _
               varNumero.proEstadoNumero, _
               varNumero.proNumero, _
               varNumero.proRegionCodeDescripcion, _
               varNumero.proRegionCode, _
               varNumero.proUserIdAprobador, _
               varNumero.proFechaAprobacion, _
               varNumero.proIncidentId, _
               varNumero.proCompanyId, _
               varNumero.proCompanyName
              
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varNumero = Nothing
    
    MetConsultarNumeros = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarNumerosSinAprobacion(iCompanyId As String) As Boolean
    Dim varComando As ADODB.Command
    Dim varComandoAux As String
    Dim varEntro As Boolean
    Dim varContador As Integer
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claNumero
    Dim varCantidadClasificaciones As Integer
    Dim i As Integer
    On Error GoTo ErrManager
    
    
    
    Set varComando = New ADODB.Command
    
     'Limpia la coleccion
    If Me.Count > 0 Then
        For i = Me.Count To 1 Step -1
            Me.Remove (i)
        Next i
    End If
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_ConsultaNumerosSinAprobacion"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@chEstadoNumero").Value = Me.proEstado
    varComando.Parameters.Item("@vchNumeroInicial").Value = Me.proNumeroInicial
    varComando.Parameters.Item("@vchNumeroFinal").Value = Me.proNumeroFinal
    varComando.Parameters.Item("@chUserId").Value = Me.proUsuario
    varComando.Parameters.Item("@iCompanyId").Value = iCompanyId
    
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute
    
    Set varNumero = New claNumero
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varNumero.proRegionCode = ""
        Else
            varNumero.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("chRegionName")) Then
            varNumero.proRegionCodeDescripcion = ""
        Else
            varNumero.proRegionCodeDescripcion = Trim(varResultado.Fields("chRegionName"))
        End If
        
        If IsNull(varResultado.Fields("vchNumero")) Then
            varNumero.proNumero = ""
        Else
            varNumero.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        If IsNull(varResultado.Fields("chEstadoNumero")) Then
            varNumero.proEstadoNumero = ""
        Else
            varNumero.proEstadoNumero = Trim(varResultado.Fields("chEstadoNumero"))
        End If
        
        If IsNull(varResultado.Fields("vchDescripcionEstado")) Then
            varNumero.proEstadoNumeroDescripcion = ""
        Else
            varNumero.proEstadoNumeroDescripcion = Trim(varResultado.Fields("vchDescripcionEstado"))
        End If
        
        
        If IsNull(varResultado.Fields(5)) Then
            varNumero.proClasificacionDescripcion = ""
        Else
            varNumero.proClasificacionDescripcion = Trim(varResultado.Fields(5))
        End If
        
        
        If IsNull(varResultado.Fields("chUpdateBy")) Then
            varNumero.proUpdateBy = ""
        Else
            varNumero.proUpdateBy = Trim(varResultado.Fields("chUpdateBy"))
        End If
        
        If IsNull(varResultado.Fields("dtUpdateDate")) Then
            varNumero.proUpdateDate = ""
        Else
            varNumero.proUpdateDate = Trim(varResultado.Fields("dtUpdateDate"))
        End If
        If IsNull(varResultado.Fields("chUserIdAprobador")) Then
            varNumero.proUserIdAprobador = ""
        Else
            varNumero.proUserIdAprobador = Trim(varResultado.Fields("chUserIdAprobador"))
        End If
        If IsNull(varResultado.Fields("dtFechaAprobacion")) Then
            varNumero.proFechaAprobacion = ""
        Else
            varNumero.proFechaAprobacion = Trim(varResultado.Fields("dtFechaAprobacion"))
        End If
        
        If IsNull(varResultado.Fields("iIncidentId")) Then
            varNumero.proIncidentId = ""
        Else
            varNumero.proIncidentId = Trim(varResultado.Fields("iIncidentId"))
        End If
        If IsNull(varResultado.Fields("iCompanyId")) Then
            varNumero.proCompanyId = "0"
        Else
            varNumero.proCompanyId = Trim(varResultado.Fields("iCompanyId"))
        End If
        If IsNull(varResultado.Fields("vchCompanyName")) Then
            varNumero.proCompanyName = ""
        Else
            varNumero.proCompanyName = Trim(varResultado.Fields("vchCompanyName"))
        End If
        
        
        Me.Add Me.proConexion, _
               1, _
               varNumero.proUpdateDate, _
               varNumero.proUpdateBy, _
               varNumero.proClasificacionId, _
               varNumero.proClasificacionDescripcion, _
               varNumero.proEstadoNumeroDescripcion, _
               varNumero.proEstadoNumero, _
               varNumero.proNumero, _
               varNumero.proRegionCodeDescripcion, _
               varNumero.proRegionCode, _
               varNumero.proUserIdAprobador, _
               varNumero.proFechaAprobacion, _
               varNumero.proIncidentId, _
               varNumero.proCompanyId, _
               varNumero.proCompanyName
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varNumero = Nothing
    
    MetConsultarNumerosSinAprobacion = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Function MetClasificacionManual(iClasificacionId As Integer, iModo As Integer, ByRef strError As String) As Integer
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claNumero
    Dim i As Integer
    Dim iContador As Integer
    Dim iError As Integer
    On Error GoTo ErrManager
    
    
    Set varComando = New ADODB.Command
    
    If Not MetAbrirConexion Then
        MetClasificacionManual = 0
        Exit Function
    End If
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_ClasificacionNumeroManual"
    varComando.CommandTimeout = 0
    varComando.Parameters.Append varComando.CreateParameter("Return", adInteger, adParamReturnValue, , 0)
    varComando.Parameters.Append varComando.CreateParameter("chRegionCode", adChar, adParamInput, 10, "")
    varComando.Parameters.Append varComando.CreateParameter("vchNumero", adVarChar, adParamInput, 50, "")
    varComando.Parameters.Append varComando.CreateParameter("iClasificacionId", adInteger, adParamInput, 4, iClasificacionId)
    varComando.Parameters.Append varComando.CreateParameter("iModo", adInteger, adParamInput, 4, iModo)
    
    iContador = 0
    iError = 0
    strError = ""
    For Each varNumero In Me
        If varNumero.proSeleccionado = "S" Then
            varComando.Parameters.Item("chRegionCode").Value = varNumero.proRegionCode
            varComando.Parameters.Item("vchNumero").Value = varNumero.proNumero
            Set varResultado = varComando.Execute
            If Not varResultado Is Nothing Then
                If varResultado.State <> adStateClosed Then
                    If Not varResultado.EOF Then
                        varResultado.GetRows
                    End If
                End If
            End If
            Set varResultado = Nothing
            If varComando.Parameters.Item("Return").Value > 0 Then
                iContador = iContador + 1
            Else
                iError = varComando.Parameters.Item("Return").Value
                If iError <= 0 Then
                    If strError <> "" Then strError = strError & ", "
                    strError = strError & varNumero.proNumero
                End If
            End If
        End If
    Next
    
    MetClasificacionManual = iContador
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetVerificarAsignacionOnyx(ByRef strError As String) As Boolean
'*******************************************************************
'       DESCRIPCION: Funcion que verifica si los nros. libres
'       seleccionados para asignar a un cliente estan disponibles
'       en softswitch
'
'       PARAMETROS:
'           -         strError String que devuelve el mensaje de error
'       RETORNO:
'           -        retorna true o false si el nro. esta libre o no
'                    en el softswitch
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Desarrollo inicial de la funcion
'       VERSION:       1.0.000
'       REQUERIMIENTO: 5290
'       FECHA:       2009/09/02
'*******************************************************************

    Dim varComando As ADODB.Command
    Dim varNumero As claNumero
    Dim paramNumeros As String
    Dim parEtiqueta As String
    
    On Error GoTo ErrManager
    

    Set varComando = New ADODB.Command
    Set varComando.ActiveConnection = Me.proConexion
    varComando.CommandType = adCmdStoredProc
    varComando.CommandText = "CRMCP_VerificarAsignacionOnyx"
    
    For Each varNumero In Me
        If varNumero.proSeleccionado = "S" Then
            paramNumeros = paramNumeros + varNumero.proNumero & ","
        End If
    Next
    
    varComando.Parameters.Item("@numerosList").Value = Mid$(paramNumeros, 1, Len(paramNumeros) - 1)
    varComando.Parameters.Item("@vchMensaje").Value = strError
    varComando.Execute
    
    If varComando.Parameters.Item(0) = 1 Then
        strError = varComando.Parameters.Item("@vchMensaje").Value
        MetVerificarAsignacionOnyx = False
    Else
        MetVerificarAsignacionOnyx = True
    End If
    
    
    Set varComando = Nothing

    Exit Function
ErrManager:
     SubGMuestraError
End Function

Public Function MetCambiarEstadoNumeros(chEstadoDestino As String, strEstadoDescripcion As String, strLogin As String) As String

    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claNumero
    Dim iError As Integer
    Dim strError As String
    Dim iContador As Integer
    Dim iPosicion As Integer
    Dim varMensajeResultado As String
    Dim varCollectionCantidades As Collection
    Dim varCollectionEstados As Collection
    Dim varTemp As Variant
    Dim i As Integer
    
    On Error GoTo ErrManager
    
    Set varComando = New ADODB.Command
    Set varCollectionCantidades = New Collection
    Set varCollectionEstados = New Collection
            
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CRMCP_CambiarEstadoNumero"
    varComando.CommandTimeout = 0
    varComando.Parameters.Append varComando.CreateParameter("Return", adInteger, adParamReturnValue, , 0)
    varComando.Parameters.Append varComando.CreateParameter("vchNumero", adVarChar, adParamInput, 50, "")
    varComando.Parameters.Append varComando.CreateParameter("chEstadoDestino", adChar, adParamInput, 10, "")
    varComando.Parameters.Append varComando.CreateParameter("vchUser", adVarChar, adParamInput, 30, "")
    varComando.Parameters.Append varComando.CreateParameter("vchMensaje", adVarChar, adParamOutput, 50, "")
    varComando.Parameters.Append varComando.CreateParameter("vchFechaActualizacion", adVarChar, adParamOutput, 30, "")
    varComando.Parameters.Append varComando.CreateParameter("chEstadoDestinoOutput", adChar, adParamOutput, 10, "")
    varComando.Parameters.Append varComando.CreateParameter("vchDescripcionEstadoOutput", adVarChar, adParamOutput, 255, "")
    varComando.Parameters.Append varComando.CreateParameter("vchUserOutput", adVarChar, adParamOutput, 30, "")
    
    
    iError = 0
    strError = ""
    varComando.Parameters.Item("chEstadoDestino").Value = chEstadoDestino
    varComando.Parameters.Item("vchUser").Value = strLogin
    For Each varNumero In Me
        If varNumero.proSeleccionado = "S" Then
            varComando.Parameters.Item("vchNumero").Value = varNumero.proNumero
            Set varResultado = varComando.Execute
            If Not varResultado Is Nothing Then
                If varResultado.State <> adStateClosed Then
                    If Not varResultado.EOF Then
                        varResultado.GetRows
                    End If
                End If
            End If
            Set varResultado = Nothing
            If varComando.Parameters.Item("Return").Value > 0 Then
                ' se ejecuto exitosamente, ahora hay que agregar el mensaje a la coleccion
                varNumero.proObservacionesCambioEstado = varComando.Parameters.Item("vchMensaje").Value
                If Left(varNumero.proObservacionesCambioEstado, 5) <> "Error" Then
                    iPosicion = FunGAgregarACollection(varNumero.proEstadoNumeroDescripcion, varCollectionEstados)
                    Call SubSumarEnIndex(varCollectionCantidades, iPosicion)
                    varNumero.proEstadoNumeroDescripcion = Trim(varComando.Parameters.Item("vchDescripcionEstadoOutput").Value)
                    varNumero.proEstadoNumero = Trim(varComando.Parameters.Item("chEstadoDestinoOutput").Value)
                    varNumero.proUpdateBy = Trim(varComando.Parameters.Item("vchUserOutput").Value)
                    varNumero.proUpdateDate = Trim(varComando.Parameters.Item("vchFechaActualizacion").Value)
                    
                End If
            Else
                iError = varComando.Parameters.Item("Return").Value
                If iError <= 0 Then
                    If strError <> "" Then strError = strError & ", "
                    strError = strError & varNumero.proNumero
                End If
            End If
        End If
    Next
    
    If varCollectionEstados.Count = 0 Then
        varMensajeResultado = "No se pudo cambiar el estado de ninguno de los números."
    Else
        varMensajeResultado = "Se cambiaron de estado los siguientes números:" & Chr(13) & Chr(13)
        iContador = 0
        ' ahora armar el mensaje de resultado
        For Each varTemp In varCollectionEstados
            iContador = iContador + 1
            varMensajeResultado = varMensajeResultado & Chr(13) & _
                        varCollectionCantidades.Item(iContador) & _
                        " Números de estado " & varTemp & " a " & _
                        strEstadoDescripcion
        Next
    End If
    
    
    MetCambiarEstadoNumeros = varMensajeResultado
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Private Function FunGAgregarACollection(varObjeto As Variant, ByRef colCollection As Collection) As Integer
    Dim varexiste As Boolean
    Dim varContador As Integer
    Dim varTemp As Variant
    varContador = 0
    varexiste = False
    
    For Each varTemp In colCollection
        If varTemp = varObjeto Then
            varContador = varContador + 1
            varexiste = True
            Exit For
        End If
        varContador = varContador + 1
    Next
    
    If varexiste = False Then
        colCollection.Add (varObjeto)
        varContador = varContador + 1
    End If
    FunGAgregarACollection = varContador
    
End Function

Private Sub SubSumarEnIndex(ByRef colCollection As Collection, intIndex As Integer)

    Dim intTemp As Integer
    
    ' esto asegura de que haya n elementos en la coleccion
    If intIndex > colCollection.Count Then
        While intIndex > colCollection.Count
            Call colCollection.Add(0, CStr(intIndex))
        Wend
    End If
    
    intTemp = colCollection.Item(CStr(intIndex))
    intTemp = intTemp + 1
    colCollection.Remove (CStr(intIndex))
    Call colCollection.Add(intTemp, CStr(intIndex))
    
End Sub
