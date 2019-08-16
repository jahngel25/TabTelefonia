VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claDatosProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***********************************************************************
'* Descripcion:
'*
'*
'*
'*
'*
'* Parametros:
'*
'*
'*
'*
'*
'*
'*
'**************************************************************************
'**********************************************************************
' MODIFICADO POR :      CARLOS ALBERTO BARRERA
' DESCRIPCION CAMBIO:   Se pasa como parametro la propiedad del id del cliente
' VERSION: 1.0.100
' FECHA: SEPTIEMBRE 7/2009
'****************************************************************

'**********************************************************************
' MODIFICADO POR :      IVAN MAURICIO FONSECA I&T
' DESCRIPCION CAMBIO:   Con el cambio se controla que solo los tipos con la categoria de incidentes
'                       correctos puedan modificar la numeración.
'                       Se adiciona el metodo MetConsultaIncidenteTelefonia
' VERSION: 2.0.000
' FECHA: 15/Junio/2012
'****************************************************************


Option Explicit

Public proCodigoEnlace As String
Public proProductId As String
Public proProductName As String
Public proProductNumber As String
Public proDatosProductoId As String
Public proEnvioPublicoLocal As String
Public proEnvioPublicoLD As String
Public proEnvioCorpLocal As String
Public proEnvioCorpLD As String

'Registro del cliente que tiene telefonía nacional / local
Public proClienteNacionalId As String
Public proNombreClienteNacional As String
Public proClienteLocalId As String
Public proNombreClienteLocal As String

Public proComentarios As String
Public proRecordStatus As String
Public proCodigoRegion As String
Public proTipoTelefonia As String

' Incidente de venta
Public proiVentaid As String

'Ciudad destino
Public proCiudadId As Long
Public proNombreCiudad As String

'Estrato
Public proiEstratoid As String
Public proDescripcionEstrato As String

'Uso
Public proUsoServicioId As Long
Public proDescripcionUso As String

Public proDatosProductoIncident As colDatosProductoIncident
Public proDetalleDatosProducto As colDetalleDatosProducto
Public proNovedadDetalleDatosProducto As colNovedadDetalleDatosProducto

Public proParametrosProducto As colParametroProducto
Public proProductMaster As colProductMaster

Public proDatosProductoNumero As colDatosProductoNumero
Public proNovedadNumero As colNovedadNumero

Public proServiciosxNumero As colServiciosxNumero

Public proNumeracionCorporativa As colNumeracionCorporativa
Public proNovedadNumeracionCorporativa As colNovedadNumeracionCorporativa

'Variable para verificar la existencia del incidente en la coleccion
Public proIncidentId As String
Public proVirtual As String 'Agregado por Carlos Castelblanco 2006/07/26

Public proConexion As ADODB.Connection

Public Function MetConsultar() As Boolean
    On Error GoTo ErrManager:
    Dim varResultado As New ADODB.Recordset
    Dim varCommand As New ADODB.Command
    
    'Llamar al procedimiento de consulta, pasando como parámetro la llave de CT_DATOSPRODUCTO
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_CONSULTARDATOSPRODUCTO"
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IDATOSPRODUCTOID", adInteger, adParamInput, 0, proDatosProductoId)
    varResultado.Open varCommand
    If Not varResultado.EOF Then
        proProductId = IIf(IsNull(varResultado!IPRODUCTID), "", varResultado!IPRODUCTID)
        proEnvioPublicoLocal = IIf(IsNull(varResultado!BENVIOPUBLICOLOCAL), "", varResultado!BENVIOPUBLICOLOCAL)
        proEnvioPublicoLD = IIf(IsNull(varResultado!BENVIOPUBLICOLD), "", varResultado!BENVIOPUBLICOLD)
        proEnvioCorpLocal = IIf(IsNull(varResultado!BENVIOCORPLOCAL), "", varResultado!BENVIOCORPLOCAL)
        proEnvioCorpLD = IIf(IsNull(varResultado!BENVIOCORPLD), "", varResultado!BENVIOCORPLD)
        proComentarios = IIf(IsNull(varResultado!VCHCOMENTARIOS), "", varResultado!VCHCOMENTARIOS)
        proRecordStatus = IIf(IsNull(varResultado!TIRECORDSTATUS), "", varResultado!TIRECORDSTATUS)
        proiVentaid = IIf(IsNull(varResultado!IVENTAID), "", varResultado!IVENTAID)
        proClienteNacionalId = IIf(IsNull(varResultado!icompanyid), "", varResultado!icompanyid)
        proNombreClienteNacional = IIf(IsNull(varResultado!vchcompanyname), "", varResultado!vchcompanyname)
        proClienteLocalId = IIf(IsNull(varResultado!ICOMPANYIDLOCAL), "", varResultado!ICOMPANYIDLOCAL)
        proNombreClienteLocal = IIf(IsNull(varResultado!VCHCOMPANYNAMELOCAL), "", varResultado!VCHCOMPANYNAMELOCAL)
        proCiudadId = IIf(IsNull(varResultado!VCHDATAVALUECIUDAD), 0, varResultado!VCHDATAVALUECIUDAD)
        proNombreCiudad = IIf(IsNull(varResultado!VCHPARAMETERDESCCIUDAD), "", varResultado!VCHPARAMETERDESCCIUDAD)
        proiEstratoid = varResultado!IESTRATOCIUDADID
        proDescripcionEstrato = IIf(IsNull(varResultado!VCHNOMBREESTRATO), "", varResultado!VCHNOMBREESTRATO)
        proUsoServicioId = varResultado!IUSOSERVICIOID
        proDescripcionUso = IIf(IsNull(varResultado!VCHDESCRIPCIONUSO), "", varResultado!VCHDESCRIPCIONUSO)
        proCodigoRegion = IIf(IsNull(varResultado!VCHCODIGOREGION), "", varResultado!VCHCODIGOREGION)
        proTipoTelefonia = IIf(IsNull(varResultado!TIPOTELEFONIA), "", varResultado!TIPOTELEFONIA)
    End If
    
    Set varResultado = Nothing
    MetConsultar = Me.MetConsultaProducto
    Exit Function
ErrManager:
    SubGMuestraError
    MetConsultar = False
End Function

Public Function MetConsultarIncidentes() As Boolean
    On Error GoTo ErrManager
    
    Set Me.proDatosProductoIncident = Nothing
    Set Me.proDatosProductoIncident = New colDatosProductoIncident
    Set Me.proDatosProductoIncident.proConexion = Me.proConexion
    Me.proDatosProductoIncident.proDatosProductoId = Me.proDatosProductoId
    
    If Me.proDatosProductoIncident.MetConsultarxDatosProducto Then
        MetConsultarIncidentes = True
    Else
        MetConsultarIncidentes = False
    End If
    
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarDetalles() As Boolean
    On Error GoTo ErrManager
    
    Set Me.proDetalleDatosProducto = Nothing
    Set Me.proDetalleDatosProducto = New colDetalleDatosProducto
    Set Me.proDetalleDatosProducto.proConexion = Me.proConexion
    Me.proDetalleDatosProducto.proDatosProductoId = Me.proDatosProductoId
    
    If Me.proDetalleDatosProducto.MetConsultarxDatosProducto Then
        MetConsultarDetalles = True
    Else
        MetConsultarDetalles = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetVerificarExistenciaIncidente() As Boolean
    Dim varContador As Integer
    Dim varEncontro As Boolean
    On Error GoTo ErrManager
    
    varEncontro = False
    For varContador = 1 To Me.proDatosProductoIncident.Count
        If Me.proIncidentId = Me.proDatosProductoIncident.Item(varContador).proIncidentId Then
            varEncontro = True
            Exit For
        End If
    Next varContador
    
    MetVerificarExistenciaIncidente = varEncontro
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetAgregarIncidente(parDatosProductoIncident As EDCAdminVoz.claDatosProductoIncident) As Boolean
    On Error GoTo ErrManager
    
    If Me.proDatosProductoIncident Is Nothing Then
        Set Me.proDatosProductoIncident = New colDatosProductoIncident
        Set Me.proDatosProductoIncident.proConexion = Me.proConexion
    End If
    
    Me.proDatosProductoIncident.Add Me.proConexion, _
                                    parDatosProductoIncident.proDescripcion, _
                                    parDatosProductoIncident.proTipo, _
                                    parDatosProductoIncident.proCategoria, _
                                    parDatosProductoIncident.proFechaModificacion, _
                                    parDatosProductoIncident.proIncidentId, _
                                    parDatosProductoIncident.proDatosProductoId
    
    
    MetAgregarIncidente = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetAgregarDetalle(parDetalleDatosProducto As claDetalleDatosProducto) As Boolean
    On Error GoTo ErrManager
    
    If Me.proDetalleDatosProducto Is Nothing Then
        Set Me.proDetalleDatosProducto = New colDetalleDatosProducto
        Set Me.proDetalleDatosProducto.proConexion = Me.proConexion
    End If
    
    Me.proDetalleDatosProducto.Add Me.proConexion, _
                                    parDetalleDatosProducto.proRecordStatus, _
                                    parDetalleDatosProducto.proUser40, parDetalleDatosProducto.proUser39, _
                                    parDetalleDatosProducto.proUser38, parDetalleDatosProducto.proUser37, _
                                    parDetalleDatosProducto.proUser36, parDetalleDatosProducto.proUser35, _
                                    parDetalleDatosProducto.proUser34, parDetalleDatosProducto.proUser33, _
                                    parDetalleDatosProducto.proUser32, parDetalleDatosProducto.proUser31, _
                                    parDetalleDatosProducto.proUser30, parDetalleDatosProducto.proUser29, _
                                    parDetalleDatosProducto.proUser28, parDetalleDatosProducto.proUser27, _
                                    parDetalleDatosProducto.proUser26, parDetalleDatosProducto.proUser25, _
                                    parDetalleDatosProducto.proUser24, parDetalleDatosProducto.proUser23, _
                                    parDetalleDatosProducto.proUser22, parDetalleDatosProducto.proUser21, _
                                    parDetalleDatosProducto.proUser20, parDetalleDatosProducto.proUser19, _
                                    parDetalleDatosProducto.proUser18, parDetalleDatosProducto.proUser17, _
                                    parDetalleDatosProducto.proUser16, parDetalleDatosProducto.proUser15, _
                                    parDetalleDatosProducto.proUser14, parDetalleDatosProducto.proUser13, _
                                    parDetalleDatosProducto.proUser12, parDetalleDatosProducto.proUser11, _
                                    parDetalleDatosProducto.proUser10, parDetalleDatosProducto.proUser9, _
                                    parDetalleDatosProducto.proUser8, parDetalleDatosProducto.proUser7, _
                                    parDetalleDatosProducto.proUser6, parDetalleDatosProducto.proUser5, _
                                    parDetalleDatosProducto.proUser4, parDetalleDatosProducto.proUser3, _
                                    parDetalleDatosProducto.proUser2, parDetalleDatosProducto.proUser1, _
                                    parDetalleDatosProducto.proStatusId, parDetalleDatosProducto.proDetalleDatosProductoId, _
                                    parDetalleDatosProducto.proDatosProductoId
    
    
    MetAgregarDetalle = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetAgregarNovedadDetalle(parNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto) As Boolean
    On Error GoTo ErrManager
    
    If Me.proNovedadDetalleDatosProducto Is Nothing Then
        Set Me.proNovedadDetalleDatosProducto = New colNovedadDetalleDatosProducto
        Set Me.proNovedadDetalleDatosProducto.proConexion = Me.proConexion
    End If
    
    Me.proNovedadDetalleDatosProducto.Add Me.proConexion, _
                                    parNovedadDetalleDatosProducto.proRecordStatus, parNovedadDetalleDatosProducto.proUser40, parNovedadDetalleDatosProducto.proUser39, _
                                    parNovedadDetalleDatosProducto.proUser38, parNovedadDetalleDatosProducto.proUser37, parNovedadDetalleDatosProducto.proUser36, _
                                    parNovedadDetalleDatosProducto.proUser35, parNovedadDetalleDatosProducto.proUser34, parNovedadDetalleDatosProducto.proUser33, _
                                    parNovedadDetalleDatosProducto.proUser32, parNovedadDetalleDatosProducto.proUser31, parNovedadDetalleDatosProducto.proUser30, _
                                    parNovedadDetalleDatosProducto.proUser29, parNovedadDetalleDatosProducto.proUser28, parNovedadDetalleDatosProducto.proUser27, _
                                    parNovedadDetalleDatosProducto.proUser26, parNovedadDetalleDatosProducto.proUser25, parNovedadDetalleDatosProducto.proUser24, _
                                    parNovedadDetalleDatosProducto.proUser23, parNovedadDetalleDatosProducto.proUser22, parNovedadDetalleDatosProducto.proUser21, _
                                    parNovedadDetalleDatosProducto.proUser20, parNovedadDetalleDatosProducto.proUser19, parNovedadDetalleDatosProducto.proUser18, _
                                    parNovedadDetalleDatosProducto.proUser17, parNovedadDetalleDatosProducto.proUser16, parNovedadDetalleDatosProducto.proUser15, _
                                    parNovedadDetalleDatosProducto.proUser14, parNovedadDetalleDatosProducto.proUser13, parNovedadDetalleDatosProducto.proUser12, _
                                    parNovedadDetalleDatosProducto.proUser11, parNovedadDetalleDatosProducto.proUser10, parNovedadDetalleDatosProducto.proUser9, _
                                    parNovedadDetalleDatosProducto.proUser8, parNovedadDetalleDatosProducto.proUser7, parNovedadDetalleDatosProducto.proUser6, _
                                    parNovedadDetalleDatosProducto.proUser5, parNovedadDetalleDatosProducto.proUser4, parNovedadDetalleDatosProducto.proUser3, _
                                    parNovedadDetalleDatosProducto.proUser2, parNovedadDetalleDatosProducto.proUser1, parNovedadDetalleDatosProducto.proStatusId, _
                                    parNovedadDetalleDatosProducto.proDetalleDatosProductoId, parNovedadDetalleDatosProducto.proIncidentId, _
                                    parNovedadDetalleDatosProducto.proDatosProductoId, parNovedadDetalleDatosProducto.proTipoNovedadId, _
                                    parNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId, _
                                    parNovedadDetalleDatosProducto.proContadorNumeros
    
    
    MetAgregarNovedadDetalle = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetAgregarNovedadNumeroPublico(parNovedadNumero As claNovedadNumero) As Boolean
    On Error GoTo ErrManager
    
    If Me.proNovedadNumero Is Nothing Then
        Set Me.proNovedadNumero = New colNovedadNumero
        Set Me.proNovedadNumero.proConexion = Me.proConexion
    End If
    
    Me.proNovedadNumero.Add Me.proConexion, _
                                    parNovedadNumero.proFechaLiberacion, _
                                    parNovedadNumero.proFechaReserva, _
                                    parNovedadNumero.proTipoNovedadId, _
                                    parNovedadNumero.proIncidentId, _
                                    parNovedadNumero.proNumero, _
                                    parNovedadNumero.proDatosProductoId, _
                                    parNovedadNumero.proRegionName, _
                                    parNovedadNumero.proRegionCode, _
                                    parNovedadNumero.proNovedadNumeroId, _
                                    parNovedadNumero.proClasificacionDescripcion, _
                                    parNovedadNumero.proTipoLinea, _
                                    parNovedadNumero.proTipoLineaAnterior, _
                                    parNovedadNumero.proPublicar, _
                                    parNovedadNumero.proAsociaNovedad
    
    MetAgregarNovedadNumeroPublico = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetAgregarNovedadNumeracionCorporativa(parNovedadNumeracionCorporativa As claNovedadNumeracionCorporativa) As Boolean
    On Error GoTo ErrManager
    
    If Me.proNovedadNumeracionCorporativa Is Nothing Then
        Set Me.proNovedadNumeracionCorporativa = New colNovedadNumeracionCorporativa
        Set Me.proNovedadNumeracionCorporativa.proConexion = Me.proConexion
    End If
    
    Me.proNovedadNumeracionCorporativa.Add Me.proConexion, _
                                    parNovedadNumeracionCorporativa.proMarcacion, _
                                    parNovedadNumeracionCorporativa.proTipoNovedadId, _
                                    parNovedadNumeracionCorporativa.proIncidentId, _
                                    parNovedadNumeracionCorporativa.proDatosProductoId, _
                                    parNovedadNumeracionCorporativa.proVirtual
                                    'parNovedadNumeracionCorporativa.proVirtual Agregado por Carlos Castelblanco 2006/07/26
                                    
    
    MetAgregarNovedadNumeracionCorporativa = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetInsertar() As Boolean
    On Error GoTo ErrManager:
    Dim varCommand As New ADODB.Command
    
    'Llamar al procedimiento de inserción en la tabla CT_DATOSPRODUCTO
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_INSERTARDATOSPRODUCTO"
    varCommand.Parameters.Append varCommand.CreateParameter("@RETURN", adInteger, adParamReturnValue, 0, 0)
    varCommand.Parameters.Append varCommand.CreateParameter("@IDATOSPRODUCTOID", adInteger, adParamOutput, 0, 0)
    varCommand.Parameters.Append varCommand.CreateParameter("@IPRODUCTID", adVarChar, adParamInput, 10, Val(proProductId))
    varCommand.Parameters.Append varCommand.CreateParameter("@CHPRODUCTNUMBER", adVarChar, adParamInput, 20, proProductNumber)
    varCommand.Parameters.Append varCommand.CreateParameter("@BENVIOPUBLICOLOCAL", adBoolean, adParamInput, 0, Val(proEnvioPublicoLocal))
    varCommand.Parameters.Append varCommand.CreateParameter("@BENVIOPUBLICOLD", adBoolean, adParamInput, 0, Val(proEnvioPublicoLD))
    varCommand.Parameters.Append varCommand.CreateParameter("@BENVIOCORPLOCAL", adBoolean, adParamInput, 0, Val(proEnvioCorpLocal))
    varCommand.Parameters.Append varCommand.CreateParameter("@BENVIOCORPLD", adBoolean, adParamInput, 0, Val(proEnvioCorpLD))
    varCommand.Parameters.Append varCommand.CreateParameter("@ICOMPANYID", adInteger, adParamInput, 0, IIf(Val(proClienteNacionalId) = 0, Null, proClienteNacionalId))
    varCommand.Parameters.Append varCommand.CreateParameter("@ICOMPANYLOCALID", adInteger, adParamInput, 0, IIf(Val(proClienteLocalId) = 0, Null, proClienteLocalId))
    varCommand.Parameters.Append varCommand.CreateParameter("@VCHCOMENTARIOS", adVarChar, adParamInput, 900, proComentarios)
    varCommand.Parameters.Append varCommand.CreateParameter("@TIRECORDSTATUS", adUnsignedTinyInt, adParamInput, 0, proRecordStatus)
    varCommand.Parameters.Append varCommand.CreateParameter("@IVENTAID", adInteger, adParamInput, 0, Val(proiVentaid))
    varCommand.Parameters.Append varCommand.CreateParameter("@IESTRATOCIUDADID", adInteger, adParamInput, 0, Val(proiEstratoid))
    varCommand.Parameters.Append varCommand.CreateParameter("@IUSOSERVICIOID", adInteger, adParamInput, 0, proUsoServicioId)
    varCommand.Execute
    If varCommand.Parameters("@RETURN") <> 0 Then Err.Raise 0, , "Error en CP_INSERTARDATOSPRODUCTO"
    proDatosProductoId = IIf(IsNull(varCommand.Parameters("@IDATOSPRODUCTOID")), "0", varCommand.Parameters("@IDATOSPRODUCTOID"))

    MetInsertar = True
    Exit Function
ErrManager:
    SubGMuestraError
    MetInsertar = False
End Function

Public Function MetActualizar() As Boolean
    On Error GoTo ErrManager:
    Dim varCommand As New ADODB.Command
    
    If Me.proEnvioPublicoLocal = "True" Or Me.proEnvioPublicoLocal = "1" Then
        Me.proEnvioPublicoLocal = 1
    Else
        Me.proEnvioPublicoLocal = 0
    End If
    
    If Me.proEnvioPublicoLD = "True" Or Me.proEnvioPublicoLD = "1" Then
        Me.proEnvioPublicoLD = 1
    Else
        Me.proEnvioPublicoLD = 0
    End If
    
    If Me.proEnvioCorpLocal = "True" Or Me.proEnvioCorpLocal = "1" Then
        Me.proEnvioCorpLocal = 1
    Else
        Me.proEnvioCorpLocal = 0
    End If
    
    If Me.proEnvioCorpLD = "True" Or Me.proEnvioCorpLD = "1" Then
        Me.proEnvioCorpLD = 1
    Else
        Me.proEnvioCorpLD = 0
    End If

    'Llamar al procedimiento para actualizar la tabla CT_DATOSPRODUCTO
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_ACTUALIZARDATOSPRODUCTO"
    varCommand.Parameters.Append varCommand.CreateParameter("@RETURN", adInteger, adParamReturnValue, 0, 0)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IDATOSPRODUCTOID", adInteger, adParamInput, 0, proDatosProductoId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IPRODUCTID", adVarChar, adParamInput, 10, Val(proProductId))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_CHPRODUCTNUMBER", adVarChar, adParamInput, 20, proProductNumber)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_BENVIOPUBLICOLOCAL", adBoolean, adParamInput, 0, Val(proEnvioPublicoLocal))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_BENVIOPUBLICOLD", adBoolean, adParamInput, 0, Val(proEnvioPublicoLD))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_BENVIOCORPLOCAL", adBoolean, adParamInput, 0, Val(proEnvioCorpLocal))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_BENVIOCORPLD", adBoolean, adParamInput, 0, Val(proEnvioCorpLD))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_ICOMPANYID", adInteger, adParamInput, 0, IIf(Val(proClienteNacionalId) = 0, Null, proClienteNacionalId))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_ICOMPANYLOCALID", adInteger, adParamInput, 0, IIf(Val(proClienteLocalId) = 0, Null, proClienteLocalId))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_VCHCOMENTARIOS", adVarChar, adParamInput, 900, proComentarios)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_TIRECORDSTATUS", adUnsignedTinyInt, adParamInput, 0, proRecordStatus)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IVENTAID", adInteger, adParamInput, 0, Val(proiVentaid))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IESTRATOCIUDADID", adInteger, adParamInput, 0, IIf(proiEstratoid = "", 0, proiEstratoid))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IUSOSERVICIOID", adInteger, adParamInput, 0, proUsoServicioId)
    varCommand.Execute
    If varCommand.Parameters("@RETURN") <> 0 Then Err.Raise 0, , "Error en CP_ACTUALIZARDATOSPRODUCTO"
    MetActualizar = True
    Exit Function
ErrManager:
    SubGMuestraError
    MetActualizar = False
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "DELETE FROM CT_DatosProducto  " & _
                 "WHERE iDatosProductoId = " & Me.proDatosProductoId
    
    Me.proConexion.Execute varComando
        
    MetEliminar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetGuardar() As Boolean
    On Error GoTo ErrManager
    
    If Val(Trim(Me.proDatosProductoId)) = 0 Then
        If Me.MetInsertar Then
            MetGuardar = True
            Me.MetConsultar
        Else
            MetGuardar = False
        End If
    Else
        If Me.MetActualizar Then
            MetGuardar = True
        Else
            MetGuardar = False
        End If
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetGuardarColeccionIncidentes() As Boolean
    Dim varContador As Integer
    Dim varContadorInterno As Integer
    Dim varEncontro As Boolean
    Dim varDatosProductoIncident As colDatosProductoIncident
    On Error GoTo ErrManager
    
    Set varDatosProductoIncident = New colDatosProductoIncident
    
    varDatosProductoIncident.proDatosProductoId = Me.proDatosProductoId
    Set varDatosProductoIncident.proConexion = Me.proConexion
    If Not varDatosProductoIncident.MetConsultarxDatosProducto Then
        MsgBox "Error al recuperar la informacion de los incidentes.", vbCritical, App.Title
    End If
    
    varContador = 1
    While varContador <= Me.proDatosProductoIncident.Count
        Me.proDatosProductoIncident.Item(varContador).proDatosProductoId = Me.proDatosProductoId
        
        varEncontro = False
        If Not varDatosProductoIncident Is Nothing Then
            varContadorInterno = 1
            While varContadorInterno <= varDatosProductoIncident.Count And Not varEncontro
                If Val(Me.proDatosProductoIncident.Item(varContador).proDatosProductoId) = Val(varDatosProductoIncident.Item(varContadorInterno).proDatosProductoId) And _
                    Val(Me.proDatosProductoIncident.Item(varContador).proIncidentId) = Val(varDatosProductoIncident.Item(varContadorInterno).proIncidentId) Then
                    varEncontro = True
                End If
                If Not varEncontro Then
                    varContadorInterno = varContadorInterno + 1
                End If
            Wend
        End If
            
        If Not varEncontro Then
            'Me.proDatosProductoIncident.Item(varContador).MetGuardar ("A")
        'Else
            Me.proDatosProductoIncident.Item(varContador).MetGuardar ("I")
        End If
        varContador = varContador + 1
    Wend
    
    MetGuardarColeccionIncidentes = True
    
    Exit Function
ErrManager:
    MetGuardarColeccionIncidentes = False
    SubGMuestraError
End Function


Public Function MetConsultaProducto() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager:
    
    Set varResultado = New ADODB.Recordset
    
    'Buscar la informacion del codigo de enlace y el nombre del producto
    varComando = "SELECT    CP.chProductNumber,     " & _
                 "          PM.vchDescription,      " & _
                 "          CP.vchSerialNumber      " & _
                 "FROM      Customer_Product    CP, " & _
                 "          Product_Master      PM  " & _
                 "WHERE     CP.iProductId       = " & Me.proProductId & " " & _
                 "AND       PM.chProductNumber  = CP.chProductNumber    "

    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("vchSerialNumber")) Then
            Me.proCodigoEnlace = ""
        Else
            Me.proCodigoEnlace = varResultado.Fields("vchSerialNumber")
        End If
    
        If IsNull(varResultado.Fields("vchDescription")) Then
            Me.proProductName = ""
        Else
            Me.proProductName = varResultado.Fields("vchDescription")
        End If
    
        If IsNull(varResultado.Fields("chProductNumber")) Then
            Me.proProductNumber = ""
        Else
            Me.proProductNumber = varResultado.Fields("chProductNumber")
        End If
    End If
    MetConsultaProducto = True
    Set varResultado = Nothing
    Exit Function
    
ErrManager:
    SubGMuestraError
End Function


Public Function MetConsultarProductMaster() As Boolean
    On Error GoTo ErrManager
        
        Set Me.proProductMaster = New colProductMaster
        Set Me.proProductMaster.proConexion = Me.proConexion
        
        Me.proProductMaster.proIncidentId = Trim(Me.proDatosProductoIncident.Item(1).proIncidentId)
        
        If Me.proProductMaster.MetConsultarxIncidente Then
            Me.proProductNumber = Me.proProductMaster.Item(1).proProductNumber
            Me.proProductName = Me.proProductMaster.Item(1).proDescription
            MetConsultarProductMaster = True
        Else
            MetConsultarProductMaster = False
        End If
    Exit Function
ErrManager:
    SubGMuestraError
End Function

'"--------------------------------------------------------------------------------------------
'"<summary> Método que consulta los parametros x producto
'"</summary>
'"<remark>MetConsultarParametrosProducto(Optional pariClienteId As Long = 0)<para><version>1.0.000</version><cambio>Se adiciono para pasar como parametro el id del cliente</cambio><fecha>30/06/2009</fecha></remark>
'"</remark>
'"<param name = "pariClienteId">Id del cliente</param>
'"<returns> Booleano
'"<example> <code>if (MetConsultarParametrosProducto(pariClienteId)) then </code></example>
'"<history>
'"</history>
'"----------------------------------------------------------------------------------------------


Public Function MetConsultarParametrosProducto(Optional pariClienteId As Long = 0) As Boolean '* 1.0.100 Inicio Se Modifica el metodo, pasandole el parametro del id del cliente opcional
    On Error GoTo ErrManager
    
    Set Me.proParametrosProducto = Nothing
    Set Me.proParametrosProducto = New colParametroProducto
    Set Me.proParametrosProducto.proConexion = Me.proConexion
    Me.proParametrosProducto.proProductNumber = Me.proProductNumber
    
    If Me.proParametrosProducto.metConsultarxProducto(pariClienteId) Then '*1.0.100 Se pasa el parametro del id del cliente
        MetConsultarParametrosProducto = True
    Else
        MetConsultarParametrosProducto = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarDatosProductoNumero() As Boolean
    On Error GoTo ErrManager
    
    Set Me.proDatosProductoNumero = Nothing
    Set Me.proDatosProductoNumero = New colDatosProductoNumero
    Set Me.proDatosProductoNumero.proConexion = Me.proConexion
    
    Me.proDatosProductoNumero.proDatosProductoId = Me.proDatosProductoId
    
    If Me.proDatosProductoNumero.MetConsultar(Me.proDetalleDatosProducto) Then
        MetConsultarDatosProductoNumero = True
    Else
        MetConsultarDatosProductoNumero = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetConsultarNovedadDetalleDatosProducto() As Boolean
    
    On Error GoTo ErrManager
    
    Set Me.proNovedadDetalleDatosProducto = Nothing
    Set Me.proNovedadDetalleDatosProducto = New colNovedadDetalleDatosProducto
    Set Me.proNovedadDetalleDatosProducto.proConexion = Me.proConexion
    
    Me.proNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProductoId
    Me.proNovedadDetalleDatosProducto.proIncidentId = Me.proIncidentId
    
    If Me.proNovedadDetalleDatosProducto.MetConsultarxIncidente(Me.proDetalleDatosProducto) Then
        MetConsultarNovedadDetalleDatosProducto = True
    Else
        MetConsultarNovedadDetalleDatosProducto = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarNovedadNumeros() As Boolean
    On Error GoTo ErrManager
    
    Set Me.proNovedadNumero = Nothing
    Set Me.proNovedadNumero = New colNovedadNumero
    Set Me.proNovedadNumero.proConexion = Me.proConexion
    
    Me.proNovedadNumero.proDatosProductoId = Me.proDatosProductoId
    Me.proNovedadNumero.proIncidentId = Me.proIncidentId
    
    If Me.proNovedadNumero.MetConsultar(proDetalleDatosProducto) Then
        MetConsultarNovedadNumeros = True
    Else
        MetConsultarNovedadNumeros = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarNovedadNumeracionCorporativa() As Boolean
    On Error GoTo ErrManager
    
    Set Me.proNovedadNumeracionCorporativa = Nothing
    Set Me.proNovedadNumeracionCorporativa = New colNovedadNumeracionCorporativa
    Set Me.proNovedadNumeracionCorporativa.proConexion = Me.proConexion
    
    Me.proNovedadNumeracionCorporativa.proDatosProductoId = Me.proDatosProductoId
    Me.proNovedadNumeracionCorporativa.proIncidentId = Me.proIncidentId
    
    If Me.proNovedadNumeracionCorporativa.MetConsultar Then
        MetConsultarNovedadNumeracionCorporativa = True
    Else
        MetConsultarNovedadNumeracionCorporativa = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarServiciosxNumero() As Boolean
    On Error GoTo ErrManager
    
    Set Me.proServiciosxNumero = Nothing
    Set Me.proServiciosxNumero = New colServiciosxNumero
    Set Me.proServiciosxNumero.proConexion = Me.proConexion
    
    Me.proServiciosxNumero.proDatosProductoId = Me.proDatosProductoId
    
    If Me.proServiciosxNumero.MetConsultar Then
        MetConsultarServiciosxNumero = True
    Else
        MetConsultarServiciosxNumero = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarNumeracionCorporativa() As Boolean
    On Error GoTo ErrManager
        
    Set Me.proNumeracionCorporativa = Nothing
    Set Me.proNumeracionCorporativa = New colNumeracionCorporativa
    Set Me.proNumeracionCorporativa.proConexion = Me.proConexion
    
    Me.proNumeracionCorporativa.proDatosProductoId = Me.proDatosProductoId
    
    If Me.proNumeracionCorporativa.FunGConsulta Then
        MetConsultarNumeracionCorporativa = True
    Else
        MetConsultarNumeracionCorporativa = False
    End If
        
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetActualizarDetalle(parDetalleDatosProducto As claDetalleDatosProducto) As Boolean
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    For varContador = 1 To Me.proDetalleDatosProducto.Count
        If Me.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = _
            parDetalleDatosProducto.proDetalleDatosProductoId Then
                                                
            Me.proDetalleDatosProducto.Item(varContador).proRecordStatus = parDetalleDatosProducto.proRecordStatus
            Me.proDetalleDatosProducto.Item(varContador).proUser40 = parDetalleDatosProducto.proUser40
            Me.proDetalleDatosProducto.Item(varContador).proUser39 = parDetalleDatosProducto.proUser39
            Me.proDetalleDatosProducto.Item(varContador).proUser38 = parDetalleDatosProducto.proUser38
            Me.proDetalleDatosProducto.Item(varContador).proUser37 = parDetalleDatosProducto.proUser37
            Me.proDetalleDatosProducto.Item(varContador).proUser36 = parDetalleDatosProducto.proUser36
            Me.proDetalleDatosProducto.Item(varContador).proUser35 = parDetalleDatosProducto.proUser35
            Me.proDetalleDatosProducto.Item(varContador).proUser34 = parDetalleDatosProducto.proUser34
            Me.proDetalleDatosProducto.Item(varContador).proUser33 = parDetalleDatosProducto.proUser33
            Me.proDetalleDatosProducto.Item(varContador).proUser32 = parDetalleDatosProducto.proUser32
            Me.proDetalleDatosProducto.Item(varContador).proUser31 = parDetalleDatosProducto.proUser31
            Me.proDetalleDatosProducto.Item(varContador).proUser30 = parDetalleDatosProducto.proUser30
            Me.proDetalleDatosProducto.Item(varContador).proUser29 = parDetalleDatosProducto.proUser29
            Me.proDetalleDatosProducto.Item(varContador).proUser28 = parDetalleDatosProducto.proUser28
            Me.proDetalleDatosProducto.Item(varContador).proUser27 = parDetalleDatosProducto.proUser27
            Me.proDetalleDatosProducto.Item(varContador).proUser26 = parDetalleDatosProducto.proUser26
            Me.proDetalleDatosProducto.Item(varContador).proUser25 = parDetalleDatosProducto.proUser25
            Me.proDetalleDatosProducto.Item(varContador).proUser24 = parDetalleDatosProducto.proUser24
            Me.proDetalleDatosProducto.Item(varContador).proUser23 = parDetalleDatosProducto.proUser23
            Me.proDetalleDatosProducto.Item(varContador).proUser22 = parDetalleDatosProducto.proUser22
            Me.proDetalleDatosProducto.Item(varContador).proUser21 = parDetalleDatosProducto.proUser21
            Me.proDetalleDatosProducto.Item(varContador).proUser20 = parDetalleDatosProducto.proUser20
            Me.proDetalleDatosProducto.Item(varContador).proUser19 = parDetalleDatosProducto.proUser19
            Me.proDetalleDatosProducto.Item(varContador).proUser18 = parDetalleDatosProducto.proUser18
            Me.proDetalleDatosProducto.Item(varContador).proUser17 = parDetalleDatosProducto.proUser17
            Me.proDetalleDatosProducto.Item(varContador).proUser16 = parDetalleDatosProducto.proUser16
            Me.proDetalleDatosProducto.Item(varContador).proUser15 = parDetalleDatosProducto.proUser15
            Me.proDetalleDatosProducto.Item(varContador).proUser14 = parDetalleDatosProducto.proUser14
            Me.proDetalleDatosProducto.Item(varContador).proUser13 = parDetalleDatosProducto.proUser13
            Me.proDetalleDatosProducto.Item(varContador).proUser12 = parDetalleDatosProducto.proUser12
            Me.proDetalleDatosProducto.Item(varContador).proUser11 = parDetalleDatosProducto.proUser11
            Me.proDetalleDatosProducto.Item(varContador).proUser10 = parDetalleDatosProducto.proUser10
            Me.proDetalleDatosProducto.Item(varContador).proUser9 = parDetalleDatosProducto.proUser9
            Me.proDetalleDatosProducto.Item(varContador).proUser8 = parDetalleDatosProducto.proUser8
            Me.proDetalleDatosProducto.Item(varContador).proUser7 = parDetalleDatosProducto.proUser7
            Me.proDetalleDatosProducto.Item(varContador).proUser6 = parDetalleDatosProducto.proUser6
            Me.proDetalleDatosProducto.Item(varContador).proUser5 = parDetalleDatosProducto.proUser5
            Me.proDetalleDatosProducto.Item(varContador).proUser4 = parDetalleDatosProducto.proUser4
            Me.proDetalleDatosProducto.Item(varContador).proUser3 = parDetalleDatosProducto.proUser3
            Me.proDetalleDatosProducto.Item(varContador).proUser2 = parDetalleDatosProducto.proUser2
            Me.proDetalleDatosProducto.Item(varContador).proUser1 = parDetalleDatosProducto.proUser1
            Me.proDetalleDatosProducto.Item(varContador).proStatusId = parDetalleDatosProducto.proStatusId
            Me.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = parDetalleDatosProducto.proDetalleDatosProductoId
            Me.proDetalleDatosProducto.Item(varContador).proDatosProductoId = parDetalleDatosProducto.proDatosProductoId
            
            Exit For
        End If
    Next varContador
    MetActualizarDetalle = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetActualizarNovedadDetalle(parNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto) As Boolean
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    For varContador = 1 To Me.proNovedadDetalleDatosProducto.Count
        If Me.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId = _
            parNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId Then
                                                
            Me.proNovedadDetalleDatosProducto.Item(varContador).proRecordStatus = parNovedadDetalleDatosProducto.proRecordStatus
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser40 = parNovedadDetalleDatosProducto.proUser40
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser39 = parNovedadDetalleDatosProducto.proUser39
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser38 = parNovedadDetalleDatosProducto.proUser38
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser37 = parNovedadDetalleDatosProducto.proUser37
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser36 = parNovedadDetalleDatosProducto.proUser36
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser35 = parNovedadDetalleDatosProducto.proUser35
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser34 = parNovedadDetalleDatosProducto.proUser34
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser33 = parNovedadDetalleDatosProducto.proUser33
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser32 = parNovedadDetalleDatosProducto.proUser32
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser31 = parNovedadDetalleDatosProducto.proUser31
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser30 = parNovedadDetalleDatosProducto.proUser30
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser29 = parNovedadDetalleDatosProducto.proUser29
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser28 = parNovedadDetalleDatosProducto.proUser28
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser27 = parNovedadDetalleDatosProducto.proUser27
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser26 = parNovedadDetalleDatosProducto.proUser26
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser25 = parNovedadDetalleDatosProducto.proUser25
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser24 = parNovedadDetalleDatosProducto.proUser24
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser23 = parNovedadDetalleDatosProducto.proUser23
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser22 = parNovedadDetalleDatosProducto.proUser22
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser21 = parNovedadDetalleDatosProducto.proUser21
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser20 = parNovedadDetalleDatosProducto.proUser20
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser19 = parNovedadDetalleDatosProducto.proUser19
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser18 = parNovedadDetalleDatosProducto.proUser18
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser17 = parNovedadDetalleDatosProducto.proUser17
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser16 = parNovedadDetalleDatosProducto.proUser16
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser15 = parNovedadDetalleDatosProducto.proUser15
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser14 = parNovedadDetalleDatosProducto.proUser14
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser13 = parNovedadDetalleDatosProducto.proUser13
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser12 = parNovedadDetalleDatosProducto.proUser12
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser11 = parNovedadDetalleDatosProducto.proUser11
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser10 = parNovedadDetalleDatosProducto.proUser10
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser9 = parNovedadDetalleDatosProducto.proUser9
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser8 = parNovedadDetalleDatosProducto.proUser8
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser7 = parNovedadDetalleDatosProducto.proUser7
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser6 = parNovedadDetalleDatosProducto.proUser6
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser5 = parNovedadDetalleDatosProducto.proUser5
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser4 = parNovedadDetalleDatosProducto.proUser4
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser3 = parNovedadDetalleDatosProducto.proUser3
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser2 = parNovedadDetalleDatosProducto.proUser2
            Me.proNovedadDetalleDatosProducto.Item(varContador).proUser1 = parNovedadDetalleDatosProducto.proUser1
            Me.proNovedadDetalleDatosProducto.Item(varContador).proStatusId = parNovedadDetalleDatosProducto.proStatusId
            Me.proNovedadDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = parNovedadDetalleDatosProducto.proDetalleDatosProductoId
            Me.proNovedadDetalleDatosProducto.Item(varContador).proIncidentId = parNovedadDetalleDatosProducto.proIncidentId
            Me.proNovedadDetalleDatosProducto.Item(varContador).proDatosProductoId = parNovedadDetalleDatosProducto.proDatosProductoId
            Me.proNovedadDetalleDatosProducto.Item(varContador).proTipoNovedadId = parNovedadDetalleDatosProducto.proTipoNovedadId
            Me.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId = parNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId
            
            Exit For
        End If
    Next varContador
    MetActualizarNovedadDetalle = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminarDetalle(parRegistro As Integer) As Boolean
    On Error GoTo ErrManager
    
    Me.proDetalleDatosProducto.Item(parRegistro).proRecordStatus = 0
    
    If Me.proDetalleDatosProducto.Item(parRegistro).MetGuardar Then
        MetEliminarDetalle = True
    Else
        MetEliminarDetalle = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminarNovedadDetalle(parRegistro As Integer) As Boolean
    On Error GoTo ErrManager
    
    Me.proNovedadDetalleDatosProducto.Item(parRegistro).proRecordStatus = 0
    
    If Me.proNovedadDetalleDatosProducto.Item(parRegistro).MetGuardar Then
        MetEliminarNovedadDetalle = True
    Else
        MetEliminarNovedadDetalle = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetDevolverEnlace(varVenta As String) As String
   Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager:
    
    Set varResultado = New ADODB.Recordset
    
    'Buscar la informacion del codigo de enlace y el nombre del producto
    varComando = "SELECT    isnull(CP.vchSerialNumber,'') as Enlace," & _
                 "          I.vchUser7   " & _
                 "          FROM customer_product CP " & _
                 "          INNER JOIN Incident I " & _
                 "          ON CP.vchUser1 = I.iIncidentId " & _
                 " where I.vchUser7 = '" & varVenta & "'"
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        MetDevolverEnlace = Trim(varResultado("Enlace"))
    Else
        MetDevolverEnlace = ""
    End If
 Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetDevolverVenta(varEnlace As String) As String
   Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager:
    
    Set varResultado = New ADODB.Recordset
    
    'Buscar la informacion del codigo de enlace y el nombre del producto
    varComando = "SELECT    CP.vchSerialNumber," & _
                 "          isnull(I.vchUser7,'') as Venta " & _
                 "          FROM customer_product CP " & _
                 "          INNER JOIN Incident I " & _
                 "          ON CP.vchUser1 = I.iIncidentId " & _
                 " where cp.vchSerialNumber = '" & varEnlace & "'"
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        MetDevolverVenta = Trim(varResultado("Venta"))
    End If
 Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarCiudadDestino() As Boolean
    On Error GoTo ErrManager:
    Dim varComando As String
    Dim varResultado As New ADODB.Recordset
    MetConsultarCiudadDestino = True
    If Me.proiVentaid = "" Then Exit Function
    'Buscar la informacion de la ciudad
    varComando = "SELECT ED.VCHDATAVALUE VCHDATAVALUECIUDAD, RD.VCHPARAMETERDESC VCHPARAMETERDESCCIUDAD, R.VCHCODIGO VCHCODIGOREGION " & _
        "FROM EXPANSIONDATA ED " & _
        "LEFT OUTER JOIN REFERENCEDEFINITION RD ON  ED.VCHDATAVALUE = RD.IPARAMETERID AND RD.IREFERENCEID = 888 " & _
        "LEFT OUTER JOIN CT_TELEFONIACIUDAD R ON  ED.VCHDATAVALUE = R.ICIUDADID AND R.CHTIPO = 'R' " & _
        "WHERE   ED.CHFIELDNAME = 'incident.sales.expand60' AND ED.ISYSTEMID = " & Me.proiVentaid
    varResultado.Open varComando, Me.proConexion
    If varResultado.EOF Then Exit Function
    Me.proCiudadId = IIf(IsNull(varResultado!VCHDATAVALUECIUDAD), "0", varResultado!VCHDATAVALUECIUDAD)
    Me.proCodigoRegion = IIf(IsNull(varResultado!VCHCODIGOREGION), "", varResultado!VCHCODIGOREGION)
    Me.proNombreCiudad = IIf(IsNull(varResultado!VCHPARAMETERDESCCIUDAD), "", varResultado!VCHPARAMETERDESCCIUDAD)
    Exit Function
ErrManager:
    MetConsultarCiudadDestino = False
    SubGMuestraError
End Function

' Inicio 2.0.000 Se valida si el tipo del incidente puede alterar la numeracion
Public Function MetConsultaIncidenteTelefonia(TipoIncidente As String, ClasificacionIncidente As String) As String
    
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager:
    
    Set varResultado = New ADODB.Recordset
    
    varComando = "SELECT    iIncidentTypeId     " & _
                 "FROM      CT_CierreIncidentesTelefonia IT  " & _
                 "WHERE     IT.iIncidentTypeId  =  " & TipoIncidente

    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        MetConsultaIncidenteTelefonia = Trim(varResultado("iIncidentTypeId"))
    End If
    
 Exit Function
ErrManager:
    SubGMuestraError
End Function
' Fin 2.0.000


