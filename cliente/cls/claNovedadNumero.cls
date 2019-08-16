VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNovedadNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proNovedadNumeroId As String
Public proRegionCode As String
Public proRegionName As String
Public proNumero As String
Public proDatosProductoId As String
Public proIncidentId As String
Public proTipoNovedadId As String
Public proFechaReserva As String
Public proFechaLiberacion As String
Public proConexion As ADODB.Connection
Public proSeleccion As String
Public proClasificacionDescripcion As String 'Agregado por Carlos Castelblanco 2006/07/28
Public proTipoLinea As Long
Public proPublicar As String
Public proTipoLineaAnterior As Long ' Tipo de línea al que se relaciona el número
Public proAsociaNovedad As String
Public procolServiciosxReserva As colserviciosxreserva
Public Function MetAgregarServiciosxReserva(proclaServiciosxReserva As claServiciosxReserva)
    On Error GoTo ErrManager
    
    If Me.procolServiciosxReserva Is Nothing Then
        Set Me.procolServiciosxReserva = New colserviciosxreserva
        Set Me.procolServiciosxReserva.proConexion = Me.proConexion
    End If
    
    Me.procolServiciosxReserva.Add proclaServiciosxReserva.proNovedadNumeroId, _
                                    proclaServiciosxReserva.proServicioSuplementarioId, _
                                    proclaServiciosxReserva.proTipoNovedadId, Me.proConexion
        MetAgregarServiciosxReserva = True

    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function FunGInsertar()
    On Error GoTo ErrManager:
    Dim varResultado As New ADODB.Recordset
    Dim varCommand As New ADODB.Command
    
    'Llamar al procedimiento de inserción en la tabla CT_NOVEDADNUMEROS
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_INSERTARNOVEDADNUMEROS"
    varCommand.Parameters.Append varCommand.CreateParameter("@P_CHREGIONCODE", adVarChar, adParamInput, 10, proRegionCode)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_VCHNUMERO", adVarChar, adParamInput, 50, proNumero)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IDATOSPRODUCTOID", adInteger, adParamInput, 0, proDatosProductoId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IINCIDENTID", adInteger, adParamInput, 0, proIncidentId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_ITIPONOVEDADID", adInteger, adParamInput, 0, proTipoNovedadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_DTFECHARESERVA", adDBTimeStamp, adParamInput, 0, IIf(proFechaReserva = "", "01/01/1900", proFechaReserva))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_DTFECHALIBERACION", adDBTimeStamp, adParamInput, 0, IIf(proFechaLiberacion = "", "01/01/1900", proFechaLiberacion))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_INOVEDADDETALLEDATOSPRODUCTOID", adInteger, adParamInput, 0, proTipoLinea)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_CHPUBLICAR", adVarChar, adParamInput, 1, IIf(proPublicar = "", "N", proPublicar))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IDETALLEDATOSPRODUCTOANTERIORID", adInteger, adParamInput, 0, proTipoLineaAnterior)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_chAsociaNovedad", adVarChar, adParamInput, 1, proAsociaNovedad)
    varResultado.Open varCommand
    If Not varResultado.EOF Then
        proNovedadNumeroId = IIf(IsNull(varResultado!INOVEDADNUMEROID), "0", varResultado!INOVEDADNUMEROID)
    Else
        proNovedadNumeroId = "0"
    End If

    Set varResultado = Nothing
    FunGInsertar = True
    Exit Function
ErrManager:
    FunGInsertar = False
    '    SubGMuestraError
    If (proConexion.Errors.Item(0).NativeError = 50001 Or proConexion.Errors.Item(0).NativeError = 50002) Then
        SubGMuestraErrorReservaDuplicado
    Else
        SubGMuestraError
    End If

End Function

Public Function FunGModificar() As Boolean
    On Error GoTo ErrManager:
    Dim varCommand As New ADODB.Command
    
    'Llamar al procedimiento de modificación de la tabla CT_NOVEDADNUMEROS
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_ACTUALIZARNOVEDADNUMEROS"
    varCommand.Parameters.Append varCommand.CreateParameter("@RETURN", adInteger, adParamReturnValue, 0, 0)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_INOVEDADNUMEROID", adInteger, adParamInput, 0, proNovedadNumeroId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_CHREGIONCODE", adVarChar, adParamInput, 10, proRegionCode)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_VCHNUMERO", adVarChar, adParamInput, 50, proNumero)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IDATOSPRODUCTOID", adInteger, adParamInput, 0, proDatosProductoId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IINCIDENTID", adInteger, adParamInput, 0, proIncidentId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_ITIPONOVEDADID", adInteger, adParamInput, 0, proTipoNovedadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_DTFECHARESERVA", adDBTimeStamp, adParamInput, 0, IIf(proFechaReserva = "", "01/01/1900", proFechaReserva))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_DTFECHALIBERACION", adDBTimeStamp, adParamInput, 0, IIf(proFechaLiberacion = "", "01/01/1900", proFechaLiberacion))
    varCommand.Parameters.Append varCommand.CreateParameter("@P_INOVEDADDETALLEDATOSPRODUCTOID", adInteger, adParamInput, 0, proTipoLinea)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_CHPUBLICAR", adVarChar, adParamInput, 1, proPublicar)
    varCommand.Execute
    If varCommand("@RETURN") <> 0 Then Err.Raise 0, , "Error en CP_ACTUALIZARDATOSPRODUCTO"
    FunGModificar = True
    Exit Function
ErrManager:
    FunGModificar = False
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "Delete From  CT_NovedadNumeros " & _
     "  Where iNovedadNumeroId =  " & proNovedadNumeroId
     
    proConexion.Execute varComando
    
    FunGEliminar = True
    Exit Function
ErrManager:
    FunGEliminar = False
    SubGMuestraError
End Function
Public Function FunGEliminarxNumero() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "Delete From  CT_NovedadNumeros " & _
     "  Where chRegionCode =  '" & Me.proRegionCode & "' AND vchNumero = '" & Me.proNumero & "'"
     
    proConexion.Execute varComando
    
    FunGEliminarxNumero = True
    Exit Function
ErrManager:
    FunGEliminarxNumero = False
    SubGMuestraError
End Function


Private Sub Class_Initialize()
    On Error GoTo ErrorManager
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
Public Function FunGEliminarServiciosxReserva() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    Set Me.procolServiciosxReserva = Nothing
    Set Me.procolServiciosxReserva = New colserviciosxreserva
    Set Me.procolServiciosxReserva.proConexion = Me.proConexion
    
    Me.procolServiciosxReserva.proNovedadNumeroId = Me.proNovedadNumeroId
    
    If Me.procolServiciosxReserva.FunGConsulta Then
        FunGEliminarServiciosxReserva = True
    Else
        FunGEliminarServiciosxReserva = False
    End If
    
    Dim i As Integer
    i = 1
    For i = 1 To Me.procolServiciosxReserva.Count
        procolServiciosxReserva.Item(i).FunGEliminar
    Next
    
    Exit Function
ErrManager:
    FunGEliminarServiciosxReserva = False
    SubGMuestraError
End Function

Public Function MetConsultarServiciosxReserva() As Boolean
    On Error GoTo ErrManager
        
    Set Me.procolServiciosxReserva = Nothing
    Set Me.procolServiciosxReserva = New colserviciosxreserva
    Set Me.procolServiciosxReserva.proConexion = Me.proConexion
    
    Me.procolServiciosxReserva.proNovedadNumeroId = Me.proNovedadNumeroId
    
    If Me.procolServiciosxReserva.FunGConsulta Then
        MetConsultarServiciosxReserva = True
    Else
        MetConsultarServiciosxReserva = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


