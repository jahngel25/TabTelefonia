VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claParametroProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' FECHA: SEPTIEMBRE 7 /2009
'******************************************************************************

Option Explicit

Public proProductNumber As String
Public proCampo As String
Public proEtiqueta As String
Public proTipo As String
Public proTamaño As String
Public proIDInterfase As String
Public proPosicionInterfase As String
Public proMascara As String
Public proObligatorioVenta As String
Public proObligatorioAtencion As String
Public proObligatorioOT As String

Public proEditableVenta As String
Public proEditableAtencion As String
Public proEditableOT As String


Public proValidarRepetidos As String
Public proCampoPadre As String
Public ProOrden As String

Public proValorIdPadre As String

Public proValores As colValoresCampoProducto

'Conexion
Public proConexion As ADODB.Connection

Public Function MetInsertar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
           varComando = "INSERT INTO CT_ParametrosXProducto (  chProductNumber, " & _
                                       "chCampo, vchEtiqueta, " & _
                                       "chTipo, iTamaño, " & _
                                       "bIDInterfase, iPosicionInterfase, " & _
                                       "vchMascara, bObligatorioVenta, " & _
                                       "bObligatorioAtencion, bObligatorioOT, " & _
                                       "bEditableVenta, " & _
                                       "bEditableAtencion, bEditableOT, " & _
                                       "cValidarRepetidos, chCampoPadre, " & _
                                       "iOrden )   " & _
                                "VALUES ('" & Me.proProductNumber & "', '" & _
                               Me.proCampo & "', '" & Me.proEtiqueta & "', '" & _
                               Me.proTipo & "', " & Me.proTamaño & ", " & _
                               Me.proIDInterfase & ", " & Me.proPosicionInterfase & ", '" & _
                               Me.proMascara & "', " & Me.proObligatorioVenta & ", " & _
                               proObligatorioAtencion & ", " & Me.proObligatorioOT & ", " & _
                               Me.proEditableVenta & ", " & _
                               proEditableAtencion & ", " & Me.proEditableOT & ", " & _
                               Me.proValidarRepetidos & ", '" & Me.proCampoPadre & "', " & _
                               Me.ProOrden & ")"
    Me.proConexion.Execute varComando
    MetInsertar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetActualizarOrden() As Boolean
Dim varComando As String

    On Error GoTo ErrManager
    
    varComando = "UPDATE CT_ParametrosXProducto " & _
                 "SET  iOrden=" & Me.ProOrden & _
                 " WHERE chProductNumber = '" & Me.proProductNumber & "' " & _
                 " AND   chCampoPadre  = '" & Me.proCampoPadre & "' " & _
                 " AND   chCampo  = '" & Me.proCampo & "' "
    Me.proConexion.Execute varComando
    
    MetActualizarOrden = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetActualizar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
       varComando = "UPDATE CT_ParametrosXProducto " & _
                 "SET   vchEtiqueta         = '" & Me.proEtiqueta & "', " & _
                 "      chTipo              = '" & Me.proTipo & "', " & _
                 "      iTamaño             =  " & Me.proTamaño & ", " & _
                 "      bIdInterfase        =  " & Me.proIDInterfase & ", " & _
                 "      iPosicionInterfase  =  " & Me.proPosicionInterfase & ", " & _
                 "      vchMascara          = '" & Me.proMascara & "', " & _
                 "      bObligatorioVenta   =  " & Me.proObligatorioVenta & ", " & _
                 "      bObligatorioAtencion=  " & Me.proObligatorioAtencion & ", " & _
                 "      bObligatorioOT      =  " & Me.proObligatorioOT & ", " & _
                 "      beditableVenta      =  " & Me.proEditableVenta & ", " & _
                 "      beditableAtencion   =  " & Me.proEditableAtencion & ", " & _
                 "      beditableOT         =  " & Me.proEditableOT & ", " & _
                 "      cValidarRepetidos   = '" & Me.proValidarRepetidos & "' " & _
                 " WHERE chProductNumber = '" & Me.proProductNumber & "' " & _
                 " AND   chCampo         = '" & Me.proCampo & "' "
                 
    Me.proConexion.Execute varComando
    
    MetActualizar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetTieneHijos() As Boolean
Dim varResultados As ADODB.Recordset
Dim varValordatos As claValordatos
Dim varPosicion As Integer

On Error GoTo ErrorManager

        MetTieneHijos = False
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset

        varResultados.Open "SELECT     ISNULL(COUNT(*), 0) AS cuantos " _
                                        & " FROM         CT_PARAMETROSXPRODUCTO " _
                                        & " WHERE    (chCampoPadre = '" & Me.proCampoPadre & "') AND (chProductNumber = '" & Me.proProductNumber & "') AND (chCampo = '" & Me.proCampo & "')", Me.proConexion
        If Not varResultados.EOF Then
            If varResultados.Fields("cuantos") > 0 Then
                MetTieneHijos = True
            End If
        End If
    
    Exit Function
ErrorManager:
    SubGMuestraError
End Function

Public Function MetExistenRelaciones() As Boolean
Dim varResultados As ADODB.Recordset
Dim varValordatos As claValordatos
Dim varPosicion As Integer

On Error GoTo ErrorManager

        MetExistenRelaciones = False
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset

        varResultados.Open "select isnull(count(*),0) as cuantos " _
                            & " from         ct_detalledatosproducto inner join " _
                            & " ct_datosproducto on ct_detalledatosproducto.idatosproductoid = ct_datosproducto.idatosproductoid " _
                            & " group by ct_datosproducto.chproductnumber,  ct_detalledatosproducto." & Trim(Me.proCampo) _
                            & " having      (ct_datosproducto.chproductnumber = '" & Me.proProductNumber & "') and (ct_detalledatosproducto." & Trim(Me.proCampo) & " is not null) ", Me.proConexion
        If Not varResultados.EOF Then
            If varResultados.Fields("cuantos") > 0 Then
                MetExistenRelaciones = True
            End If
        End If
    
    Exit Function
ErrorManager:
    SubGMuestraError
End Function
Public Function MetEliminarValoresCampo() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "DELETE FROM CT_VALORESCAMPOxPRODUCTO " & _
                 "WHERE chProductNumber = '" & Me.proProductNumber & "' " & _
                 "AND   chCampo         = '" & Me.proCampo & "'"
    Me.proConexion.Execute varComando
    
    MetEliminarValoresCampo = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "DELETE FROM CT_ParametrosXProducto " & _
                 "WHERE chProductNumber = '" & Me.proProductNumber & "' " & _
                 "AND   chCampo         = '" & Me.proCampo & "'"
        
    Me.proConexion.Execute varComando
    
    MetEliminar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

'"--------------------------------------------------------------------------------------------
'"<summary> Método que consulta los valores de acuerdo con el id de un cliente
'"</summary>
'"<remark>MetConsultarValores(Optional pariClienteId As Long = 0)<para><version>1.0.100</version><cambio>Se adiciono para pasar como parametro el id del cliente</cambio><fecha>30/06/2009</fecha></remark>
'"</remark>
'"<param name = "pariClienteId">Id del cliente</param>
'"<returns> Booleano
'"<example> <code>if (MetConsultarValores(pariClienteId)) then </code></example>
'"<history>
'"</history>
'"----------------------------------------------------------------------------------------------
Public Function MetConsultarValores(Optional pariClienteId As Long = 0) As Boolean '1.0.100 Se pasa como parametro el id del cliente
    On Error GoTo ErrManager
    
    Set Me.proValores = Nothing
    Set Me.proValores = New colValoresCampoProducto
    Set Me.proValores.proConexion = Me.proConexion
    
    Me.proValores.proProductNumber = Trim(Me.proProductNumber)
    Me.proValores.proCampo = Trim(Me.proCampo)
    Me.proValores.proValidarRepetidos = Me.proValidarRepetidos
    
    If Trim(Me.proValorIdPadre) = "" Then
       Me.proValores.proValorIdPadre = 0
    Else
        Me.proValores.proValorIdPadre = Me.proValorIdPadre
    End If
    
    If Me.proValores.MetConsultarxCampoProducto(pariClienteId) Then '1.0.100 Se pasa como parametro el id del cliente
        MetConsultarValores = True
    Else
        MetConsultarValores = False
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetAgregarValor(parValorCampoProducto As claValoresCampoProducto) As Boolean
    On Error GoTo ErrManager
    
    Me.proValores.Add Me.proConexion, _
                      parValorCampoProducto.proValorDesc, _
                      parValorCampoProducto.proValorId, _
                      parValorCampoProducto.proValorIdPadre, _
                      parValorCampoProducto.proCampo, _
                      parValorCampoProducto.proProductNumber, 0, 0, 0
                      
    MetAgregarValor = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetMostrarVentanaEdicion(parCampoPadre1 As String, _
                                         parCampoPadre2 As String, _
                                         parCampoPadre3 As String, _
                                         parValorPadre1 As String, _
                                         parValorPadre2 As String) As Boolean
    On Error GoTo ErrManager
    
    Set frmEdicionParametros.proConexion = Me.proConexion
    
    Set frmEdicionParametros.proParametroProducto = New claParametroProducto
    Set frmEdicionParametros.proParametroProducto.proConexion = Me.proConexion
    Set frmEdicionParametros.proParametroProducto.proValores = Me.proValores
    
    frmEdicionParametros.proParametroProducto.proCampo = Me.proCampo
    frmEdicionParametros.proParametroProducto.proCampoPadre = Me.proCampoPadre
    frmEdicionParametros.proParametroProducto.proEtiqueta = Me.proEtiqueta
    frmEdicionParametros.proParametroProducto.proIDInterfase = Me.proIDInterfase
    frmEdicionParametros.proParametroProducto.proMascara = Me.proMascara
    frmEdicionParametros.proParametroProducto.proObligatorioAtencion = Me.proObligatorioAtencion
    frmEdicionParametros.proParametroProducto.proObligatorioOT = Me.proObligatorioOT
    frmEdicionParametros.proParametroProducto.proObligatorioVenta = Me.proObligatorioVenta
    frmEdicionParametros.proParametroProducto.proEditableAtencion = Me.proEditableAtencion
    frmEdicionParametros.proParametroProducto.proEditableOT = Me.proEditableOT
    frmEdicionParametros.proParametroProducto.proEditableVenta = Me.proEditableVenta
    frmEdicionParametros.proParametroProducto.ProOrden = Me.ProOrden
    frmEdicionParametros.proParametroProducto.proPosicionInterfase = Me.proPosicionInterfase
    frmEdicionParametros.proParametroProducto.proProductNumber = Me.proProductNumber
    frmEdicionParametros.proParametroProducto.proTamaño = Me.proTamaño
    frmEdicionParametros.proParametroProducto.proTipo = Me.proTipo
    frmEdicionParametros.proParametroProducto.proValidarRepetidos = Me.proValidarRepetidos
    frmEdicionParametros.proParametroProducto.proValorIdPadre = Me.proValorIdPadre
    
    frmEdicionParametros.proCampoPadre1 = parCampoPadre1
    frmEdicionParametros.proCampoPadre2 = parCampoPadre2
    frmEdicionParametros.proCampoPadre3 = parCampoPadre3
    frmEdicionParametros.proValorPadre1 = parValorPadre1
    frmEdicionParametros.proValorPadre2 = parValorPadre2
    
    frmEdicionParametros.proEdicionCliente = True
    
    frmEdicionParametros.Show (vbModal)
    
    MetMostrarVentanaEdicion = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetValidarInformacionCampo(parCampoPadre1 As String, _
                                            parCampoPadre2 As String, _
                                            parCampoPadre3 As String, _
                                            parValorPadre1 As String, _
                                            parValorPadre2 As String, _
                                            parValorPadre3 As String, _
                                            parDatosProductoId As String, _
                                            ByRef parEtiqueta As String) As Boolean

    Dim varComando As ADODB.Command
    On Error GoTo ErrManager
    
    Set varComando = New ADODB.Command
    Set varComando.ActiveConnection = Me.proConexion
    varComando.CommandType = adCmdStoredProc
    varComando.CommandText = "CP_ValidarExistenciaValorCampoProducto"
    
    If Trim(parValorPadre1) = "" Then
        parValorPadre1 = 0
    End If
    
    If Trim(parValorPadre2) = "" Then
        parValorPadre2 = 0
    End If
    
    If Trim(parValorPadre3) = "" Then
        parValorPadre3 = 0
    End If
    
    varComando.Parameters.Item("@chCampoPadre1").Value = Trim(parCampoPadre1)
    varComando.Parameters.Item("@chCampoPadre2").Value = Trim(parCampoPadre2)
    varComando.Parameters.Item("@chCampoPadre3").Value = Trim(parCampoPadre3)
    varComando.Parameters.Item("@iValorIdPadre1").Value = parValorPadre1
    varComando.Parameters.Item("@iValorIdPadre2").Value = parValorPadre2
    varComando.Parameters.Item("@iValorIdPadre3").Value = parValorPadre3
    varComando.Parameters.Item("@iDatosProductoId").Value = parDatosProductoId
    varComando.Parameters.Item("@chProductNumber").Value = Trim(Me.proProductNumber)
    varComando.Parameters.Item("@vchEtiqueta").Value = parEtiqueta
    
    varComando.Execute
    
    If varComando.Parameters.Item(0) = 1 Then
        parEtiqueta = varComando.Parameters.Item("@vchEtiqueta").Value
        MetValidarInformacionCampo = False
    Else
        MetValidarInformacionCampo = True
    End If
    Exit Function
ErrManager:
    SubGMuestraError
End Function
                                            

