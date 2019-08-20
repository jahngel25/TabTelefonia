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
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proProductNumber As String
Public proCampo As String
Public proEtiqueta As String
Public proTipo As String
Public proTamaño As Long
Public proIDInterfase As Integer
Public proPosicionInterfase As Long
Public proMascara As String
Public proObligatorioVenta As Integer
Public proObligatorioAtencion As Integer
Public proObligatorioOT As Integer
Public proValidarRepetidos As String
Public proCampoPadre As String
Public proOrden As Long


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
                                       "cValidarRepetidos, chCampoPadre, " & _
                                       "iOrden )   " & _
                                "VALUES ('" & Me.proProductNumber & "', '" & _
                               Me.proCampo & "', '" & Me.proEtiqueta & "', '" & _
                               Me.proTipo & "', " & Me.proTamaño & ", " & _
                               Me.proIDInterfase & ", " & Me.proPosicionInterfase & ", '" & _
                               Me.proMascara & "', " & Me.proObligatorioVenta & ", " & _
                               proObligatorioAtencion & ", " & Me.proObligatorioOT & ", " & _
                               Me.proValidarRepetidos & ", '" & Me.proCampoPadre & "', " & _
                               Me.proOrden & ")"
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
                 "SET  iOrden=" & Me.proOrden & _
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

Public Function MetConsultar() As Boolean
    On Error GoTo ErrManager
    
    Set Me.proValores = Nothing
    Set Me.proValores = New colValoresCampoProducto
    Set Me.proValores.proConexion = Me.proConexion
    
    Me.proValores.proProductNumber = Me.proProductNumber
    Me.proValores.proCampo = Me.proCampo
    Me.proValores.proValidarRepetidos = Me.proValidarRepetidos
    
    If Me.proValores.MetConsultarxCampoProducto Then
        MetConsultar = True
    Else
        MetConsultar = False
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
                      parValorCampoProducto.proProductNumber
                      
    MetAgregarValor = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
