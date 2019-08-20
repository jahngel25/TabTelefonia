VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claConsultaNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proConexion As ADODB.Connection
Public proNumeros As colNumero
Public proValoresCampoProducto As colValoresCampoProducto
Public proTipoLineaEdicion As colTipoLineaEdicion
Public proNo As String
Public proTipoLineaBasico As Boolean
Public proCodigoTipoLineaBasica As String
Public proIndiceTipoLineaEdicion As Long
Public proCodCiudad As String
Public proSeleccionInstalado As Boolean
Public Sub MetMostrarVentanaConsulta()
    On Error GoTo ErrManager
    
    Set frmConsultaNumeros.proConexion = Me.proConexion
    Set frmConsultaNumeros.proNumeros = Me.proNumeros
    Set frmConsultaNumeros.proValoresCampoProducto = Me.proValoresCampoProducto
    Set frmConsultaNumeros.proTipoLineaEdicion = Me.proTipoLineaEdicion
    frmConsultaNumeros.proNo = Me.proNo
    frmConsultaNumeros.proCodCiudad = Me.proCodCiudad
    frmConsultaNumeros.Show (vbModal)
    Set Me.proNumeros = frmConsultaNumeros.proNumeros
    Me.proTipoLineaBasico = frmConsultaNumeros.proTipoLineaBasico
    Me.proCodigoTipoLineaBasica = frmConsultaNumeros.proCodigoTipoLineaBasica
    Me.proIndiceTipoLineaEdicion = frmConsultaNumeros.proIndiceTipoLineaEdicion
    If frmConsultaNumeros.proIndiceInstalado = 0 Then
        proSeleccionInstalado = False
    Else
        If frmConsultaNumeros.proIndiceSeleccionado >= frmConsultaNumeros.proIndiceInstalado Then
            proSeleccionInstalado = True
        Else
            proSeleccionInstalado = False
        End If
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


