VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claEstadoOrigenDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'       DESCRIPCION: Clase con propiedades de la tabla CT_ESTADOSORIGENDESTINO
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Desarrollo inicial de la clase
'       VERSION:              1.0.000
'       REQUERIMIENTO:        5322
'       FECHA:       2009/10/10
'*******************************************************************
Option Explicit

Public proTipoEstado As String
Public proEstadoNumero As String
Public proDescripcion As String

Public proConexion As ADODB.Connection


