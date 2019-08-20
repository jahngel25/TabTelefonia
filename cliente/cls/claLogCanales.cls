VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claLogCanales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************
'       DESCRIPCION: Clase con propiedades de la tabla CT_LOG_CANALES
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Desarrollo inicial de la clase
'       VERSION:              1.0.000
'       REQUERIMIENTO:        3488
'       FECHA:       2009/07/30
'*******************************************************************
Option Explicit

Public proLogCanalesId As Long
Public proIncidentId As String
Public proIncidentCategory As String
Public proTipoLinea As String
Public proSerialNumber As String
Public proCompanyId As String
Public proCanalesEnUso As Integer
Public proDiferenciaCanales As Integer
Public proCanalesCalculados As Integer
Public proUserId As String
Public proFechaNovedad As Date

Public proConexion As ADODB.Connection



