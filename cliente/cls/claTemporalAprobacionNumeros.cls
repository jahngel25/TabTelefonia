VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claTemporalAprobacionNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************************
'   Clase datos temporales en la aprobación de números
'   Autor: Fernando González
'   Fecha: 16-Nov-2006
'******************************************************************************************************
Option Explicit
Public temUserID As String
Public temRegionCode As String
Public temRegionname As String
Public temNumero As String
Public temEstadonumero As String
Public temDescripcionestado As String
Public temClasificacion As String
Public temUpdateby As String
Public temUseridaprobador As String
Public temFechaaprobacion As String
Public temUpdatedate As String
Public temIncidentid As String
Public temCompanyid As String
Public temCompanyname As String

Public temConexion As ADODB.Connection
