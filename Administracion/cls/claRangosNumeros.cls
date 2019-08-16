VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claRangosNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : claRangosNumerosMódulo de clase
' Fecha  : 12/10/2004 09:20
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : clase para administrar los datos de la consulta por rango de numeros
'---------------------------------------------------------------------------------------


Option Explicit
Public proInicio As String
Public proFin As String
Public proCuantos As String
Public proConexion As ADODB.Connection
