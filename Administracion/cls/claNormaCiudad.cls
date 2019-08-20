VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNormaCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proNormaCiudadId As Long
Public proNombreNorma As String
Public proCodigoNorma As String
Public proNombreCiudad As String
Public proCiudadId As Long
Public proRecordStatus As String

Public proConexion As ADODB.Connection


