VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claTipoLineaEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proNovedadDetalleDatosProductoId As String
Public proUser1 As String
Public proUser15 As String
Public proContadorNumeros As Integer
Public proNovedad As Boolean

Public Sub FunGAumentarContador()
    proContadorNumeros = proContadorNumeros + 1
End Sub

Public Sub FunGDisminuirContador()
    proContadorNumeros = proContadorNumeros - 1
End Sub

