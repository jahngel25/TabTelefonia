VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claDatosProductoNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proDatosProductoId As String
Public proRegionCode As String
Public proRegionName As String
Public proNumero As String
Public proFechaAsignacion As String
Public proClasificacionDescripcion As String 'Agregado por Carlos Castelblanco 2006/07/28
Public proTipoLinea As Long
Public proPublicar As String
Public proConexion As ADODB.Connection

Public proSeleccion As String

Public Function FunGInsertar()
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "  INSERT INTO ct_datosproductonumero  " & _
                                "  (iDatosProductoId, chRegionCode, vchNumero, dtFechaAsignacion) " & _
                                "  Values (" & proDatosProductoId & ",  '" & _
                                proRegionCode & "',  '" & _
                                proNumero & "',  '" & _
                                proFechaAsignacion & "') "
    
    proConexion.Execute varComando

    FunGInsertar = True
    
    Exit Function
ErrManager:
    FunGInsertar = False
    SubGMuestraError
End Function

Public Function FunGModificar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "  Update  ct datosproductonumero  " _
                       & " Set iDatosProductoId =  " & proDatosProductoId _
                       & ", chRegionCode =  '" & proRegionCode _
                       & "', vchNumero =  '" & proNumero _
                       & "', dtFechaAsignacion =  '" & proFechaAsignacion & "'" & _
                       "  Where iDatosProductoId =  " & proDatosProductoId & _
                       " AND chRegionCode =  '" & proRegionCode & "'" & _
                       " AND vchNumero =  '" & proNumero & "'"
    
    proConexion.Execute varComando
    
    FunGModificar = True
    Exit Function
ErrManager:
    FunGModificar = False
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "  Delete from  ct_datosproductonumero  " & _
     "  Where iDatosProductoId =  " & proDatosProductoId & _
     " AND chRegionCode =  '" & proRegionCode & "'" & _
     " AND vchNumero =  '" & proNumero & "'"
    
    proConexion.Execute varComando
    
    FunGEliminar = True
    Exit Function
ErrManager:
    FunGEliminar = False
    SubGMuestraError
End Function


Private Sub Class_Initialize()
    On Error GoTo ErrorManager
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub


