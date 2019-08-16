VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNumeracionCorporativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proDatosProductoId As Long
Public proMarcacion As String
Public proVirtual As String 'Agregado por Carlos Castelblanco 2006/07/26

Public proSeleccion As String

'Conexion
Public proConexion As ADODB.Connection

Public Function FunGInsertar()
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    ', chVirtual Agregado por Carlos Castelblanco 2006/07/26:
    varComando = "  insert into CT_NUMERACIONCORPORATIVA  " & _
     "  (iDatosProductoId, vchMarcacion, chVirtual) " & _
     "  Values (" & proDatosProductoId & ",  '" & proMarcacion & "' , '" & proVirtual & "')"
    
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
    
    ' " chVirtual = '" & proVirtual & " ' " & _ Agregado por Carlos Castelblanco 2006/07/26:
    varComando = "  Update  CT_NUMERACIONCORPORATIVA " & _
     "  Set iDatosProductoId =  " & proDatosProductoId & ", vchMarcacion =  '" & proMarcacion & "'," & _
     "  chVirtual = '" & proVirtual & " ' " & _
     "  Where iDatosProductoId =  " & proDatosProductoId & " AND vchMarcacion =  '" & proMarcacion & "'"
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
    
    varComando = "  Delete from  CT_NUMERACIONCORPORATIVA  " & _
     "  Where iDatosProductoId =  " & proDatosProductoId & " AND vchMarcacion =  '" & proMarcacion & "'"
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


