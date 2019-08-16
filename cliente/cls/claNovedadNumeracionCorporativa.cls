VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNovedadNumeracionCorporativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proDatosProductoId As Long
Public proIncidentId As Long
Public proTipoNovedadId As Long
Public proMarcacion As String
Public proVirtual As String 'Agregado por Carlos Castelblanco 2006/07/26

Public proSeleccion As String

'Conexion
Public proConexion As ADODB.Connection

Public Function FunGInsertar()
    Dim varComando As String
        
    On Error GoTo ErrManager
    
    ', proVirtual Agregado por Carlos Castelblanco 2006/07/26
    ' "  chVirtual = '" & proVirtual & "')" Agregado por Carlos Castelblanco 2006/07/26
    varComando = "  INSERT INTO ct_novedadnumeracioncorporativa  " & _
                            "  (iDatosProductoId, iIncidentId, iTipoNovedadId, vchMarcacion, chVirtual) " & _
                            "  Values (" & proDatosProductoId & ",  " & proIncidentId & ",  " & proTipoNovedadId & ",  '" & proMarcacion & "', " & _
                            "   '" & proVirtual & "')"
                            
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
    
    '", chVirtual =  '" & proVirtual & "'" & _ Agregado por Carlos Castelblanco 2006/07/26
    varComando = "  Update  ct_novedadnumeracioncorporativa " & _
                            "  Set iDatosProductoId =  " & proDatosProductoId & _
                            ", iIncidentId =  " & proIncidentId & _
                            ", iTipoNovedadId =  " & proTipoNovedadId & _
                            ", vchMarcacion =  '" & proMarcacion & "'" & _
                            ", chVirtual =  '" & proVirtual & "'" & _
                            "  Where iDatosProductoId =  " & proDatosProductoId & _
                            " AND iIncidentId =  " & proIncidentId & _
                            " AND iTipoNovedadId =  " & proTipoNovedadId
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
    
    varComando = "  Delete from  ct_novedadnumeracioncorporativa  " & _
                            "  Where iDatosProductoId =  " & proDatosProductoId & _
                            " AND iIncidentId =  " & proIncidentId & _
                            " AND iTipoNovedadId =  " & proTipoNovedadId & _
                            " AND vchMarcacion =  " & "'" & proMarcacion & "'"
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

