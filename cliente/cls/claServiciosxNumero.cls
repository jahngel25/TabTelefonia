VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claServiciosxNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proServicioSuplementarioId As String
Public proNombreServicio As String
Public proTipoServicio As String
Public proDatosProductoId As String
Public proRegionCode As String
Public proRegionName As String
Public proNumero As String
Public proFechaActivacion As String

Public proConexion As ADODB.Connection

Public Function FunGInsertar()
Dim varComando As String
Dim varResultado As ADODB.Recordset
    
On Error GoTo ErrManager
    
    varComando = "  INSERT INTO ct_serviciosxnumero  " & _
                            "  (iServicioSuplementarioId, iDatosProductoId, chRegionCode, vchNumero, dtFechaActivacion) " & _
                            "  Values (" & proServicioSuplementarioId & ",  " & _
                            proDatosProductoId & ",  '" & _
                            proRegionCode & "',  '" & _
                            proNumero & "',  " & _
                            proFechaActivacion & ")  "
                            
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
    
varComando = "  Update  ct_serviciosxnumero " & _
                        "  Set iServicioSuplementarioId = " & proServicioSuplementarioId & ", " & _
                        " iDatosProductoId = " & proDatosProductoId & ", " & _
                        " chRegionCode = '" & proRegionCode & "', " & _
                        " vchNumero = '" & proNumero & "', " & _
                        " dtFechaActivacion = " & proFechaActivacion & _
                        "  Where iServicioSuplementarioId =  " & proServicioSuplementarioId & _
                        " And iDatosProductoId =  " & proDatosProductoId & _
                        " And chRegionCode =  '" & proRegionCode & _
                        "' And vchNumero =  '" & proNumero & _
                        "' And dtFechaActivacion =  '" & proFechaActivacion & "'"
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
    
    varComando = "  Delete from  CT_ServiciosxNumero  " & _
     "Where iServicioSuplementarioId =  " & proServicioSuplementarioId & _
     " And iDatosProductoId =  " & proDatosProductoId & _
     " And chRegionCode =  '" & proRegionCode & _
     "' And vchNumero =  '" & proNumero & "'"
     
    proConexion.Execute varComando
    
FunGEliminar = True
Exit Function
ErrManager:
FunGEliminar = False
SubGMuestraError
End Function


