VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Se agrega la propiedad proObservacionesCambioEstado
'       REQUERIMIENTO:          5322
'       VERSION:       1.0.00
'       FECHA:       2009/10/10
'*******************************************************************
Option Explicit

Public proRegionCode As String
Public proRegionCodeDescripcion As String
Public proNumero As String
Public proEstadoNumero As String
Public proEstadoNumeroDescripcion As String
Public proClasificacionId As String
Public proClasificacionDescripcion As String
Public proUpdateBy As String
Public proUpdateDate As String
Public proRecordStatus As String
Public proUserIdAprobador As String
Public proFechaAprobacion As String
 '/* 1.0.000  -  Inicio */
Public proObservacionesCambioEstado As String
 '/* 1.0.000  -  Fin */
Public proIncidentId As String
Public proCompanyId As String
Public proCompanyName As String


Public proSeleccionado As String

Public proConexion As ADODB.Connection

Public Function FunGInsertar()
Dim varComando As String
Dim varResultado As ADODB.Recordset
    
On Error GoTo ErrManager
    
    varComando = "  Insert into ct_numeros  " & _
                          "  (chRegionCode, vchNumero, chEstadoNumero, chUpdateBy, dtUpdateDate, tiRecordStatus) " & _
                         "   Values ('" & proRegionCode & "',  '" & proNumero & "',  '" & proEstadoNumero & "',  '" & proUpdateBy & "',  '" & proUpdateDate & "',  " & proRecordStatus & ") "
                 
    proConexion.Execute varComando
    
FunGInsertar = True
Exit Function
ErrManager:
FunGInsertar = False
SubGMuestraError
End Function
Public Function FunGModificarEstadoAprobado() As Boolean
Dim varComando As String
    
On Error GoTo ErrManager
    
    varComando = "  Update  CT_Numeros " & _
               "Set chUserIdAprobador = '" & proUserIdAprobador & "', " & _
               " chEstadoNumero = '" & proEstadoNumero & "', " & _
               " dtFechaAprobacion = getdate() " & _
               " Where chRegionCode =  '" & proRegionCode & "' AND vchNumero =  '" & proNumero & "'"
    proConexion.Execute varComando
    
FunGModificarEstadoAprobado = True
Exit Function
ErrManager:
FunGModificarEstadoAprobado = False
SubGMuestraError
End Function
Public Function FunGModificar() As Boolean
Dim varComando As String
    
On Error GoTo ErrManager
    
    varComando = "  Update  CT_Numeros " & _
               "Set chRegionCode = '" & proRegionCode & "', " & _
                " vchNumero = '" & proNumero & "', " & _
                " chEstadoNumero = '" & proEstadoNumero & "', " & _
                " chUpdateBy = '" & proUpdateBy & "', " & _
                " dtUpdateDate = '" & proUpdateDate & "', " & _
                " tiRecordStatus = " & proRecordStatus & _
                " Where chRegionCode =  '" & proRegionCode & "' AND vchNumero =  '" & proNumero & "'"
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
    
    varComando = "  Delete from  ct_numeros  " & _
           "Where chRegionCode =  '" & proRegionCode & "' AND vchNumero =  '" & proNumero & "'"
    proConexion.Execute varComando
    
FunGEliminar = True
Exit Function
ErrManager:
FunGEliminar = False
SubGMuestraError
End Function


Private Sub Class_Initialize()
    On Error GoTo ErrorManager
    
        Me.proSeleccionado = "N"
        
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

