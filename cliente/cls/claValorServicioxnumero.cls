VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claValorServicioxnumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public proServicioSuplementario As String
Public proRegionCode As String
Public proNumero As String
Public proValor As String
Public proTipoServicio As String


Public proConexion As ADODB.Connection
Public proSeleccion As String

Public Function FunGInsertar()
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    varComando = "INSERT INTO CT_VALORSERVICIOXNUMERO " & _
     "  (iServicioSuplementarioId,chRegionCode,vchNumero,vchValor) " & _
     "  Values (" & _
                          proServicioSuplementario & ",  '" & _
                          proRegionCode & "', '" & _
                          proNumero & "', '" & _
                          proValor & "')  "
    
    proConexion.Execute varComando
    
    FunGInsertar = True
    Exit Function
    
ErrManager:
    FunGInsertar = False
    SubGMuestraError
End Function
Public Function FunGInsertarTodos()
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    
    Set varComando = New ADODB.Command
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_InsertarTodosValorservxNumero"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@iServicioSuplementarioId").Value = Me.proServicioSuplementario
    varComando.Parameters.Item("@chRegionCode").Value = Me.proRegionCode
    varComando.Parameters.Item("@vchNumero").Value = Me.proNumero
    varComando.Parameters.Item("@vchValor").Value = Me.proValor
    
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute

    FunGInsertarTodos = True
    Exit Function
    
ErrManager:
    FunGInsertarTodos = False
    SubGMuestraError
End Function
Public Function FunGModificar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "Update  CT_VALORSERVICIOXNUMERO " & _
                " Set vchValor =  '" & proValor & _
                "'  Where iServicioSuplementario = " & proServicioSuplementario & " AND chRegionCode = '" & Me.proRegionCode & "' AND vchNumero = '" & Me.proNumero & "' "
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
    
    varComando = "Delete From  CT_VALORSERVICIOXNUMERO " & _
     "  Where  iServicioSuplementarioId = " & proServicioSuplementario & " AND chRegionCode = '" & Me.proRegionCode & "' AND vchNumero = '" & Me.proNumero & "' "
     
    proConexion.Execute varComando
    
    FunGEliminar = True
    Exit Function
ErrManager:
    FunGEliminar = False
    SubGMuestraError
End Function

Public Function FunGEliminarTodos() As Boolean
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
        
    Set varComando = New ADODB.Command
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_EliminarTodosValorservxNumero"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@iServicioSuplementarioId").Value = Me.proServicioSuplementario
    varComando.Parameters.Item("@chRegionCode").Value = Me.proRegionCode
    varComando.Parameters.Item("@vchNumero").Value = Me.proNumero
    varComando.Parameters.Item("@vchValor").Value = Me.proValor
    
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute
    
    FunGEliminarTodos = True
    Exit Function
ErrManager:
    FunGEliminarTodos = False
    SubGMuestraError
End Function

