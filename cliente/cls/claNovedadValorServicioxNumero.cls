VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNovedadValorServicioxNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public proNovedadNumeroId As String
Public proServicioSuplementario As String
Public proValor As String
Public proNumero As String
Public proRegion As String




Public proConexion As ADODB.Connection
Public proSeleccion As String

Public Function FunGInsertar()
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Debe verificar que no exista
    varResultado.Open "SELECT   Count(iNovedadNumeroId) numeroreg " & _
                  "FROM     CT_NOVEDADVALORSERVICIOXNUMERO    " & _
                "WHERE    iNovedadNumeroId  =  " & proNovedadNumeroId & " AND iServicioSuplementarioId = " & proServicioSuplementario & " AND vchNumero = '" & proNumero & "' AND chRegionCode = '" & proRegion & "' ", Me.proConexion

    If varResultado.EOF = False Then
     If CInt(varResultado("numeroreg")) = 0 Then
        'Si no existe se inserta
                varComando = "INSERT INTO CT_NOVEDADVALORSERVICIOXNUMERO " & _
             "  (iNovedadNumeroId,iServicioSuplementarioId,vchValor,vchNumero,chRegionCode) " & _
             "  Values (" & _
                                  proNovedadNumeroId & ",  " & _
                                  proServicioSuplementario & ",  '" & _
                                  proValor & "', '" & _
                                  proNumero & "', '" & proRegion & "' )  "
            
            proConexion.Execute varComando
     Else
        FunGEliminar
        
                varComando = "INSERT INTO CT_NOVEDADVALORSERVICIOXNUMERO " & _
             "  (iNovedadNumeroId,iServicioSuplementarioId,vchValor,vchNumero,chregionCode) " & _
             "  Values (" & _
                                  proNovedadNumeroId & ",  " & _
                                  proServicioSuplementario & ",  '" & _
                                  proValor & "', '" & _
                                  proNumero & "', '" & proRegion & "' )  "
            
            proConexion.Execute varComando
     
     End If
    End If
    Set varResultado = Nothing
    
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
    varComando.CommandText = "CP_InsertarTodosNovedadValorservxNumero"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@iNovedadNumeroId").Value = Me.proNovedadNumeroId
    varComando.Parameters.Item("@iServicioSuplementarioId").Value = Me.proServicioSuplementario
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
    
    varComando = "Update  CT_NOVEDADVALORSERVICIOXNUMERO " & _
                " Set vchValor =  '" & proValor & _
                "'  Where iNovedadNumeroId =  " & proNovedadNumeroId & " AND iServicioSuplementarioId = " & proServicioSuplementario
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
    
    varComando = "Delete From  CT_NOVEDADVALORSERVICIOXNUMERO " & _
     "  Where iNovedadNumeroId =  " & proNovedadNumeroId & " AND iServicioSuplementarioId = " & proServicioSuplementario & " AND vchNumero = '" & proNumero & "'"
     
    proConexion.Execute varComando
    
    FunGEliminar = True
    Exit Function
ErrManager:
    FunGEliminar = False
    SubGMuestraError
End Function
Public Function FunGEliminarTodos()
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    
    Set varComando = New ADODB.Command
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CP_EliminarTodosNovedadValorservxNumero"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@iNovedadNumeroId").Value = Me.proNovedadNumeroId
    varComando.Parameters.Item("@iServicioSuplementarioId").Value = Me.proServicioSuplementario
    varComando.Parameters.Item("@vchValor").Value = Me.proValor
    
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute

    FunGEliminarTodos = True
    Exit Function
    
ErrManager:
    FunGEliminarTodos = False
    SubGMuestraError
End Function
