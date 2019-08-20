VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claUsersClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proUserId As String
Public proUserName As String
Public proClasificacionId As String
Public proClasificacionDescripcion As String
Public proRecordStatus As String

Public proResultado As String
Public proMensaje As String

Public proSeleccionado As String

Public proConexion As ADODB.Connection

Public Function FunGInsertar()
Dim varComando As ADODB.Command
Dim varResultado As ADODB.Recordset
    
On Error GoTo ErrManager


Set varComando = New ADODB.Command
    
varComando.CommandType = adCmdStoredProc
varComando.ActiveConnection = Me.proConexion
varComando.CommandText = "CP_InsertaUsersClasificacion"
varComando.CommandTimeout = 0

varComando.Parameters.Item("@iClasificacionId").Value = Me.proClasificacionId
varComando.Parameters.Item("@chUserId").Value = Me.proUserId

Set varResultado = New ADODB.Recordset
Set varResultado = varComando.Execute

While Not varResultado.EOF
  Me.proResultado = varResultado.Fields("Resultado")
  Me.proMensaje = varResultado.Fields("Mensaje")
 varResultado.MoveNext
Wend

    
FunGInsertar = True
Exit Function
ErrManager:
FunGInsertar = False
SubGMuestraError
End Function
Public Function FunGEliminar() As Boolean
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

    'Eliminación Lógica coloca el recordstatus = 0
    Me.proRecordStatus = 0
    
    FunGEliminar = Me.FunGModificar
    
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function



Public Function FunGModificar() As Boolean

On Error GoTo ErrorManager

    'Ejecución de Insert
    Me.proConexion.Execute "UPDATE CT_UsersClasificacion " & _
      "SET  tiRecordStatus = '" & Me.proRecordStatus & "'" & _
       "WHERE chUserId = '" & Me.proUserId & "' AND iClasificacionId =  " & proClasificacionId
    
    If Me.proConexion.Errors.Count Then Exit Function
    FunGModificar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function


Private Sub Class_Initialize()
    On Error GoTo ErrorManager
    
        Me.proSeleccionado = "N"
        
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub



