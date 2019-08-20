VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claNorma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proNormaId As Long
Public proNormaCiudadId As Long
Public proNombreNorma As String
Public proCodigoNorma As String
Public proNombreCiudad As String
Public proCiudadId As Long
Public proTipoLineaId As Long
Public proUsoServicioId As Long
Public proTipoLinea As String
Public proUsoServicio As String
Public proEstratos As String
Public proRecordStatus As String
Public proConexion As ADODB.Connection
Public Function FunGInsertar() As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Set varResultado = New ADODB.Recordset
    On Error GoTo ErrManager
    
    Set varCommand = New ADODB.Command
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_InsertarNorma"
    varCommand.Parameters.Append varCommand.CreateParameter("@iCiudadId", adInteger, adParamInput, , Me.proCiudadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@iValorId", adInteger, adParamInput, , proTipoLineaId)
    varCommand.Parameters.Append varCommand.CreateParameter("@iUsoServicioId", adInteger, adParamInput, , Me.proUsoServicioId)
    varCommand.Parameters.Append varCommand.CreateParameter("@iNormaCiudadId", adInteger, adParamInput, , Me.proNormaCiudadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@vchUsuarioInserto", adVarChar, adParamInput, 50, frmAdminVoz.proLogin)
    varCommand.Parameters.Append varCommand.CreateParameter("@iNormaId", adInteger, adParamInputOutput, , 0)
    varCommand.Execute
    If varCommand.Parameters.Item("@iNormaId").Value <> 0 Then
        proNormaId = varCommand.Parameters.Item("@iNormaId").Value
    End If
    FunGInsertar = True
    Exit Function
ErrManager:
    If Err.Number = -2147217900 Then
        MsgBox "La norma a configurar ya está definida y se encuentra activa", vbInformation, "Seleccione un valor para cada campo"
    Else
        SubGMuestraError
    End If
End Function
Public Function FunGInsertarEstrato(parEstratoId As Long) As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Set varResultado = New ADODB.Recordset
    On Error GoTo ErrManager
    
    Set varCommand = New ADODB.Command
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_InsertarEstratosNorma"
    varCommand.Parameters.Append varCommand.CreateParameter("@iNormaId", adInteger, adParamInput, , Me.proNormaId)
    varCommand.Parameters.Append varCommand.CreateParameter("@iEstratoCiudadId", adInteger, adParamInput, , parEstratoId)
    varCommand.Execute
    FunGInsertarEstrato = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function FunGEliminarEstrato(parEstratoId As Long) As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Set varResultado = New ADODB.Recordset
    On Error GoTo ErrManager
    
    Set varCommand = New ADODB.Command
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_EliminarEstratosNorma"
    varCommand.Parameters.Append varCommand.CreateParameter("@iNormaId", adInteger, adParamInput, , Me.proNormaId)
    varCommand.Parameters.Append varCommand.CreateParameter("@iEstratoCiudadId", adInteger, adParamInput, , parEstratoId)
    varCommand.Execute
    FunGEliminarEstrato = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Function FunGEliminar() As Boolean
    Dim varCommand As ADODB.Command
    On Error GoTo ErrManager
    Set varCommand = New ADODB.Command
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_EliminarNorma"
    varCommand.Parameters.Append varCommand.CreateParameter("@iNormaId", adInteger, adParamInput, , proNormaId)
    varCommand.Execute
    FunGEliminar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


