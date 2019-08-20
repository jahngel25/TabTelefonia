VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claEstratoCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proEstratoCiudadId As Long
Public proNombreEstrato As String
Public proNombreCiudad As String
Public proCiudadId As Long
Public proSeleccionado As Boolean
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
    varCommand.CommandText = "CP_InsertarTelefoniaCiudad"
    
    varCommand.Parameters.Append varCommand.CreateParameter("@iTelefoniaCiudadId", adInteger, adParamInputOutput, , 0)
    varCommand.Parameters.Append varCommand.CreateParameter("@chTipo", adChar, adParamInput, 1, "E")
    varCommand.Parameters.Append varCommand.CreateParameter("@iCiudadId", adInteger, adParamInput, , proCiudadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@vchCodigo", adVarChar, adParamInput, 20, "")
    varCommand.Parameters.Append varCommand.CreateParameter("@vchNombre", adVarChar, adParamInput, 50, proNombreEstrato)
    varCommand.Parameters.Append varCommand.CreateParameter("@vchUsuarioInserto", adVarChar, adParamInput, 50, frmAdminVoz.proLogin)
    
    varCommand.Execute
    If varCommand.Parameters.Item("@iTelefoniaCiudadId").Value <> 0 Then
        proEstratoCiudadId = varCommand.Parameters.Item("@iTelefoniaCiudadId").Value
    End If
    FunGInsertar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    On Error GoTo ErrManager
    
    Set varCommand = New ADODB.Command
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_EliminarTelefoniaCiudad"
    
    varCommand.Parameters.Append varCommand.CreateParameter("@iTelefoniaCiudadId", adInteger, adParamInput, , proEstratoCiudadId)
    
    varCommand.Execute
    
    FunGEliminar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

