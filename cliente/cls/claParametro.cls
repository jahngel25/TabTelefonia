VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public proValor As String
Public proAcronimo As String

'Conexion
Public proConexion As ADODB.Connection

Public Function FunGConsultar() As Boolean
    If proValor = "" Then
        Dim varResultado As ADODB.Recordset
        Dim varCommand As ADODB.Command
        Set varResultado = New ADODB.Recordset
        On Error GoTo ErrManager
        Set varCommand = New ADODB.Command
        varCommand.CommandType = adCmdStoredProc
        varCommand.ActiveConnection = Me.proConexion
        varCommand.CommandText = "CP_Gral_BuscarParametro"
        varCommand.Parameters.Append varCommand.CreateParameter("@vchAcronimo", adVarChar, adParamInput, 50, proAcronimo)
        varCommand.Parameters.Append varCommand.CreateParameter("@iAplicacionId", adInteger, adParamInput, , 20)
        varResultado.Open varCommand
        While Not varResultado.EOF
            proValor = varResultado("vchValor")
            varResultado.MoveNext
        Wend
    End If
    FunGConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


