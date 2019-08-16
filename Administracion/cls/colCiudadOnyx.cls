VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colCiudadOnyx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public proConexion As ADODB.Connection

Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proCiudadId As Long, proNombre As String, Optional sKey As String) As claCiudadOnyx
        
    Dim objNewMember As claCiudadOnyx
    Set objNewMember = New claCiudadOnyx
        
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proCiudadId = proCiudadId
    objNewMember.proNombre = proNombre
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Function FunGConsulta() As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Set varCommand = New ADODB.Command
    Set varResultado = New ADODB.Recordset
    On Error GoTo ErrManager
    
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "sp_fac_ListarCiudades"
    varResultado.Open varCommand
    While Not varResultado.EOF
        Me.Add Me.proConexion, _
               varResultado.Fields("iCiudadId"), _
               Trim(varResultado.Fields("vchCiudad")), _
               ""
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    
    FunGConsulta = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Property Get Item(vntIndexKey As Variant) As claCiudadOnyx
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


