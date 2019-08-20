VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colNormaCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection
Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, _
                    proCiudadId As Long, _
                    proCodigoNorma As String, _
                    proNombreCiudad As String, _
                    proNombreNorma As String, _
                    proNormaCiudadId As Long, _
                    Optional sKey As String) As claNormaCiudad
    Dim varNorma As claNormaCiudad
    Set varNorma = New claNormaCiudad
    If IsObject(proConexion) Then
        Set varNorma.proConexion = proConexion
    Else
        varNorma.proConexion = proConexion
    End If
    With varNorma
        .proCiudadId = proCiudadId
        .proCodigoNorma = proCodigoNorma
        .proNombreCiudad = proNombreCiudad
        .proNombreNorma = proNombreNorma
        .proNormaCiudadId = proNormaCiudadId
    End With
    If Len(sKey) = 0 Then
        mCol.Add varNorma
    Else
        mCol.Add varNorma, sKey
    End If
    Set Add = varNorma
    Set varNorma = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As claNormaCiudad
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

Public Function FunGConsulta(ParCiudadId As Long, ParRegistroActivo As Integer) As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Set varCommand = New ADODB.Command
    Set varResultado = New ADODB.Recordset
    On Error GoTo ErrManager
    
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_ConsultarTelefoniaCiudad"
    varCommand.CommandTimeout = 0
        
    varCommand.Parameters.Append varCommand.CreateParameter("@iCiudadId", adInteger, adParamInput, , ParCiudadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@chTipo", adVarChar, adParamInput, 1, "N")
    varCommand.Parameters.Append varCommand.CreateParameter("@tiRecordStatus", adTinyInt, adParamInput, , ParRegistroActivo)
    
    With varResultado
        .Open varCommand
        While Not .EOF
            Me.Add Me.proConexion, _
                     .Fields("iCiudadId"), _
                     .Fields("vchCodigo"), _
                     .Fields("vchParameterDesc"), _
                     .Fields("vchNombre"), _
                     .Fields("iTelefoniaCiudadId")
        .MoveNext
        Wend
    End With
    Set varResultado = Nothing
    FunGConsulta = True
    Exit Function
ErrManager:
    FunGConsulta = False
    SubGMuestraError
End Function





