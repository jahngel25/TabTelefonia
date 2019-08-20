VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColNorma"
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
                    proNormaId As Long, _
                    proUsoServicioId As Long, _
                    proUsoServicio As String, _
                    proTipoLineaId As Long, _
                    proTipoLinea As String, _
                    proNormaCiudadId As Long, _
                    proEstratos As String, _
                    Optional sKey As String) As claNorma
    Dim varNorma As claNorma
    Set varNorma = New claNorma
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
        .proNormaId = proNormaId
        .proUsoServicioId = proUsoServicioId
        .proUsoServicio = proUsoServicio
        .proTipoLinea = proTipoLinea
        .proTipoLineaId = proTipoLineaId
        .proNormaCiudadId = proNormaCiudadId
        .proEstratos = proEstratos
    End With
    If Len(sKey) = 0 Then
        mCol.Add varNorma
    Else
        mCol.Add varNorma, sKey
    End If
    Set Add = varNorma
    Set varNorma = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As claNorma
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
    Dim varResultadoE As ADODB.Recordset
    Dim varCommandE As ADODB.Command
    Set varCommand = New ADODB.Command
    Set varResultado = New ADODB.Recordset
    Dim varestrato As String
    On Error GoTo ErrManager
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_ConsultarNorma"
    varCommand.CommandTimeout = 0
    varCommand.Parameters.Append varCommand.CreateParameter("@iCiudadId", adInteger, adParamInput, , ParCiudadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@tiRecordStatus", adTinyInt, adParamInput, , ParRegistroActivo)
    With varResultado
        .Open varCommand
        While Not .EOF
            varestrato = ""
            Set varCommandE = New ADODB.Command
            Set varResultadoE = New ADODB.Recordset
            varCommandE.CommandType = adCmdStoredProc
            varCommandE.ActiveConnection = Me.proConexion
            varCommandE.CommandText = "CP_ConsultarEstratosporNorma"
            varCommandE.CommandTimeout = 0
            varCommandE.Parameters.Append varCommandE.CreateParameter("@iCiudadId", adInteger, adParamInput, , .Fields("iNormaId"))
            varResultadoE.Open varCommandE
            While Not varResultadoE.EOF
                varestrato = varestrato & IIf(varestrato = "", "", ",") & varResultadoE.Fields("vchNombre")
                varResultadoE.MoveNext
            Wend
            Me.Add Me.proConexion, _
                    .Fields("iCiudadId"), _
                    .Fields("vchCodigo"), _
                    .Fields("vchNombreCiudad"), _
                    .Fields("vchNombre"), _
                    .Fields("iNormaId"), _
                    .Fields("iUsoServicioId"), _
                    .Fields("vchUsoServicio"), _
                    .Fields("iTipoLineaId"), _
                    .Fields("vchTipoLinea"), _
                    .Fields("iNormaCiudadId"), _
                    varestrato
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
Public Function FunGConsultaporParametros(ParCiudadId As Long, parUsoservicioid As Long, parTipolineaid As Long, parEstratoId As Long) As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Dim varResultadoE As ADODB.Recordset
    Dim varCommandE As ADODB.Command
    Set varCommand = New ADODB.Command
    Set varResultado = New ADODB.Recordset
    Dim varestrato As String
    On Error GoTo ErrManager
    With varCommand
        .CommandType = adCmdStoredProc
        .ActiveConnection = Me.proConexion
        .CommandText = "CP_ConsultarNormaporParametros"
        .CommandTimeout = 0
        .Parameters.Append varCommand.CreateParameter("@iCiudadId", adInteger, adParamInput, , ParCiudadId)
        .Parameters.Append varCommand.CreateParameter("@iUsoServicioId", adInteger, adParamInput, , parUsoservicioid)
        .Parameters.Append varCommand.CreateParameter("@iTipoLineaId", adInteger, adParamInput, , parTipolineaid)
        .Parameters.Append varCommand.CreateParameter("@iEstratoId", adInteger, adParamInput, , parEstratoId)
        .Parameters.Append varCommand.CreateParameter("@iNormaId", adInteger, adParamOutput, , 0)
        .Parameters.Append varCommand.CreateParameter("@vchCodigo", adVarChar, adParamOutput, 10, "")
        .Parameters.Append varCommand.CreateParameter("@vchNombre", adVarChar, adParamOutput, 50, "")
        .Execute
        Dim varNorma As Long
        If IsNull(.Parameters("@iNormaId")) Then
            varNorma = 0
        Else
            varNorma = .Parameters("@iNormaId")
        End If
        If varNorma <> 0 Then
                Me.Add Me.proConexion, _
                    0, _
                    .Parameters("@vchCodigo"), _
                    0, _
                    .Parameters("@vchNombre"), _
                    .Parameters("@iNormaId"), _
                    0, _
                    "", _
                    0, _
                    "", _
                    0, _
                    ""
        End If
    End With
    FunGConsultaporParametros = True
    Exit Function
ErrManager:
    FunGConsultaporParametros = False
    SubGMuestraError

End Function
Public Function FunGEliminar(parItem As Integer) As Boolean
 Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Dim objNewMember As claEstratoCiudad
    Dim varexiste As Integer
    Set objNewMember = New claEstratoCiudad
    On Error GoTo ErrManager
    Set varCommand = New ADODB.Command
    Set varResultado = New ADODB.Recordset
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_ConsultarServiciosporNorma"
    varCommand.Parameters.Append varCommand.CreateParameter("@iNorma", adInteger, adParamInput, , Me.Item(parItem).proNormaId)
    varResultado.Open varCommand
    While Not varResultado.EOF
        varexiste = 1
        varResultado.MoveNext
    Wend
    If varexiste = 1 Then
        MsgBox "No es posible desactivar la norma porque se encuentra relacionada en algún servicio", vbInformation
        FunGEliminar = True
    Else
        If MsgBox("¿Esta seguro de desactivar la configuración de la norma?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            FunGEliminar = True
            Exit Function
        End If
        Set Me.Item(parItem).proConexion = Me.proConexion
        FunGEliminar = Me.Item(parItem).FunGEliminar
        Me.Remove (parItem)
    End If
    Exit Function
ErrManager:
    SubGMuestraError
End Function



