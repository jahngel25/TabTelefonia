VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colEstratoCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mCol As Collection
Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, _
                    proEstratoCiudadId As Long, _
                    proNombreEstrato As String, _
                    proNombreCiudad As String, _
                    proCiudadId As Long, _
                    proSeleccionado As Boolean, _
                    Optional sKey As String) As claEstratoCiudad
                    
                    
    Dim objNewMember As claEstratoCiudad
    Set objNewMember = New claEstratoCiudad
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proEstratoCiudadId = proEstratoCiudadId
    objNewMember.proNombreEstrato = proNombreEstrato
    objNewMember.proNombreCiudad = proNombreCiudad
    objNewMember.proCiudadId = proCiudadId
    objNewMember.proSeleccionado = proSeleccionado
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As claEstratoCiudad
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
    varCommand.Parameters.Append varCommand.CreateParameter("@chTipo", adVarChar, adParamInput, 1, "E")
    varCommand.Parameters.Append varCommand.CreateParameter("@tiRecordStatus", adTinyInt, adParamInput, , ParRegistroActivo)
    
    varResultado.Open varCommand
    While Not varResultado.EOF
        Me.Add Me.proConexion, _
               varResultado.Fields("iTelefoniaCiudadId"), _
               varResultado.Fields("vchNombre"), _
               varResultado.Fields("vchParameterDesc"), _
               varResultado.Fields("iCiudadId"), _
               True
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    FunGConsulta = True
    Exit Function
ErrManager:
    FunGConsulta = False
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
    varCommand.CommandText = "CP_ConsultarServiciosPorEstrato"
    varCommand.Parameters.Append varCommand.CreateParameter("@iEstratoID", adInteger, adParamInputOutput, , Me.Item(parItem).proEstratoCiudadId)
    
    varResultado.Open varCommand
    varexiste = 0
    While Not varResultado.EOF
        varexiste = 1
        varResultado.MoveNext
    Wend
    If varexiste = 1 Then
        MsgBox "No es posible desactivar el estrato porque se encuentra relacionado a algún servicio o norma", vbInformation
    Else
        If MsgBox("¿Está seguro de desactivar el estrato " & Me.Item(parItem).proNombreEstrato & " de la ciudad " & Me.Item(parItem).proNombreCiudad & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
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

Public Function FunGConsultaNorma(PariNormaId As Long, parTipolineaid As Long, parUsoservicioid As Long, ParCiudadId As Long) As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varCommand As ADODB.Command
    Set varCommand = New ADODB.Command
    Set varResultado = New ADODB.Recordset
    On Error GoTo ErrManager
    
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_ConsultarEstratosPorNorma"
    varCommand.CommandTimeout = 0
        
    varCommand.Parameters.Append varCommand.CreateParameter("@iNormaId", adInteger, adParamInput, , PariNormaId)
    
    varResultado.Open varCommand
    While Not varResultado.EOF
        Me.Add Me.proConexion, _
                varResultado.Fields("iTelefoniaCiudadId"), _
                varResultado.Fields("vchNombre"), _
                "", _
                varResultado.Fields("iCiudadId"), _
                True
        varResultado.MoveNext
    Wend
    
    varCommand.CommandText = "CP_ConsultarEstratosLibres"
    varCommand.CommandTimeout = 0
    varCommand.Parameters.Delete (0)
    varCommand.Parameters.Append varCommand.CreateParameter("@iCiudadId", adInteger, adParamInput, , ParCiudadId)
    varCommand.Parameters.Append varCommand.CreateParameter("@iUsoServicioId", adInteger, adParamInput, , parUsoservicioid)
    varCommand.Parameters.Append varCommand.CreateParameter("@iTipoLineaId", adInteger, adParamInput, , parTipolineaid)
    varResultado.Close
    varResultado.Open varCommand
    While Not varResultado.EOF
        Me.Add Me.proConexion, _
                varResultado.Fields("iTelefoniaCiudadId"), _
                varResultado.Fields("vchNombre"), _
                "", _
                varResultado.Fields("iCiudadId"), _
                False
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    
    FunGConsultaNorma = True
    Exit Function
ErrManager:
    FunGConsultaNorma = False
    SubGMuestraError
End Function


