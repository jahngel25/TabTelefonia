VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claOperacionOnyx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const OTS = 1, ATENCIONES = 2, VENTAS = 3
Public proIncidentCategory As String
Public proIncidentTypeId As String
Public proTipoNovedadId As String
Public proTipoSeccionId As String
Public proConexion As ADODB.Connection
Public Function MetInsertar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varDato As String
    If Me.proTipoSeccionId = "*" Then
        varDato = "NULL"
    Else
        varDato = "'" & Me.proTipoSeccionId & "'"
    End If
    On Error GoTo ErrManager
     
    varComando = "Insert Into CT_OPERACIONONYX  " & _
                 "(iIncidentCategory, iIncidentTypeID, " & _
                 " iTipoNovedadID,chSeccion) " & _
                 " Values (" & proIncidentCategory & "," & _
                 " " & proIncidentTypeId & "," & proTipoNovedadId & ", " & varDato & " )"
    proConexion.Execute varComando
    
    MetInsertar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetEliminar() As Boolean
    Dim varComando As String
    Dim varDato As String
    If Me.proTipoSeccionId = "*" Then
        varDato = "is NULL"
    Else
        varDato = "= '" & Trim(Me.proTipoSeccionId) & "'"
    End If
    On Error GoTo ErrManager
    varComando = "Delete from CT_OPERACIONONYX " & _
                 "Where  iIncidentTypeID    = " & Trim(Me.proIncidentTypeId) & " AND " & _
                 "      iTipoNovedadID  = " & Trim(Me.proTipoNovedadId) & " AND " & _
                 "      iIncidentCategory  = " & Trim(Me.proIncidentCategory) & " AND " & _
                 "      chSeccion  " & varDato & " "
    Me.proConexion.Execute varComando
    
    MetEliminar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Property Get proNombreSeccion() As String
   Select Case UCase(proTipoSeccionId)
        Case "T"
             proNombreSeccion = "Tipos de Linea"
        Case "P"
             proNombreSeccion = "Numeración Publica"
        Case "C"
             proNombreSeccion = "Numeración Coorporativa"
        Case "*"
             proNombreSeccion = "Todos"
   End Select
End Property
Property Get proNombreIncidente() As String
   Select Case proIncidentCategory
      Case OTS
         proNombreIncidente = "OT´s"
      Case ATENCIONES
         proNombreIncidente = "Atenciones"
      Case VENTAS
         proNombreIncidente = "Ventas"
    Case Else
        proNombreIncidente = "Categoria no definida"
   End Select
End Property
Property Get proNombreTipoIncidente() As String
Dim varResultado As ADODB.Recordset
Dim varComando As String
On Error GoTo ErrManager
If Trim(Me.proIncidentCategory) <> "" And Trim(Me.proIncidentTypeId) <> "" Then
    varComando = "SELECT  vchParameterDesc " _
    & " From ReferenceDefinition " _
    & " WHERE iReferenceId in (" _
    & "  SELECT iReferenceId " _
    & "   From Reference_Fields " _
    & "    WHERE chFieldName = 'incident.type') and tiRecordStatus = 1 and iParentId = " & proIncidentCategory & "  and  iParameterId=" & proIncidentTypeId
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.proConexion
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("vchParameterDesc")) Then
            proNombreTipoIncidente = ""
        Else
            proNombreTipoIncidente = varResultado.Fields("vchParameterDesc")
        End If
    End If
End If
Exit Property
ErrManager:
    SubGMuestraError

End Property
Property Get proNombreTipoNovedad() As String
Dim varResultado As ADODB.Recordset
Dim varComando As String
On Error GoTo ErrManager
If Trim(Me.proTipoNovedadId) <> "" Then
    varComando = "SELECT  vchDescripcionNovedad  " _
    & " From CT_TIPOSNOVEDAD" _
    & " WHERE  iTipoNovedadId = " & proTipoNovedadId & " and tiRecordStatus = 1 "
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.proConexion
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("vchDescripcionNovedad")) Then
            proNombreTipoNovedad = ""
        Else
            proNombreTipoNovedad = varResultado.Fields("vchDescripcionNovedad")
        End If
    End If
End If
Exit Property
ErrManager:
    SubGMuestraError

End Property


