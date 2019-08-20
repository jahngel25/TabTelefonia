VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claTiposNovedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public proConexion As ADODB.Connection
Public proRecordStatus As String
Public proDescripcionNovedad As String
Public proTipoNovedadId As String
Public Function MetModificar() As Boolean
    On Error GoTo ErrManager
    
    If Trim(Me.proTipoNovedadId) = "" Then
        If Me.MetInsertar Then
            MetModificar = True
        Else
            MetModificar = False
        End If
    Else
        If Me.MetActualizar Then
            MetModificar = True
        Else
            MetModificar = False
        End If
    End If
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetInsertar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
     
    varComando = "Insert Into CT_TIPOSNOVEDAD     " & _
                 "(vchDescripcionNovedad, tiRecordStatus) " & _
                 "Values ('" & _
                 Trim(Me.proDescripcionNovedad) & "', " & _
                 Trim(Me.proRecordStatus) & ")"
                 
    proConexion.Execute varComando
    
    varComando = "Select    Max(iTipoNovedadID)       " & _
                 "From      CT_TIPOSNOVEDAD       "
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields(0)) Then
            Me.proTipoNovedadId = 0
        Else
            Me.proTipoNovedadId = Trim(varResultado.Fields(0))
        End If
    End If
    
    Set varResultado = Nothing
    
    MetInsertar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetActualizar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "Update    CT_TIPOSNOVEDAD   " & _
                 "Set   vchDescripcionNovedad    = '" & Trim(Me.proDescripcionNovedad) & "', " & _
                 "      tiRecordStatus  = " & Trim(Me.proRecordStatus) & " " & _
                 "where iTipoNovedadID        = " & Trim(Me.proTipoNovedadId)
    
    Me.proConexion.Execute varComando
    
    MetActualizar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    
    varComando = "Delete from CT_TIPOSNOVEDAD " & _
                 "Where iTipoNovedadID = " & Me.proTipoNovedadId
                 
    Me.proConexion.Execute varComando
    
    MetEliminar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    varComando = "Select vchDescripcionNovedad   " & _
                 "From   CT_TIPOSNOVEDAD  " & _
                 "Where  iTipoNovedadID = " & Val(Me.proTipoNovedadId) & " " & _
                 "And    tiRecordStatus = 1"
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("vchDescripcionNovedad")) Then
            Me.proDescripcionNovedad = ""
        Else
            Me.proDescripcionNovedad = Trim(varResultado.Fields("vchDescripcionNovedad"))
        End If
    End If
    
    Set varResultado = Nothing
    
    MetConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarxDescripcion() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    Me.proTipoNovedadId = 0
    
    varComando = "Select    iTipoNovedadID        " & _
                 "From      CT_TIPOSNOVEDAD   " & _
                 "Where     vchDescripcionNovedad = '" & Me.proDescripcionNovedad & "'"
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("iTipoNovedadID")) Then
            Me.proTipoNovedadId = 0
        Else
            Me.proTipoNovedadId = Trim(varResultado.Fields("iTipoNovedadID"))
        End If
    End If
    
    Set varResultado = Nothing
    
    MetConsultarxDescripcion = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


