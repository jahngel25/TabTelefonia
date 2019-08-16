VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proValorId As String
Public proValorDesc As String
Public proRecordStatus As String

'Conexion a la base
Public proConexion As ADODB.Connection

Public Function MetModificar() As Boolean
    On Error GoTo ErrManager
    
    If Trim(Me.proValorId) = "" Then
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
     
    varComando = "Insert Into CT_VALORDATOS     " & _
                 "(vchValorDesc, tiRecordStatus) " & _
                 "Values ('" & _
                 Trim(Me.proValorDesc) & "', " & _
                 Trim(Me.proRecordStatus) & ")"
                 
    proConexion.Execute varComando
    
    varComando = "Select    Max(iValorId)       " & _
                 "From      CT_ValorDatos       "
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields(0)) Then
            Me.proValorId = 0
        Else
            Me.proValorId = Trim(varResultado.Fields(0))
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
    
    varComando = "Update    CT_VALORDATOS   " & _
                 "Set   vchValorDesc    = '" & Trim(Me.proValorDesc) & "', " & _
                 "      tiRecordStatus  = " & Trim(Me.proRecordStatus) & " " & _
                 "where iValorId        = " & Trim(Me.proValorId)
    
    Me.proConexion.Execute varComando
    
    MetActualizar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    
    varComando = "Delete from CT_VALORDATOS " & _
                 "Where iValorId = " & Me.proValorId
                 
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
    
    varComando = "Select vchValorDesc   " & _
                 "From   CT_ValorDatos  " & _
                 "Where  iValorId = " & Val(Me.proValorId) & " " & _
                 "And    tiRecordStatus = 1"
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("vchValorDesc")) Then
            Me.proValorDesc = ""
        Else
            Me.proValorDesc = Trim(varResultado.Fields("vchValorDesc"))
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
    
    Me.proValorId = 0
    
    varComando = "Select    iValorId        " & _
                 "From      CT_ValorDatos   " & _
                 "Where     vchValorDesc = '" & Me.proValorDesc & "'"
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("iValorId")) Then
            Me.proValorId = 0
        Else
            Me.proValorId = Trim(varResultado.Fields("iValorId"))
        End If
    End If
    
    Set varResultado = Nothing
    
    MetConsultarxDescripcion = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
