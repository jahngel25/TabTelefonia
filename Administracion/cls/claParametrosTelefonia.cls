VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claParametrosTelefonia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proParametro As String
Public proValor As String

Public proConexion As ADODB.Connection

Public Function MetConsultarParametro() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    varComando = "Select    vchParametro,           " & _
                 "          vchValor                " & _
                 "From      CT_ParametrosTelefonia  " & _
                 "Where     vchParametro = '" & Me.proParametro & "'"
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("vchParametro")) Then
            Me.proParametro = ""
        Else
            Me.proParametro = varResultado.Fields("vchParametro")
        End If
    
        If IsNull(varResultado.Fields("vchValor")) Then
            Me.proValor = ""
        Else
            Me.proValor = varResultado.Fields("vchValor")
        End If
    End If
    
    MetConsultarParametro = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


