VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claClienteTelefonia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proCompanyId As String
Public proGrupoCentrex As String
Public proCallSource As String

Public proConexion As ADODB.Connection

Public Function MetInsertar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "Insert into CT_ClienteTelefonia ( iCompanyId,     " & _
                 "                                  vchGrupoCentrex," & _
                 "                                  vchCallSource ) " & _
                 "Values (" & Me.proCompanyId & ", '" & _
                              Me.proGrupoCentrex & "', '" & _
                              Me.proCallSource & "')"
                              
    Me.proConexion.Execute varComando
    
    MetInsertar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetModificar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "Update    CT_ClienteTelefonia " & _
                 "Set       vchGrupoCentrex = '" & Me.proGrupoCentrex & "', " & _
                 "          vchCallSource   = '" & Me.proCallSource & "' " & _
                 "Where     iCompanyId = " & Me.proCompanyId
                 
    Me.proConexion.Execute varComando
    
    MetModificar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
        
    varComando = "Delete from CT_ClienteTelefonia " & _
                 "Where iCompanyId = " & Me.proCompanyId
                 
    Me.proConexion.Execute varComando
    
    MetEliminar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetGuardar() As Boolean
    On Error GoTo ErrManager
    
    If MetValidarExistenciaCliente Then
        If Me.MetModificar Then
            MetGuardar = True
        Else
            MetGuardar = False
        End If
    Else
        If Me.MetInsertar Then
            MetGuardar = True
        Else
            MetGuardar = False
        End If
    End If
            
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetValidarExistenciaCliente() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varCantidad As String
    On Error GoTo ErrManager
    
    varComando = "Select    Count(*) " & _
                 "from      CT_ClienteTelefonia " & _
                 "where     iCompanyId = " & Me.proCompanyId
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If IsNull(varResultado.Fields(0).Value) Then
        varCantidad = ""
    Else
        varCantidad = Trim(varResultado.Fields(0).Value)
    End If
    
    If varCantidad = "" Or varCantidad = "0" Then
        MetValidarExistenciaCliente = False
    Else
        MetValidarExistenciaCliente = True
    End If
    
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetValidarExistenciaGrupoCentrex() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varCantidad As String
    On Error GoTo ErrManager
    
    varComando = "Select    Count(*) " & _
                 "from      CT_ClienteTelefonia " & _
                 "where     vchGrupoCentrex = '" & Me.proGrupoCentrex & "' " & _
                 "and       iCompanyId <> " & Me.proCompanyId
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If IsNull(varResultado.Fields(0).Value) Then
        varCantidad = ""
    Else
        varCantidad = Trim(varResultado.Fields(0).Value)
    End If
    
    If varCantidad = "" Or varCantidad = "0" Then
        MetValidarExistenciaGrupoCentrex = False
    Else
        MetValidarExistenciaGrupoCentrex = True
    End If
    
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetValidarExistenciaCallSource() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varCantidad As String
    On Error GoTo ErrManager
    
    varComando = "Select    Count(*) " & _
                 "from      CT_ClienteTelefonia " & _
                 "where     vchCallSource = '" & Me.proCallSource & "' " & _
                 "and       iCompanyId <> " & Me.proCompanyId
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If IsNull(varResultado.Fields(0).Value) Then
        varCantidad = ""
    Else
        varCantidad = Trim(varResultado.Fields(0).Value)
    End If
    
    If varCantidad = "" Or varCantidad = "0" Then
        MetValidarExistenciaCallSource = False
    Else
        MetValidarExistenciaCallSource = True
    End If
    
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarxCliente() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    varComando = "Select    iCompanyId,         " & _
                 "          vchGrupoCentrex,    " & _
                 "          vchCallSource       " & _
                 "From      CT_ClienteTelefonia " & _
                 "Where     iCompanyId = " & Me.proCompanyId
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iCompanyId").Value) Then
            Me.proCompanyId = ""
        Else
            Me.proCompanyId = Trim(varResultado.Fields("iCompanyId").Value)
        End If
        
        If IsNull(varResultado.Fields("vchCallSource").Value) Then
            Me.proCallSource = ""
        Else
            Me.proCallSource = Trim(varResultado.Fields("vchCallSource").Value)
        End If
        
        If IsNull(varResultado.Fields("vchGrupoCentrex").Value) Then
            Me.proGrupoCentrex = ""
        Else
            Me.proGrupoCentrex = Trim(varResultado.Fields("vchGrupoCentrex").Value)
        End If
        
        varResultado.MoveNext
    Wend
    
    MetConsultarxCliente = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
