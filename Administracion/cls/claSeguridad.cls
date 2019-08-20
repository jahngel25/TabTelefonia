VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Propiedades de la clase
Public proUserId As String
Public proAplicacionId As Integer
Public proPrivilegios As String

Public proInsUpdDel As String
Public proCount As Integer

'Conexion
Public proConexion As ADODB.Connection

Function FunGInsertar() As Boolean
    On Error GoTo ErrManager
    
    Dim varComando As String
    varComando = "Insert Into CT_SeguridadAplicacion" & _
                 "(chUserId, iAplicacionId, vchPrivilegios) " & _
                 "Values ('" & _
                 Trim(Me.proUserId) & "', " & _
                 Trim(Me.proAplicacionId) & ", '" & _
                 Trim(Me.proPrivilegios) & "')"
                 
    proConexion.Execute varComando
    
    proInsUpdDel = "I"
    
    FunGInsertar = True
    
    Exit Function
ErrManager:
    If Err.Number = -2147217900 Then
        MsgBox "El usuario debe estar creado en Onyx.", vbInformation, App.Title
    Else
        SubGMuestraError
    End If
End Function

Function FunGModificar() As Boolean
    On Error GoTo ErrManager
    
    Dim varComando As String
    
    varComando = "Update CT_SeguridadAplicacion " & _
                 "set  vchPrivilegios = '" & Trim(Me.proPrivilegios) & _
                 "' where iAplicacionId = " & Trim(Me.proAplicacionId) & _
                 " and chUserId = '" & Trim(Me.proUserId) & "'"
                 
    proConexion.Execute varComando
    
    proInsUpdDel = "U"
    
    FunGModificar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Function FunGConsultaCount() As Boolean
    On Error GoTo ErrManager
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    Set varResultado = New ADODB.Recordset
    
    varComando = "select count(*) " & _
                 "from   CT_SeguridadAplicacion " & _
                 "where  iAplicacionId =  " & Trim(Me.proAplicacionId) & _
                 "  and  chUserId      = '" & Trim(Me.proUserId) & "'"
    
    varResultado.Open varComando, Me.proConexion
    
    If IsNull(varResultado.Fields(0)) Then
        Me.proCount = 0
    Else
        Me.proCount = varResultado.Fields(0)
    End If
    
    Set varResultado = Nothing
    FunGConsultaCount = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Function FunGEliminar() As Boolean
    On Error GoTo ErrManager
    
    Dim varComando As String
    
    varComando = "delete from CT_SeguridadAplicacion " & _
                 "where chUserId = '" & Trim(Me.proUserId) & "' " & _
                 "and   iAplicacionId = " & Trim(Me.proAplicacionId)
    
    Me.proConexion.Execute varComando
    
    FunGEliminar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


