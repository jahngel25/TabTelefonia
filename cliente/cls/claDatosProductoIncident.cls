VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claDatosProductoIncident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proDatosProductoId As String
Public proIncidentId As String
Public proFechaModificacion As String

'Información adicional del incidente
Public proDescripcion As String
Public proTipo As String
Public proCategoria As String

Public proConexion As ADODB.Connection

Public Function MetInsertar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "INSERT INTO CT_DatosProductoIncident  ( iDatosProductoId,     " & _
                 "                                        iIncidentId,          " & _
                 "                                        dtFechaModificacion ) " & _
                 "VALUES (" & Me.proDatosProductoId & ", " & _
                              Me.proIncidentId & ", '" & _
                              Me.proFechaModificacion & "')"
                              
    Me.proConexion.Execute varComando
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetActualizar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    
    varComando = "UPDATE    CT_DatosProductoIncident    " & _
                 "SET       iIncidentId         =  " & Me.proIncidentId & ", " & _
                 "          dtFechaModificacion = '" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "' " & _
                 "WHERE     iDatosProductoId = " & Me.proDatosProductoId
                 
    Me.proConexion.Execute varComando
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "DELETE FROM CT_DatosProductoIncident  " & _
                 "WHERE iDatosProductoId = " & Me.proDatosProductoId & " " & _
                 "AND   iIncidentId      = " & Me.proIncidentId
                 
    Me.proConexion.Execute varComando
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetGuardar(parInsUpd As String) As Boolean
    On Error GoTo ErrManager
    
    If parInsUpd = "I" Then
        If Me.MetInsertar Then
            MetGuardar = True
        Else
            MetGuardar = True
        End If
    Else
        If Me.MetActualizar Then
            MetGuardar = True
        Else
            MetGuardar = False
        End If
    End If
    
    MetGuardar = True
    Exit Function
ErrManager:
    MetGuardar = False
End Function

