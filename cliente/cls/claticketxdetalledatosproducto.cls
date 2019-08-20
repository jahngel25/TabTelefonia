VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claticketxdetalledatosproducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proiIncidentId As Long
Public proiDatosProductoId As Long
Public proiDetalleDatosProductoId As Long


'Conexion
Public proConexion As ADODB.Connection

Public Function FunGInsertar()
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    varComando = "  INSERT INTO ct_ticketxdetalledatosproducto  " & _
     "  (iIncidentId, iDatosProductoId, iDetalleDatosProductoId  ) " & _
     "  Values (" & Me.proiIncidentId & ",  " & Me.proiDatosProductoId & ",  " & Me.proiDetalleDatosProductoId & " )  "
    
    Me.proConexion.Execute varComando
    
    
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    FunGInsertar = False
    SubGMuestraError
End Function

Public Function FunGGuardar() As Boolean
    On Error GoTo ErrManager
    
    If Trim(Me.proiIncidentId) <> "" Then
        If Me.FunGInsertar Then
            FunGGuardar = True
        Else
            FunGGuardar = False
        End If
    End If
    
    Exit Function
ErrManager:
    FunGGuardar = False
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
    Dim varComando As String
    
    On Error GoTo ErrManager
    
    varComando = "  Delete from  ct_ticketxdetalledatosproducto  " & _
     "  Where iIncidentId =  " & Me.proiIncidentId
    Me.proConexion.Execute varComando
    
    FunGEliminar = True
    Exit Function
ErrManager:
    FunGEliminar = False
    SubGMuestraError
End Function
Private Sub Class_Initialize()
    On Error GoTo ErrorManager
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
