VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proClasificacionId As String
Public proClasificacion As String
Public proRecordStatus As String
Public proColRegla As colReglasClasificacion
Public proSeleccionado As String

Public proConexion As ADODB.Connection



Public Function FunGInsertar() As Boolean

Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

     Me.proRecordStatus = 1
    'Ejecución de Insert
    Me.proConexion.Execute "INSERT INTO CT_Clasificacion (" & _
      "vchClasificacion, tiRecordStatus)" & _
      "VALUES (" & _
     "'" & Me.proClasificacion & "', '1')"
    
    If Me.proConexion.Errors.Count = 0 Then 'Si no hubo error al ingresar los datos
        'Consulta del código que acaba de insertar
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT proClasificacionId = max(iClasificacionId) FROM CT_Clasificacion", Me.proConexion
                       
        If Not (varResultados.EOF) Then
                Me.proClasificacionId = varResultados("proClasificacionId").Value
        End If
    End If

    FunGInsertar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGModificar() As Boolean

On Error GoTo ErrorManager

    'Ejecución de Insert
    Me.proConexion.Execute "UPDATE CT_Clasificacion " & _
      "SET vchClasificacion = '" & Me.proClasificacion & "', " & _
      "    tiRecordStatus = '" & Me.proRecordStatus & "'" & _
       "WHERE iClasificacionId = " & Me.proClasificacionId
    
    If Me.proConexion.Errors.Count Then Exit Function
    FunGModificar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGGuardar() As Boolean
'***********************************************************************
' OBJETIVO: Toma la decisión de Insertar o Modificar
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
On Error GoTo ErrorManager

    If Len(Trim(Me.proClasificacionId)) = 0 Then
        FunGGuardar = Me.FunGInsertar
    Else
        FunGGuardar = Me.FunGModificar
    End If
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

    'Eliminación Lógica coloca el recordstatus = 0
    Me.proRecordStatus = 0
    
    FunGEliminar = Me.FunGModificar
    
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function





