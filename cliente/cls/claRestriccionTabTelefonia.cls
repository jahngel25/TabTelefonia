VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claRestriccionTabTel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colRestriccionTabFacturacion"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Public proiIncidentTypeId As String
Public proConexion As ADODB.Connection
Public Property Get colRestriccionTabFac() As colRestriccionTabTel
    If mvarcolRestriccionTabFac Is Nothing Then
        Set mvarcolRestriccionTabFac = New colRestriccionTabTel
    End If
    Set colRestriccionTabFac = mvarcolRestriccionTabFac
End Property
Public Property Set colRestriccionTabFac(vData As colRestriccionTabTel)
    Set mvarcolRestriccionTabFac = vData
End Property

Public Function FunGInsertar() As Boolean
'***********************************************************************
' OBJETIVO: Inserta una Restricción
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Germán Fajardo
' FECHA: 17/06/2004
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    proConexion.Execute "INSERT INTO CT_RestriccionTabTelefonia (" & _
      "iIncidentTypeId)" & _
      "VALUES ('" & proiIncidentTypeId & ")"

    FunGInsertar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGModificar() As Boolean
'***********************************************************************
' OBJETIVO: Modifica una Restricción
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Germán Fajardo
' FECHA: 17/06/2004
'************************************************************************
On Error GoTo ErrorManager

    'Ejecución de Insert
    Me.proConexion.Execute "UPDATE CT_RestriccionTabTelefonia " & _
      "SET iIncidentTypeId = '" & proiIncidentTypeId _
      & " WHERE iIncidentTypeId = " & proiIncidentTypeId
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
' AUTOR: Germán Fajardo
' FECHA: 19/05/2004
'************************************************************************
On Error GoTo ErrorManager

    If Len(Trim(proiIncidentTypeId)) = 0 Then
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

    'Se asegura del deseo de eliminación del tramo
    'Ejecución de Insert
    Me.proConexion.Execute "DELETE FROM CT_RestriccionTabTelefonia " & _
        "WHERE iIncidentTypeId = " & proiIncidentTypeId
     
    Set varResultados = Nothing
    
    If Me.proConexion.Errors.Count Then Exit Function
    FunGEliminar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

