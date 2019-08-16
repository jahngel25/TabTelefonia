VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claReglasClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Modulo    : claReglasClasificacionM�dulo de clase
' Fecha  : 27/09/2004 14:56
' Author    : Germ�n A. Fajardo G -  Inform�tica & Tecnologia LTDA.
' Prop�sito   : Administrar la tabla CT_ReglasClasificacion
'---------------------------------------------------------------------------------------


Option Explicit

'propiedades de la tabla
Public proiClasificacionId As String
Public proiReglaId As String

' Clases relacionadas
Public proColReglas As colRegla
Public procolClasificacion As colClasificacion

Public ProDescripcionRegla As String

'Propiedad de conexi�n
Public proConexion As ADODB.Connection

Public Function FunGInsertar() As Boolean
'***********************************************************************
' OBJETIVO: Inserta una Variable
'***********************************************************************
'�PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Germ�n A. Fajardo G.
' FECHA: 02/08/2004
'************************************************************************
On Error GoTo ErrorManager

        
    'Ejecuci�n de Insert
    Me.proConexion.Execute "INSERT INTO CT_REGLASCLASIFICACION (" & _
      "iClasificacionId , iReglaId  )" & _
      "VALUES (" & Me.proiClasificacionId & ", " & proiReglaId & ")"
    
    FunGInsertar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function FunGGuardar() As Boolean
'***********************************************************************
' OBJETIVO: Toma la decisi�n de Insertar o Modificar
'***********************************************************************
'�PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Germ�n A. Fajardo G.
' FECHA: 02/08/2004
'************************************************************************
On Error GoTo ErrorManager

        FunGGuardar = Me.FunGInsertar

    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

    'Se asegura del deseo de eliminaci�n del tramo
    'Ejecuci�n de Insert
    Me.proConexion.Execute "DELETE FROM CT_REGLASCLASIFICACION " & _
        "WHERE iClasificacionId = " & Me.proiClasificacionId & " And iReglaId = " & proiReglaId
     
    Set varResultados = Nothing
    
    If Me.proConexion.Errors.Count Then Exit Function
    FunGEliminar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function


