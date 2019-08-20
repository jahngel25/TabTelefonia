VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proUserId As String
Public proUserName As String
Public proDescuentoId As String
Public proConceptoCobroId As String

'Propioedades adicionales del descuento
Public proValorFijo As String
Public proPorcentaje As String

'Propiedades adicionales para Aplicación
Public proAplicacionId As String
Public proPrivilegios As String

Public proConexion As ADODB.Connection

Public Function FunGEliminarApp() As Boolean
'***********************************************************************
' OBJETIVO: Inserta un usuariodescuento
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "DELETE FROM CT_SeguridadAplicacion " & _
      " WHERE chUserId = '" & Me.proUserId & "'" & _
      "  AND iAplicacionId = " & Me.proAplicacionId
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al ingresar los datos
        Exit Function
    End If

    FunGEliminarApp = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGInsertarApp() As Boolean
'***********************************************************************
' OBJETIVO: Inserta un usuariodescuento
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "INSERT INTO CT_SeguridadAplicacion (" & _
      "chUserId, iAplicacionId, vchPrivilegios )" & _
      "VALUES ('" & Me.proUserId & _
                  "', " & Me.proAplicacionId & _
                  ", '" & Me.proPrivilegios & "')"
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al ingresar los datos
        Exit Function
    End If

    FunGInsertarApp = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGInsertar() As Boolean
'***********************************************************************
' OBJETIVO: Inserta un usuariodescuento
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "INSERT INTO CT_UsersDescuento (" & _
      "iDescuentoId, chUserId )" & _
      "VALUES (" & Me.proDescuentoId & _
                  ", '" & Me.proUserId & "')"
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al ingresar los datos
        Exit Function
    End If

    FunGInsertar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
'***********************************************************************
' OBJETIVO: Elimina un usuariodescuento
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "DELETE FROM CT_UsersDescuento WHERE " & _
      "iDescuentoId = " & Me.proDescuentoId & _
      " and chUserId = " & "'" & Me.proUserId & "'"
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al eliminar los datos
        Exit Function
    End If

    FunGEliminar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGInsertarConcepto() As Boolean
'***********************************************************************
' OBJETIVO: Inserta un usuariodescuento
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "INSERT INTO CT_UsersConcepto (" & _
      "iConceptoCobroId, chUserId )" & _
      "VALUES (" & Me.proConceptoCobroId & _
                  ", '" & Me.proUserId & "')"
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al ingresar los datos
        Exit Function
    End If

    FunGInsertarConcepto = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGEliminarConcepto() As Boolean
'***********************************************************************
' OBJETIVO: Elimina un usuario del concepto
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "DELETE FROM CT_UsersConcepto WHERE " & _
      "iConceptoCobroId = " & Me.proConceptoCobroId & _
      " and chUserId = " & "'" & Me.proUserId & "'"
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al eliminar los datos
        Exit Function
    End If

    FunGEliminarConcepto = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGInsertarSinProceso() As Boolean
'***********************************************************************
' OBJETIVO: Inserta un usuariodescuento
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "INSERT INTO CT_UsuariosSinProceso (" & _
      "chUserId)" & _
      "VALUES ('" & Me.proUserId & "')"
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al ingresar los datos
        Exit Function
    End If

    FunGInsertarSinProceso = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGEliminarSinProceso() As Boolean
'***********************************************************************
' OBJETIVO: Elimina un usuariodescuento
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "DELETE FROM CT_UsuariosSinProceso WHERE " & _
     " chUserId = " & "'" & Me.proUserId & "'"
    
    If Me.proConexion.Errors.Count Then 'Si no hubo error al eliminar los datos
        Exit Function
    End If

    FunGEliminarSinProceso = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
