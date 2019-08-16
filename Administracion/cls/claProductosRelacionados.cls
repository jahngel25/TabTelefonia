VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claProductosRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proProductNumber As String
Public provchDescription As String
Public proConexion As ADODB.Connection
Public proColProductosRelacionados As colProductosRelacionados
Public Function FunGInsertar() As Boolean
'***********************************************************************
' OBJETIVO: Inserta un producto
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Hernan Botache
' Fecha 02/09/2004
'************************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

        
    'Ejecución de Insert
    Me.proConexion.Execute "INSERT INTO CT_ProductosRelacionados (" & _
      "chProductNumber )" & _
      "VALUES (" & _
                  "'" & Me.proProductNumber & "')"
    
    FunGInsertar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function


Public Function FunGGuardar() As Boolean

    FunGGuardar = Me.FunGInsertar
    Exit Function

ErrorManager:
    SubGMuestraError
End Function

Public Function FunGEliminar() As Boolean
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

    'Se asegura del deseo de eliminación del tramo
    'Ejecución de Insert
    Me.proConexion.Execute "DELETE FROM CT_ProductosRelacionados " & _
        "WHERE chProductNumber = '" & Trim(Me.proProductNumber) & "'"
     
    Set varResultados = Nothing
    
    If Me.proConexion.Errors.Count Then Exit Function
    FunGEliminar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGEliminarProducto(parVariable As claProductosRelacionados) As Boolean
'***********************************************************************
' OBJETIVO: Elimina un producto relacionado
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Hernan Botache
' FECHA: 02/09/2004
'************************************************************************
Dim varProducto As claProductosRelacionados
Dim varEncontro As Boolean
Dim varCuenta As Integer
On Error GoTo ErrorManager


    If MsgBox("Esta seguro de excluir el producto " & parVariable.provchDescription & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
        FunGEliminarProducto = True
        Exit Function
    End If
    
    'Instancia del objeto
    Set varProducto = New claProductosRelacionados
    Set varProducto.proConexion = parVariable.proConexion
    
    'Copia las propiedades
    varProducto.provchDescription = parVariable.provchDescription
    varProducto.proProductNumber = parVariable.proProductNumber
    FunGEliminarProducto = varProducto.FunGEliminar
    
    If FunGEliminarProducto Then
         varEncontro = False
        varCuenta = 1
        While varCuenta <= Me.proColProductosRelacionados.Count And varEncontro = False
            If Me.proColProductosRelacionados.Item(varCuenta).provchDescription = parVariable.provchDescription Then
                Me.proColProductosRelacionados.Remove varCuenta
                varEncontro = True
            Else
                varCuenta = varCuenta + 1
            End If
        Wend
    End If
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function


Public Function FunGAgregarProducto(parVariable As claProductosRelacionados) As Boolean
'***********************************************************************
' OBJETIVO: Agrega un producto
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Hernan Botache
' FECHA: 02/09/2004
'************************************************************************
Dim varProducto As claProductosRelacionados
On Error GoTo ErrorManager

    Set varProducto = New claProductosRelacionados
    Set varProducto.proConexion = Me.proConexion

     varProducto.proProductNumber = parVariable.proProductNumber
    varProducto.provchDescription = parVariable.provchDescription
    FunGAgregarProducto = varProducto.FunGInsertar

    If FunGAgregarProducto Then
        'Agrega a la colecciones
        Me.proColProductosRelacionados.Add Me.proConexion, parVariable.proProductNumber, parVariable.provchDescription
    End If
    Exit Function

ErrorManager:
    SubGMuestraError
End Function


