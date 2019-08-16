VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claUsuario"
Attribute VB_Ext_KEY = "Member0" ,"claUsuario"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variable de Conexión
Public proConexion As ADODB.Connection

'Descuento para el cual los usuarios de la colección autorizan
Public proDescuentoId As String

'Concepto de Cobro para los usuarios que autorizan el concepto
Public proConceptoCobroId As String

'ID de la aplicación
Public proAplicacionId As String

'Propiedad que indica los permitidos
Public proPermitidos As String

'local variable to hold collection
Private mCol As Collection

Public Function FunGAgregarUsuarioSinProceso(parUsuario As claUsuario) As Boolean
'****************************************************************
' OBJETIVO: Agrega un usuario que no va a tener validaciones del proceso
'****************************************************************
' PARITEM:  Usuario a agregar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 21/01/2002
'****************************************************************
Dim varUsuario As claUsuario
On Error GoTo ErrorManager

        
        'Instancia del objeto usuario
        Set varUsuario = New claUsuario
        Set varUsuario.proConexion = Me.proConexion
        
        varUsuario.proUserId = parUsuario.proUserId
        
        FunGAgregarUsuarioSinProceso = varUsuario.FunGInsertarSinProceso
        
        If FunGAgregarUsuarioSinProceso = True Then
            'Elimina el segmento de la coleccion
            Me.Add parUsuario.proUserName, parUsuario.proUserId, _
                   varUsuario.proAplicacionId, varUsuario.proPrivilegios, _
                   varUsuario.proValorFijo, varUsuario.proPorcentaje
        End If
        
        Set varUsuario = Nothing
        
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Public Function FunGEliminarUsuarioSinProceso(parItem As Integer) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación de una variable
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'****************************************************************
On Error GoTo ErrorManager

        
        If MsgBox("Esta seguro de eliminar al usuario autorizado " & Trim(Me.Item(parItem).proUserName) & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            FunGEliminarUsuarioSinProceso = True
            Exit Function
        End If
    
        Set Me.Item(parItem).proConexion = Me.proConexion
        FunGEliminarUsuarioSinProceso = Me.Item(parItem).FunGEliminarSinProceso
        
        If FunGEliminarUsuarioSinProceso = True Then
            'Elimina el segmento de la coleccion
            Me.Remove parItem
        End If
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function FunGAgregarUsuarioApp(parUsuario As claUsuario) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación de un usuario de la aplicación
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'****************************************************************
Dim varUsuario As claUsuario
On Error GoTo ErrorManager

        
        'Instancia del objeto usuario
        Set varUsuario = New claUsuario
        Set varUsuario.proConexion = Me.proConexion
        
        varUsuario.proAplicacionId = Me.proAplicacionId
        varUsuario.proUserId = parUsuario.proUserId
        varUsuario.proPrivilegios = ""
        
        FunGAgregarUsuarioApp = varUsuario.FunGInsertarApp
        
        If FunGAgregarUsuarioApp = True Then
            'Elimina el segmento de la coleccion
            Me.Add parUsuario.proUserName, parUsuario.proUserId, _
                   varUsuario.proAplicacionId, varUsuario.proPrivilegios, _
                   varUsuario.proValorFijo, varUsuario.proPorcentaje
        End If
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function FunGEliminarUsuarioApp(parItem As Integer) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación de un usuario de la aplicación
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'****************************************************************
On Error GoTo ErrorManager

        
        If MsgBox("Esta seguro de eliminar al usuario " & Trim(Me.Item(parItem).proUserName) & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            FunGEliminarUsuarioApp = True
            Exit Function
        End If
    
        Set Me.Item(parItem).proConexion = Me.proConexion
        Me.Item(parItem).proAplicacionId = Me.proAplicacionId
        
        FunGEliminarUsuarioApp = Me.Item(parItem).FunGEliminarApp
        
        If FunGEliminarUsuarioApp = True Then
            'Elimina el segmento de la coleccion
            Me.Remove parItem
        End If
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function FunGEliminarUsuario(parItem As Integer) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación de una variable
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'****************************************************************
On Error GoTo ErrorManager

        
        If MsgBox("Esta seguro de eliminar al usuario autorizador " & Trim(Me.Item(parItem).proUserName) & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            FunGEliminarUsuario = True
            Exit Function
        End If
    
        Set Me.Item(parItem).proConexion = Me.proConexion
        Me.Item(parItem).proDescuentoId = Me.proDescuentoId
        FunGEliminarUsuario = Me.Item(parItem).FunGEliminar
        
        If FunGEliminarUsuario = True Then
            'Elimina el segmento de la coleccion
            Me.Remove parItem
        End If
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function FunGEliminarUsuarioConcepto(parItem As Integer) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación de un usuario del concepto
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 10/09/2001
'****************************************************************
On Error GoTo ErrorManager

        
        If MsgBox("Esta seguro de eliminar al usuario " & Trim(Me.Item(parItem).proUserName) & " del concepto?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            FunGEliminarUsuarioConcepto = True
            Exit Function
        End If
    
        Set Me.Item(parItem).proConexion = Me.proConexion
        Me.Item(parItem).proConceptoCobroId = Me.proConceptoCobroId
        FunGEliminarUsuarioConcepto = Me.Item(parItem).FunGEliminarConcepto
        
        If FunGEliminarUsuarioConcepto = True Then
            'Elimina el segmento de la coleccion
            Me.Remove parItem
        End If
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Public Function FunGEliminarTodosUsuarioConcepto() As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación de un usuario del concepto
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 10/09/2001
'****************************************************************
On Error GoTo ErrorManager

        
        Me.proConexion.Execute "DELETE FROM CT_UsersConcepto " & _
                                               "WHERE iConceptoCobroId = " & Me.proConceptoCobroId
        
        If Me.proConexion.Errors.Count <> 0 Then Exit Function
        
        FunGEliminarTodosUsuarioConcepto = True
        'Elimina todos los elementos de la colección
        While Me.Count
                Me.Remove 1
        Wend

        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function FunGConsultaSinProceso() As Boolean
'********************************************************************
' OBJETIVO: Consulta de todos los usuarios por aplicación
'********************************************************************
' PARITEM:  Item a eliminar
'********************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'********************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager
   
        

        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT U.* FROM CT_UsuariosSinProceso S, Users U " & _
                           "WHERE U.chUserId = S.chUserId " & _
                           "Order by U.chUserName", Me.proConexion
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add varResultados("chUserName").Value, _
                    varResultados("chUserId").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        Set varResultados = Nothing
        FunGConsultaSinProceso = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGConsultaxApp() As Boolean
'********************************************************************
' OBJETIVO: Consulta de todos los usuarios por aplicación
'********************************************************************
' PARITEM:  Item a eliminar
'********************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'********************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager
   
        
        If Len(Trim(Me.proAplicacionId)) = 0 Then
                FunGConsultaxApp = True
                Exit Function
        End If
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT U.*, S.vchPrivilegios  FROM CT_SeguridadAplicacion S, Users U " & _
                           "WHERE S.iAplicacionId = " & Me.proAplicacionId & _
                           " and U.chUserId = S.chUserId " & _
                           "Order by U.chUserName", Me.proConexion
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add varResultados("chUserName").Value, _
                    varResultados("chUserId").Value, _
                    Me.proAplicacionId, _
                    varResultados("vchPrivilegios").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        Set varResultados = Nothing
        FunGConsultaxApp = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function FunGConsultaAprobacionNumeros() As Boolean
'********************************************************************
' OBJETIVO: Consulta de todos los usuarios que pueden aprobar numeros por clasificación
'********************************************************************
' PARITEM:
'********************************************************************
' AUTOR: Diana Buenhombre
' FECHA: 31/01/2006
'********************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager
   
        
        
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "select Distinct U.chUserId , U.chUsername " & _
                           " from ct_usersclasificacion CU, users U " & _
                           " Where CU.tiRecordStatus = 1 " & _
                           " and CU.chUserId = U.chUserId " & _
                           " and U.tiRecordStatus = 1 ", Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add varResultados("chUserName").Value, _
                    varResultados("chUserId").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        Set varResultados = Nothing
        FunGConsultaAprobacionNumeros = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function FunGConsultaxDescuento() As Boolean
'********************************************************************
' OBJETIVO: Consulta de todos los usuarios que autorizan un descuento
'********************************************************************
' PARITEM:  Item a eliminar
'********************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'********************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager
   
        
        If Len(Trim(Me.proDescuentoId)) = 0 Then
                FunGConsultaxDescuento = True
                Exit Function
        End If
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT U.* FROM CT_UsersDescuento UD, Users U " & _
                           "WHERE UD.iDescuentoId = " & Me.proDescuentoId & _
                           " and U.chUserId = UD.chUserId " & _
                           "Order by U.chUserName", Me.proConexion
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add varResultados("chUserName").Value, _
                    varResultados("chUserId").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        Set varResultados = Nothing
        FunGConsultaxDescuento = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGConsultaxConcepto() As Boolean
'********************************************************************
' OBJETIVO: Consulta de todos los usuarios que pueden usar el concepto
'********************************************************************
' PARAMETROS:  Ninguno
'********************************************************************
' AUTOR: Raúl Cruz
' FECHA: 10/09/2001
'********************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager
   
        
        If Len(Trim(Me.proConceptoCobroId)) = 0 Then
                FunGConsultaxConcepto = True
                Exit Function
        End If
        
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT U.* FROM CT_UsersConcepto UD, Users U " & _
                           "WHERE UD.iConceptoCobroId = " & Me.proConceptoCobroId & _
                           " and U.chUserId = UD.chUserId " & _
                           "Order by U.chUserName", Me.proConexion
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add varResultados("chUserName").Value, _
                    varResultados("chUserId").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        Set varResultados = Nothing
        FunGConsultaxConcepto = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGConsultaxPermitidos() As Boolean
'********************************************************************
' OBJETIVO: Consulta de todos los usuarios permitidos por el proceso
'********************************************************************
' PARAMETROS:  Ninguno
'********************************************************************
' AUTOR: Raúl Cruz
' FECHA: 10/09/2001
'********************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager
   
        
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT U.* FROM Users U, Groups G  " & _
                           "WHERE U.chGroupId = G.chGroupId " & _
                           " AND  G.chGroupDesc like '%" & Me.proPermitidos & "%'" & _
                           " AND U.tiRecordSTatus = 1 " & _
                           " Order by U.chUserName", Me.proConexion
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add Trim(varResultados("chUserName").Value), _
                    varResultados("chUserId").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        Set varResultados = Nothing
        FunGConsultaxPermitidos = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGConsulta() As Boolean
'****************************************************************
' OBJETIVO: Consulta de todos los usuarios de ONYX
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'****************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager
   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT * FROM Users Where tiRecordStatus = 1 Order by chUserName", Me.proConexion
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add Trim(varResultados("chUserName").Value), _
                    Trim(varResultados("chUserId").Value)
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        Set varResultados = Nothing
        FunGConsulta = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function Add(proUserName As String, proUserId As String, Optional proAplicacion As Variant, Optional proPrivilegios As String, Optional proValorFijo As Variant, Optional proPorcentaje As Variant, Optional sKey As String) As claUsuario
    'create a new object
    Dim objNewMember As claUsuario
    Set objNewMember = New claUsuario

    objNewMember.proUserName = proUserName
    objNewMember.proUserId = proUserId
    
    If IsMissing(proAplicacion) = False Then
        objNewMember.proAplicacionId = proAplicacion
    End If
    If IsMissing(proPrivilegios) = False Then
        objNewMember.proPrivilegios = proPrivilegios
    End If
    If IsMissing(proValorFijo) = False Then
        objNewMember.proValorFijo = proValorFijo
    End If
    If IsMissing(proPorcentaje) = False Then
        objNewMember.proPorcentaje = proPorcentaje
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As claUsuario
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

