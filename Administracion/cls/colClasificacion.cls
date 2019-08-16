VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claClasificacion"
Attribute VB_Ext_KEY = "Member0" ,"claClasificacion"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public proConexion As ADODB.Connection
Public Function Add(proConexion As ADODB.Connection, proRecordStatus As String, proClasificacion As String, proClasificacionId As String, Optional sKey As String) As claClasificacion
    'create a new object
    Dim objNewMember As claClasificacion
    Set objNewMember = New claClasificacion


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proRecordStatus = proRecordStatus
    objNewMember.proClasificacion = proClasificacion
    objNewMember.proClasificacionId = proClasificacionId
    objNewMember.proSeleccionado = "N"
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claClasificacion
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

Public Function FunGConsulta() As Boolean

Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT * FROM CT_Clasificacion", Me.proConexion
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add Me.proConexion, varResultados("tiRecordStatus").Value, _
                    varResultados("vchClasificacion").Value, varResultados("iclasificacionid").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        FunGConsulta = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function FunGEliminarClasificacion(parItem As Integer) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación lógica de una clasificacion
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Hernan Botache
' FECHA: 02/09/2004
'****************************************************************
On Error GoTo ErrorManager

        
        If MsgBox("Esta seguro de eliminar la clasificación " & Me.Item(parItem).proClasificacion & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            FunGEliminarClasificacion = True
            Exit Function
        End If
    
        Set Me.Item(parItem).proConexion = Me.proConexion
        
        FunGEliminarClasificacion = Me.Item(parItem).FunGEliminar
        
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
