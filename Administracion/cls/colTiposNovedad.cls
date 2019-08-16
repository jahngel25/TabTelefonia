VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTiposNovedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claTiposNovedad"
Attribute VB_Ext_KEY = "Member0" ,"claTiposNovedad"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : colTiposNovedadMódulo de clase
' Fecha  : 02/07/200417:16
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   :Administrar la colección de Tipos Novedad
'---------------------------------------------------------------------------------------


Option Explicit

'Variable de Conexión
Public proConexion As ADODB.Connection

Public proTipoNovedadId As String

'local variable to hold collection
Private mCol As Collection
Public Function MetConsultar() As Boolean

    Dim varResultados As ADODB.Recordset
    Dim varComando As String
    Dim varTiposNovedad As claTiposNovedad
    On Error GoTo ErrorManager


    varComando = "SELECT iTipoNovedadId, " & _
                 "          vchDescripcionNovedad,   " & _
                 "          tiRecordStatus  " & _
                 "FROM      CT_TIPOSNOVEDAD   "
                 
    'Trae los datos de la base
    Set varResultados = New ADODB.Recordset
        
    varResultados.Open varComando, Me.proConexion
        
    Set varTiposNovedad = New claTiposNovedad
    'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
    While varResultados.EOF = False
        
        'Codigo
        If IsNull(varResultados.Fields("iTipoNovedadId")) Then
            varTiposNovedad.proTipoNovedadId = 0
        Else
            varTiposNovedad.proTipoNovedadId = varResultados.Fields("iTipoNovedadId")
        End If
        
        'Descripcion
        If IsNull(varResultados.Fields("vchDescripcionNovedad")) Then
            varTiposNovedad.proTipoNovedadId = ""
        Else
            varTiposNovedad.proDescripcionNovedad = varResultados.Fields("vchDescripcionNovedad")
        End If
        
        'Estado
        If IsNull(varResultados.Fields("tiRecordStatus")) Then
            varTiposNovedad.proRecordStatus = 0
        Else
            varTiposNovedad.proRecordStatus = varResultados.Fields("tiRecordStatus")
        End If
        
        Me.Add Me.proConexion, _
                varTiposNovedad.proRecordStatus, _
                varTiposNovedad.proDescripcionNovedad, _
                varTiposNovedad.proTipoNovedadId
                    
        'Siguiente registro del cursor
        varResultados.MoveNext
    Wend
                
        Set varResultados = Nothing
        Set varTiposNovedad = Nothing
        
        MetConsultar = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function Add(proConexion As ADODB.Connection, proRecordStatus As String, proTiposNovedadDesc As String, proTiposNovedadId As String, Optional sKey As String) As claTiposNovedad
    'create a new object
    Dim objNewMember As claTiposNovedad
    Set objNewMember = New claTiposNovedad


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proRecordStatus = proRecordStatus
    objNewMember.proDescripcionNovedad = proTiposNovedadDesc
    objNewMember.proTipoNovedadId = proTiposNovedadId
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claTiposNovedad
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

