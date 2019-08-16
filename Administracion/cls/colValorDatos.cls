VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colValordatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"colValordatos"
Attribute VB_Ext_KEY = "Member0" ,"colValordatos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Propiedad de conexión de la coleccion
Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

'Variables para consulta
Public proiValorId As String
Public provchValorDesc As String
Public protiRecordStatus As String

Public Function FunGConsulta() As Boolean
Dim varResultados As ADODB.Recordset
Dim varValordatos As claValordatos
Dim varPosicion As Integer
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
              
        Set mCol = Nothing
        Set mCol = New Collection
        varResultados.Open "SELECT DISTINCT " & _
                                       "iValorId, " & _
                                       "vchValorDesc, " & _
                                       "tiRecordStatus, " & _
                                       " FROM ct_valordatos " & _
                                       " WHERE iValorId =  " & Me.proiValorId
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
                varValordatos.proiValorId = IIf(IsNull(varResultados("iValorId")), "", varResultados.Fields("iValorId"))
                varValordatos.provchValorDesc = IIf(IsNull(varResultados("vchValorDesc")), "", varResultados.Fields("vchValorDesc"))
                varValordatos.protiRecordStatus = IIf(IsNull(varResultados("tiRecordStatus")), "", varResultados.Fields("tiRecordStatus"))
                Me.Add varValordatos.proiValorId, _
                varValordatos.provchValorDesc, _
                varValordatos.protiRecordStatus, _
                , Me.proConexion
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
        
        FunGConsulta = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function Add(proiValorId As Long, provchValorDesc As String, protiRecordStatus As Byte, Optional proConexion As Variant, Optional sKey As String) As claValordatos _
    'create a new object
    Dim objNewMember As claValordatos
    Set objNewMember = New claValordatos

    'set the properties passed into the method
    objNewMember.proiValorId = proiValorId
    objNewMember.provchValorDesc = provchValorDesc
    objNewMember.protiRecordStatus = protiRecordStatus
    If IsMissing(proConexion) = False Then
        Set objNewMember.proConexion = proConexion
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

Public Property Get Item(vntIndexKey As Variant) As claValordatos
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
