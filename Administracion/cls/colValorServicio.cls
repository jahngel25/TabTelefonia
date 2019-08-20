VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colValorServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Propiedad de conexión de la coleccion
Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

'Variables para consulta
Public proServicioSuplementarioId As String
Public proValor As String
Public proDefault As String
Public protiRecordStatus As String


Public Function FunGConsulta() As Boolean
Dim varResultados As ADODB.Recordset
Dim varValorServios As claValorServicio
Dim varPosicion As Integer
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        Set varValorServios = New claValorServicio
              
        Set mCol = Nothing
        Set mCol = New Collection
        varResultados.Open "SELECT DISTINCT " & _
                                       "iServicioSuplementarioId, " & _
                                       "vchValor, " & _
                                       "bDefault, " & _
                                       "tiRecordStatus " & _
                                       " FROM ct_valoresservicio " & _
                                       " WHERE tiRecordStatus = 1 and iServicioSuplementarioId =  " & Me.proServicioSuplementarioId, Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
                varValorServios.proServicioSuplementarioId = IIf(IsNull(varResultados("iServicioSuplementarioId")), "", varResultados.Fields("iServicioSuplementarioId"))
                varValorServios.proValor = IIf(IsNull(varResultados("vchValor")), "", varResultados.Fields("vchValor"))
                varValorServios.proDefault = IIf(IsNull(varResultados("bDefault")), "", varResultados.Fields("bDefault"))
                varValorServios.protiRecordStatus = IIf(IsNull(varResultados("tiRecordStatus")), "", varResultados.Fields("tiRecordStatus"))
                Me.Add varValorServios.proServicioSuplementarioId, _
                varValorServios.proValor, _
                varValorServios.proDefault, _
                varValorServios.protiRecordStatus, _
                Me.proConexion
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
        
        FunGConsulta = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function Add(proServicioSuplementarioId As Long, proValor As String, proDefault As String, protiRecordStatus As Byte, Optional proConexion As Variant, Optional sKey As String) As claValorServicio

    'create a new object
    Dim objNewMember As claValorServicio
    Set objNewMember = New claValorServicio

    'set the properties passed into the method
    objNewMember.proServicioSuplementarioId = proServicioSuplementarioId
    objNewMember.proDefault = proDefault
    objNewMember.proValor = proValor
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

Public Property Get Item(vntIndexKey As Variant) As claValorServicio
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


