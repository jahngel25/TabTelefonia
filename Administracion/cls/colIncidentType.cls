VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colIncidentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"claIncidentType"
Attribute VB_Ext_KEY = "Member0" ,"claIncidentType"
'local variable to hold collection
Private mCol As Collection
Option Explicit
Public proParentId As String
Public proConexion As ADODB.Connection
Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varIncidentType As claIncidentType
    On Error GoTo ErrManager
    varComando = "SELECT  iParameterId, vchParameterDesc " _
    & " From ReferenceDefinition " _
    & " WHERE iReferenceId in (" _
    & " SELECT iReferenceId " _
    & " From Reference_Fields " _
    & " WHERE chFieldName = 'incident.type') " _
    & " and tiRecordStatus = 1 " _
    & " and iParentId = " & proParentId
    Set varResultado = New ADODB.Recordset
    Set varIncidentType = New claIncidentType
    Set varIncidentType.proConexion = Me.proConexion
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        'Tipo Incidente
        If IsNull(varResultado.Fields("iParameterId")) Then
            varIncidentType.proParameterId = 0
        Else
            varIncidentType.proParameterId = varResultado.Fields("iParameterId")
        End If
        
        'Tipo Novedad
        If IsNull(varResultado.Fields("vchParameterDesc")) Then
            varIncidentType.provchParameterDesc = ""
        Else
            varIncidentType.provchParameterDesc = varResultado.Fields("vchParameterDesc")
        End If

        Me.Add proParentId, _
                varIncidentType.proParameterId, _
                varIncidentType.provchParameterDesc, proConexion
        
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varIncidentType = Nothing
    
    MetConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError

End Function



Public Function Add(proParentId As String, proParameterId As String, provchParameterDesc As String, proConexion As ADODB.Connection, Optional sKey As String) As claIncidentType
    
    'create a new object
    Dim objNewMember As claIncidentType
    Set objNewMember = New claIncidentType



    'set the properties passed into the method
    objNewMember.proParameterId = proParentId
    objNewMember.proParameterId = proParameterId
    objNewMember.provchParameterDesc = provchParameterDesc
    
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proParentId = proParentId
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claIncidentType
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

