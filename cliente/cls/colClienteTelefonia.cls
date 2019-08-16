VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colClienteTelefonia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claClienteTelefonia"
Attribute VB_Ext_KEY = "Member0" ,"claClienteTelefonia"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proCallSource As String, proGrupoCentrex As String, proCompanyId As String, Optional sKey As String) As claClienteTelefonia
    'create a new object
    Dim objNewMember As claClienteTelefonia
    Set objNewMember = New claClienteTelefonia


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proCallSource = proCallSource
    objNewMember.proGrupoCentrex = proGrupoCentrex
    objNewMember.proCompanyId = proCompanyId
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claClienteTelefonia
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

Public Function MetConsultarGrupoCentrexOcupado() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varGrupoCentrex As String
    On Error GoTo ErrManager
    
    varComando = "Select distinct vchGrupoCentrex   " & _
                 "From   CT_ClienteTelefonia        "
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("vchGrupoCentrex").Value) Then
            varGrupoCentrex = ""
        Else
            varGrupoCentrex = Trim(varResultado.Fields("vchGrupoCentrex").Value)
        End If
        
        Me.Add Me.proConexion, _
               "", _
               varGrupoCentrex, _
               ""
        varResultado.MoveNext
    Wend
    
    MetConsultarGrupoCentrexOcupado = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarCallSourceOcupado() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varCallSource As String
    On Error GoTo ErrManager
    
    varComando = "Select distinct vchCallSource " & _
                 "From   CT_ClienteTelefonia    "
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("vchCallSource").Value) Then
            varCallSource = ""
        Else
            varCallSource = Trim(varResultado.Fields("vchCallSource").Value)
        End If
        
        Me.Add Me.proConexion, _
               varCallSource, _
               "", _
               ""
        varResultado.MoveNext
    Wend
    
    MetConsultarCallSourceOcupado = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


