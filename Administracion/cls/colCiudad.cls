VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claCiudad"
Attribute VB_Ext_KEY = "Member0" ,"claCiudad"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proNombreCiudad As String, proCodigoCiudad As String, Optional sKey As String) As claCiudad
    'create a new object
    Dim objNewMember As claCiudad
    Set objNewMember = New claCiudad


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proNombreCiudad = proNombreCiudad
    objNewMember.proCodigoCiudad = proCodigoCiudad
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claCiudad
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

Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varCiudad As claCiudad
    On Error GoTo ErrManager
    
    varComando = "Select    chRegionCode,   " & _
                 "          chRegionName    " & _
                 "From      Region          " & _
                 "Order by  chRegionName    "
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varCiudad = New claCiudad
    Set varCiudad.proConexion = Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varCiudad.proCodigoCiudad = ""
        Else
            varCiudad.proCodigoCiudad = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("chRegionName")) Then
            varCiudad.proNombreCiudad = ""
        Else
            varCiudad.proNombreCiudad = Trim(varResultado.Fields("chRegionName"))
        End If
        
        Me.Add varCiudad.proConexion, _
               varCiudad.proNombreCiudad, _
               varCiudad.proCodigoCiudad
               
        varResultado.MoveNext
    Wend
    
    Set varCiudad = Nothing
    Set varResultado = Nothing
    
    MetConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

