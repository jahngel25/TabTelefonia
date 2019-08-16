VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

'Aplicación
Public proAplicacionId As String

'Conexion
Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proPrivilegios As String, proAplicacionId As Integer, proUserId As String, Optional sKey As String) As claSeguridad
    'create a new object
    Dim objNewMember As claSeguridad
    Set objNewMember = New claSeguridad


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proPrivilegios = proPrivilegios
    objNewMember.proAplicacionId = proAplicacionId
    objNewMember.proUserId = proUserId
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claSeguridad
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


Public Function FunGConsulta() As Boolean
    On Error GoTo ErrManager
    
    Dim varResultado As ADODB.Recordset
    Dim varSeguridad As claSeguridad
    
    Dim varComando As String
    
    Set varResultado = New ADODB.Recordset
    
    If Len(Trim(Me.proAplicacionId)) = 0 Then
            Exit Function
    End If
    
    varComando = "Select chUserId,  " & _
                 "       iAplicacionId, " & _
                 "       vchPrivilegios " & _
                 "From   CT_SeguridadAplicacion " & _
                 "Where  iAplicacionId =" & Me.proAplicacionId & _
                 " order by chUserId"
                 
    varResultado.Open varComando, Me.proConexion
    
    Set varSeguridad = New claSeguridad
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chUserId")) Then
            varSeguridad.proUserId = ""
        Else
            varSeguridad.proUserId = Trim(varResultado.Fields("chUserId"))
        End If
        
        If IsNull(varResultado.Fields("iAplicacionId")) Then
            varSeguridad.proAplicacionId = 0
        Else
            varSeguridad.proAplicacionId = varResultado.Fields("iAplicacionId")
        End If
        
        If IsNull(varResultado.Fields("vchPrivilegios")) Then
            varSeguridad.proPrivilegios = ""
        Else
            varSeguridad.proPrivilegios = Trim(varResultado.Fields("vchPrivilegios"))
        End If
        
        Me.Add Me.proConexion, _
                varSeguridad.proPrivilegios, _
                varSeguridad.proAplicacionId, _
                varSeguridad.proUserId
        varResultado.MoveNext
    Wend
    
    Set varSeguridad = Nothing
    Set varResultado = Nothing
    
    FunGConsulta = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function



