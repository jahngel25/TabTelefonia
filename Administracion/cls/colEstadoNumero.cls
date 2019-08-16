VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colEstadoNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claEstadoNumero"
Attribute VB_Ext_KEY = "Member0" ,"claEstadoNumero"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proConexion As ADODB.Connection


Public Function Add(proConexion As ADODB.Connection, proDescripcionEstado As String, proEstadoNumero As String, Optional sKey As String) As claEstadoNumero
    'create a new object
    Dim objNewMember As claEstadoNumero
    Set objNewMember = New claEstadoNumero


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proDescripcionEstado = proDescripcionEstado
    objNewMember.proEstadoNumero = proEstadoNumero
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claEstadoNumero
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

Public Function MetConsulta() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varEstadoNumero As claEstadoNumero
    On Error GoTo ErrManager
    
    varComando = "Select    chEstadoNumero,         " & _
                 "          vchDescripcionEstado    " & _
                 "From      CT_EstadoNumero         "
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varEstadoNumero = New claEstadoNumero
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chEstadoNumero")) Then
            varEstadoNumero.proEstadoNumero = ""
        Else
            varEstadoNumero.proEstadoNumero = varResultado.Fields("chEstadoNumero")
        End If
        
        If IsNull(varResultado.Fields("vchDescripcionEstado")) Then
            varEstadoNumero.proDescripcionEstado = ""
        Else
            varEstadoNumero.proDescripcionEstado = varResultado.Fields("vchDescripcionEstado")
        End If
        
        Me.Add Me.proConexion, _
               varEstadoNumero.proDescripcionEstado, _
               varEstadoNumero.proEstadoNumero
               
        varResultado.MoveNext
    Wend
    
    Set varEstadoNumero = Nothing
    Set varResultado = Nothing
    
    MetConsulta = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
