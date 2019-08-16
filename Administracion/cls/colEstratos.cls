VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colEstratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proEstratoID As String
'Conexion
Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, parEstratoId As String, parDescripcion As String) As claEstratos
    'create a new object
    Dim objNewMember As claEstratos
    Set objNewMember = New claEstratos
 
    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proEstratoID = parEstratoId
    objNewMember.proDescripcion = parDescripcion
    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claEstratos
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

Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varEstractos As claEstratos
    On Error GoTo ErrManager
        
    varComando = "SELECT iEstratoID , vchDescripcion " & _
                 "From  CT_TelefoniaEstratos " & _
                 "WHERE tiRecordStatus = 1 " & _
                 "ORDER BY iEstratoID "

    Set varResultado = New ADODB.Recordset
    Set varEstractos = New claEstratos
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        
        If IsNull(varResultado.Fields("iEstratoID")) Then
            varEstractos.proEstratoID = "0"
        Else
            varEstractos.proEstratoID = Trim(varResultado.Fields("iEstratoID"))
        End If
    
        If IsNull(varResultado.Fields("vchDescripcion")) Then
            varEstractos.proDescripcion = ""
        Else
            varEstractos.proDescripcion = Trim(varResultado.Fields("vchDescripcion"))
        End If
        
        Me.Add Me.proConexion, varEstractos.proEstratoID, _
                               varEstractos.proDescripcion
                        
        varResultado.MoveNext
    Wend
    Set varEstractos = Nothing
    Set varResultado = Nothing
    
    MetConsultar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function








