VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'       DESCRIPCION: Clase que mantiene una coleccion de claCategoria
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Desarrollo inicial de la clase
'       VERSION:       1.0.000
'       REQUERIMIENTO: 3488
'       FECHA:       2009/08/05
'*******************************************************************
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proCategoriaID As String
'Conexion
Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, parCategoriaID As String, parDescripcion As String) As claCategoria
    'create a new object
    Dim objNewMember As claCategoria
    Set objNewMember = New claCategoria
 
    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proCategoriaID = parCategoriaID
    objNewMember.proDescripcion = parDescripcion
    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claCategoria
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
    
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    Dim varCategoria As claCategoria
    
    On Error GoTo ErrManager
        
    
    Set varComando = New ADODB.Command
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CRMCP_ConsultaCategorias"
    varComando.CommandTimeout = 0
    
   
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute
    
    Set varCategoria = New claCategoria

    
    While Not varResultado.EOF
        varCategoria.proCategoriaID = Trim(varResultado.Fields("iParameterId"))
        varCategoria.proDescripcion = Trim(varResultado.Fields("vchParameterDesc"))
        
        Me.Add Me.proConexion, varCategoria.proCategoriaID, _
                               varCategoria.proDescripcion
                        
        varResultado.MoveNext
    Wend
    
    Set varCategoria = Nothing
    Set varResultado = Nothing
    
    MetConsultar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function






