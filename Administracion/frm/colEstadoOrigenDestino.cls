VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colEstadoOrigenDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'       DESCRIPCION: Clase que mantiene una coleccion de claEstadoOrigenDestino
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Desarrollo inicial de la clase
'       VERSION:              1.0.000
'       REQUERIMIENTO:        5322
'       FECHA:       2009/10/10
'*******************************************************************
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proConexion As ADODB.Connection
Public proTipoEstado As String


Public Function Add(proConexion As ADODB.Connection, proTipoEstado As String, proEstadoNumero As String, proDescripcionEstado As String, Optional sKey As String) As claEstadoOrigenDestino
    'create a new object
    Dim objNewMember As claEstadoOrigenDestino
    Set objNewMember = New claEstadoOrigenDestino


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proTipoEstado = proTipoEstado
    objNewMember.proEstadoNumero = proEstadoNumero
    objNewMember.proDescripcion = proDescripcionEstado
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claEstadoOrigenDestino
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

Public Function FunGConsultaEstadosPorTipo() As Boolean
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    Dim varEstado As claEstadoOrigenDestino
    
    On Error GoTo ErrManager
    
    Set varComando = New ADODB.Command
    Set varComando.ActiveConnection = Me.proConexion
    varComando.CommandType = adCmdStoredProc
    varComando.CommandText = "CRMCP_ConsultaEstadosPorTipo"
    
   
    varComando.Parameters.Item("@vchTipoEstado").Value = Me.proTipoEstado
   
   
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute
    
    Set varEstado = New claEstadoOrigenDestino
    
    While Not varResultado.EOF
        varEstado.proEstadoNumero = varResultado.Fields("chEstadoNumero")
        varEstado.proDescripcion = varResultado.Fields("vchDescripcionEstado")
        
        Me.Add Me.proConexion, _
               varEstado.proTipoEstado, _
               varEstado.proEstadoNumero, _
               varEstado.proDescripcion

        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varEstado = Nothing
    
    FunGConsultaEstadosPorTipo = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


