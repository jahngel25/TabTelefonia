VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colNovedadValorServicioxNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection

Public proSeleccionados As Integer

'Datos consulta
Public proNumero As String
Public proRegionCode As String
Public proIncidentId As String
Public proNovedadNumeroId As String
Public proServicioSuplementarioId As String
Public proTipoServicio As String 'Parametro Salida

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proNovedadNumeroId As String, proServicioSuplementarioId As String, proValor As String, Optional sKey As String) As claNovedadValorServicioxNumero
                    
    'create a new object
    Dim objNewMember As claNovedadValorServicioxNumero
    Set objNewMember = New claNovedadValorServicioxNumero


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proSeleccion = "0"
    objNewMember.proNovedadNumeroId = proNovedadNumeroId
    objNewMember.proServicioSuplementario = proServicioSuplementarioId
    objNewMember.proValor = proValor
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claNovedadValorServicioxNumero
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
    Me.proSeleccionados = 0
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varNovedad As claNovedadValorServicioxNumero
    Dim i As Integer
    On Error GoTo ErrManager
    'limpiamos la coleccion
    'Limpia la coleccion
    If Me.Count > 0 Then
        For i = Me.Count To 1 Step -1
            Me.Remove (i)
        Next i
    End If
    
    varComando = "Select    N.iNovedadNumeroId,             " & _
                 "          N.iServicioSuplementarioId,     " & _
                 "          N.vchValor,                      " & _
                 "         S.chTipoServicio                  " & _
                 "From      CT_NovedadValorServicioxNumero  N, CT_SERVICIOSSUPLEMENTARIOS S      " & _
                 "Where     N.iServicioSuplementarioId = S.iServicioSuplementarioId  " & _
                 "And       N.iNovedadNumeroId = " & Me.proNovedadNumeroId & " " & _
                 "And       N.iServicioSuplementarioId  = " & Me.proServicioSuplementarioId
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varNovedad = New claNovedadValorServicioxNumero
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iNovedadNumeroId")) Then
            varNovedad.proNovedadNumeroId = ""
        Else
            varNovedad.proNovedadNumeroId = Trim(varResultado.Fields("iNovedadNumeroId"))
        End If
        
        If IsNull(varResultado.Fields("iServicioSuplementarioId")) Then
            varNovedad.proServicioSuplementario = ""
        Else
            varNovedad.proServicioSuplementario = Trim(varResultado.Fields("iServicioSuplementarioId"))
        End If
        
        If IsNull(varResultado.Fields("vchValor")) Then
            varNovedad.proValor = ""
        Else
            varNovedad.proValor = Trim(varResultado.Fields("vchValor"))
        End If
        proTipoServicio = Trim(varResultado.Fields("chTipoServicio"))
        
        Me.Add Me.proConexion, _
               varNovedad.proNovedadNumeroId, _
               varNovedad.proServicioSuplementario, _
               varNovedad.proValor
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varNovedad = Nothing
    
    MetConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


