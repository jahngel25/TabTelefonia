VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colValorServicioxnumero"
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
Public proServicioSuplementarioId As String
Public proValor As String
Public proTipoServicio As String

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proServicioSuplementarioId As String, proRegionCode As String, proNumero As String, proValor As String, proTipoServicio As String, Optional sKey As String) As claValorServicioxnumero
                    
    'create a new object
    Dim objNewMember As claValorServicioxnumero
    Set objNewMember = New claValorServicioxnumero
    

    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proSeleccion = "0"
    objNewMember.proRegionCode = proRegionCode
    objNewMember.proNumero = proNumero
    objNewMember.proServicioSuplementario = proServicioSuplementarioId
    objNewMember.proValor = proValor
    objNewMember.proTipoServicio = proTipoServicio
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claValorServicioxnumero
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
    Dim varServicio As claValorServicioxnumero
    Dim i As Integer
    On Error GoTo ErrManager
    'limpiamos la coleccion
    'Limpia la coleccion
    If Me.Count > 0 Then
        For i = Me.Count To 1 Step -1
            Me.Remove (i)
        Next i
    End If
    
    varComando = "Select    N.chRegionCode,             " & _
                 "          N.vchNumero,                 " & _
                 "          N.iServicioSuplementarioId,     " & _
                 "          N.vchValor,                      " & _
                 "         S.chTipoServicio                  " & _
                 "From      CT_ValorServicioxNumero  N, CT_SERVICIOSSUPLEMENTARIOS S      " & _
                 "Where     N.iServicioSuplementarioId = S.iServicioSuplementarioId  " & _
                 "And       N.vchNumero = '" & Me.proNumero & "' " & _
                 "And       N.chRegionCode = '" & Me.proRegionCode & "' " & _
                 "And       N.iServicioSuplementarioId  = " & Me.proServicioSuplementarioId
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varServicio = New claValorServicioxnumero
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varServicio.proRegionCode = ""
        Else
            varServicio.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        If IsNull(varResultado.Fields("vchNumero")) Then
            varServicio.proNumero = ""
        Else
            varServicio.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        
        If IsNull(varResultado.Fields("iServicioSuplementarioId")) Then
            varServicio.proServicioSuplementario = ""
        Else
            varServicio.proServicioSuplementario = Trim(varResultado.Fields("iServicioSuplementarioId"))
        End If
        
        If IsNull(varResultado.Fields("vchValor")) Then
            varServicio.proValor = ""
        Else
            varServicio.proValor = Trim(varResultado.Fields("vchValor"))
        End If
        
        If IsNull(varResultado.Fields("chTipoServicio")) Then
            varServicio.proTipoServicio = ""
        Else
            varServicio.proTipoServicio = Trim(varResultado.Fields("chTipoServicio"))
        End If

        
        Me.Add Me.proConexion, _
               varServicio.proServicioSuplementario, _
               varServicio.proRegionCode, _
               varServicio.proNumero, _
               varServicio.proValor, _
               varServicio.proTipoServicio
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varServicio = Nothing
    
    MetConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function




