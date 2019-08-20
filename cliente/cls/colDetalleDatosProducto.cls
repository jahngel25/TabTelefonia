VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colDetalleDatosProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claDetalleDatosProducto"
Attribute VB_Ext_KEY = "Member0" ,"claDetalleDatosProducto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'************************************************************************************************
'   DESCRIPCION         : Coleccion que maneja objetos de tipo detalledatos producto que para el
'                         tab de telefonia son los tipos de linea
'   PARAMETROS          :
'                        N/A
'
'   RETORNO             : NA
'
'   EJEMPLO             :

'*************************************************************************************************
'   MODIFICADO POR      : Carlos Leonardo Villamil (I&T)
'   DESCRIPCION CAMBIO  : correccion error en la definicion de datos
'   VERCION             : 3.7.402
'   FECHA               : 16-SEP-2009
'*************************************************************************************************
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proDatosProductoId As String
Public proConexion As ADODB.Connection
Public proSeleccionados As Integer

Public Function Add(proConexion As ADODB.Connection, proRecordStatus As String, _
                    proUser40 As String, proUser39 As String, proUser38 As String, proUser37 As String, _
                    proUser36 As String, proUser35 As String, proUser34 As String, proUser33 As String, _
                    proUser32 As String, proUser31 As String, proUser30 As String, proUser29 As String, _
                    proUser28 As String, proUser27 As String, proUser26 As String, proUser25 As String, _
                    proUser24 As String, proUser23 As String, proUser22 As String, proUser21 As String, _
                    proUser20 As String, proUser19 As String, proUser18 As String, proUser17 As String, _
                    proUser16 As String, proUser15 As String, proUser14 As String, proUser13 As String, _
                    proUser12 As String, proUser11 As String, proUser10 As String, proUser9 As String, _
                    proUser8 As String, proUser7 As String, proUser6 As String, proUser5 As String, _
                    proUser4 As String, proUser3 As String, proUser2 As String, proUser1 As String, _
                    proStatusId As String, proDetalleDatosProductoId As String, proDatosProductoId As String, _
                    Optional sKey As String) As claDetalleDatosProducto
                    
    'create a new object
    Dim objNewMember As claDetalleDatosProducto
    Set objNewMember = New claDetalleDatosProducto


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proSeleccion = "0"
    objNewMember.proRecordStatus = proRecordStatus
    objNewMember.proUser40 = proUser40
    objNewMember.proUser39 = proUser39
    objNewMember.proUser38 = proUser38
    objNewMember.proUser37 = proUser37
    objNewMember.proUser36 = proUser36
    objNewMember.proUser35 = proUser35
    objNewMember.proUser34 = proUser34
    objNewMember.proUser33 = proUser33
    objNewMember.proUser32 = proUser32
    objNewMember.proUser31 = proUser31
    objNewMember.proUser30 = proUser30
    objNewMember.proUser29 = proUser29
    objNewMember.proUser28 = proUser28
    objNewMember.proUser27 = proUser27
    objNewMember.proUser26 = proUser26
    objNewMember.proUser25 = proUser25
    objNewMember.proUser24 = proUser24
    objNewMember.proUser23 = proUser23
    objNewMember.proUser22 = proUser22
    objNewMember.proUser21 = proUser21
    objNewMember.proUser20 = proUser20
    objNewMember.proUser19 = proUser19
    objNewMember.proUser18 = proUser18
    objNewMember.proUser17 = proUser17
    objNewMember.proUser16 = proUser16
    objNewMember.proUser15 = proUser15
    objNewMember.proUser14 = proUser14
    objNewMember.proUser13 = proUser13
    objNewMember.proUser12 = proUser12
    objNewMember.proUser11 = proUser11
    objNewMember.proUser10 = proUser10
    objNewMember.proUser9 = proUser9
    objNewMember.proUser8 = proUser8
    objNewMember.proUser7 = proUser7
    objNewMember.proUser6 = proUser6
    objNewMember.proUser5 = proUser5
    objNewMember.proUser4 = proUser4
    objNewMember.proUser3 = proUser3
    objNewMember.proUser2 = proUser2
    objNewMember.proUser1 = proUser1
    objNewMember.proStatusId = proStatusId
    objNewMember.proDetalleDatosProductoId = proDetalleDatosProductoId
    objNewMember.proDatosProductoId = proDatosProductoId
    objNewMember.proEliminar = False
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function
'************************************************************************************************
'   DESCRIPCION         : Retorna el indice dentro de lacoleccion de un valor de idettalledatosproductoid
'   PARAMETROS          :
'                        long Idettalledatosproductoid
'
'   RETORNO             :
'                        int  indice de coleccion
'
'   EJEMPLO             :
'*************************************************************************************************
'   MODIFICADO POR      : Carlos Leonardo Villamil (I&T)
'   DESCRIPCION CAMBIO  : correccion error en la definicion de datos
'   VERCION             : 3.7.402
'   FECHA               : 16-SEP-2009
'*************************************************************************************************
Public Property Get IndexOf(vntValue As Variant) As Integer '-->3.7.402

    Dim i As Integer
    Dim j As Integer
    j = 0
    i = 1
    While i <= mCol.Count And j = 0
        Dim vardetalle As claDetalleDatosProducto
        Set vardetalle = mCol(i)
        If vardetalle.proDetalleDatosProductoId = vntValue Then
            j = i
        End If
        i = i + 1
    Wend
    IndexOf = j
End Property

Public Property Get Item(vntIndexKey As Variant) As claDetalleDatosProducto
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
    Me.proSeleccionados = 0
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Function MetConsultarxDatosProducto() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varDetalleDatosProducto As claDetalleDatosProducto
    
    On Error GoTo ErrManager
    
    varComando = "SELECT    iDatosProductoId,   iDetalleDatosProductoId,        " & _
                 "          chStatusId,   vchUser1,    vchUser2,     vchUser3,  " & _
                 "          vchUser4,     vchUser5,    vchUser6,     vchUser7,  " & _
                 "          vchUser8,     vchUser9,    vchUser10,    vchUser11, " & _
                 "          vchUser12,    vchUser13,   vchUser14,    vchUser15, " & _
                 "          vchUser16,    vchUser17,   vchUser18,    vchUser19, " & _
                 "          vchUser20,    vchUser21,   vchUser22,    vchUser23, " & _
                 "          vchUser24,    vchUser25,   vchUser26,    vchUser27, " & _
                 "          vchUser28,    vchUser29,   vchUser30,    vchUser31, " & _
                 "          vchUser32,    vchUser33,   vchUser34,    vchUser35, " & _
                 "          vchUser36,    vchUser37,   vchUser38,    vchUser39, " & _
                 "          vchUser40,    tiRecordStatus                        " & _
                 "FROM      CT_DetalleDatosProducto                             " & _
                 "WHERE     iDatosProductoId = " & Me.proDatosProductoId
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varDetalleDatosProducto = New claDetalleDatosProducto
    
    While Not varResultado.EOF
    
        varDetalleDatosProducto.proDatosProductoId = IIf(IsNull(varResultado("iDatosProductoId")), 0, varResultado.Fields("iDatosProductoId"))
        varDetalleDatosProducto.proDetalleDatosProductoId = IIf(IsNull(varResultado("iDetalleDatosProductoId")), 0, varResultado.Fields("iDetalleDatosProductoId"))
        varDetalleDatosProducto.proStatusId = IIf(IsNull(varResultado("chStatusId")), "", varResultado.Fields("chStatusId"))
        varDetalleDatosProducto.proUser1 = IIf(IsNull(varResultado("vchUser1")), "", varResultado.Fields("vchUser1"))
        varDetalleDatosProducto.proUser2 = IIf(IsNull(varResultado("vchUser2")), "", varResultado.Fields("vchUser2"))
        varDetalleDatosProducto.proUser3 = IIf(IsNull(varResultado("vchuser3")), "", varResultado.Fields("vchuser3"))
        varDetalleDatosProducto.proUser4 = IIf(IsNull(varResultado("vchUser4")), "", varResultado.Fields("vchUser4"))
        varDetalleDatosProducto.proUser5 = IIf(IsNull(varResultado("vchUser5")), "", varResultado.Fields("vchUser5"))
        varDetalleDatosProducto.proUser6 = IIf(IsNull(varResultado("vchuser6")), "", varResultado.Fields("vchuser6"))
        varDetalleDatosProducto.proUser7 = IIf(IsNull(varResultado("vchUser7")), "", varResultado.Fields("vchUser7"))
        varDetalleDatosProducto.proUser8 = IIf(IsNull(varResultado("vchUser8")), "", varResultado.Fields("vchUser8"))
        varDetalleDatosProducto.proUser9 = IIf(IsNull(varResultado("vchUser9")), "", varResultado.Fields("vchUser9"))
        varDetalleDatosProducto.proUser10 = IIf(IsNull(varResultado("vchUser10")), "", varResultado.Fields("vchUser10"))
        varDetalleDatosProducto.proUser11 = IIf(IsNull(varResultado("vchUser11")), "", varResultado.Fields("vchUser11"))
        varDetalleDatosProducto.proUser12 = IIf(IsNull(varResultado("vchUser12")), "", varResultado.Fields("vchUser12"))
        varDetalleDatosProducto.proUser13 = IIf(IsNull(varResultado("vchUser13")), "", varResultado.Fields("vchUser13"))
        varDetalleDatosProducto.proUser14 = IIf(IsNull(varResultado("vchUser14")), "", varResultado.Fields("vchUser14"))
        varDetalleDatosProducto.proUser15 = IIf(IsNull(varResultado("vchUser15")), "", varResultado.Fields("vchUser15"))
        varDetalleDatosProducto.proUser16 = IIf(IsNull(varResultado("vchUser16")), "", varResultado.Fields("vchUser16"))
        varDetalleDatosProducto.proUser17 = IIf(IsNull(varResultado("vchUser17")), "", varResultado.Fields("vchUser17"))
        varDetalleDatosProducto.proUser18 = IIf(IsNull(varResultado("vchUser18")), "", varResultado.Fields("vchUser18"))
        varDetalleDatosProducto.proUser19 = IIf(IsNull(varResultado("vchuser19")), "", varResultado.Fields("vchuser19"))
        varDetalleDatosProducto.proUser20 = IIf(IsNull(varResultado("vchUser20")), "", varResultado.Fields("vchUser20"))
        varDetalleDatosProducto.proUser21 = IIf(IsNull(varResultado("vchUser21")), "", varResultado.Fields("vchUser21"))
        varDetalleDatosProducto.proUser22 = IIf(IsNull(varResultado("vchUser22")), "", varResultado.Fields("vchUser22"))
        varDetalleDatosProducto.proUser23 = IIf(IsNull(varResultado("vchUser23")), "", varResultado.Fields("vchUser23"))
        varDetalleDatosProducto.proUser24 = IIf(IsNull(varResultado("vchUser24")), "", varResultado.Fields("vchUser24"))
        varDetalleDatosProducto.proUser25 = IIf(IsNull(varResultado("vchUser25")), "", varResultado.Fields("vchUser25"))
        varDetalleDatosProducto.proUser26 = IIf(IsNull(varResultado("vchUser26")), "", varResultado.Fields("vchUser26"))
        varDetalleDatosProducto.proUser27 = IIf(IsNull(varResultado("vchUser27")), "", varResultado.Fields("vchUser27"))
        varDetalleDatosProducto.proUser28 = IIf(IsNull(varResultado("vchUser28")), "", varResultado.Fields("vchUser28"))
        varDetalleDatosProducto.proUser29 = IIf(IsNull(varResultado("vchUser29")), "", varResultado.Fields("vchUser29"))
        varDetalleDatosProducto.proUser30 = IIf(IsNull(varResultado("vchUser30")), "", varResultado.Fields("vchUser30"))
        varDetalleDatosProducto.proUser31 = IIf(IsNull(varResultado("vchUser31")), "", varResultado.Fields("vchUser31"))
        varDetalleDatosProducto.proUser32 = IIf(IsNull(varResultado("vchUser32")), "", varResultado.Fields("vchUser32"))
        varDetalleDatosProducto.proUser33 = IIf(IsNull(varResultado("vchUser33")), "", varResultado.Fields("vchUser33"))
        varDetalleDatosProducto.proUser34 = IIf(IsNull(varResultado("vchUser34")), "", varResultado.Fields("vchUser34"))
        varDetalleDatosProducto.proUser35 = IIf(IsNull(varResultado("vchUser35")), "", varResultado.Fields("vchUser35"))
        varDetalleDatosProducto.proUser36 = IIf(IsNull(varResultado("vchUser36")), "", varResultado.Fields("vchUser36"))
        varDetalleDatosProducto.proUser37 = IIf(IsNull(varResultado("vchUser37")), "", varResultado.Fields("vchUser37"))
        varDetalleDatosProducto.proUser38 = IIf(IsNull(varResultado("vchUser38")), "", varResultado.Fields("vchUser38"))
        varDetalleDatosProducto.proUser39 = IIf(IsNull(varResultado("vchUser39")), "", varResultado.Fields("vchUser39"))
        varDetalleDatosProducto.proUser40 = IIf(IsNull(varResultado("vchUser40")), "", varResultado.Fields("vchUser40"))
        varDetalleDatosProducto.proRecordStatus = IIf(IsNull(varResultado("tiRecordStatus")), "", varResultado.Fields("tiRecordStatus"))
        Me.Add Me.proConexion, varDetalleDatosProducto.proRecordStatus, _
                varDetalleDatosProducto.proUser40, varDetalleDatosProducto.proUser39, varDetalleDatosProducto.proUser38, varDetalleDatosProducto.proUser37, _
                varDetalleDatosProducto.proUser36, varDetalleDatosProducto.proUser35, varDetalleDatosProducto.proUser34, varDetalleDatosProducto.proUser33, _
                varDetalleDatosProducto.proUser32, varDetalleDatosProducto.proUser31, varDetalleDatosProducto.proUser30, varDetalleDatosProducto.proUser29, _
                varDetalleDatosProducto.proUser28, varDetalleDatosProducto.proUser27, varDetalleDatosProducto.proUser26, varDetalleDatosProducto.proUser25, _
                varDetalleDatosProducto.proUser24, varDetalleDatosProducto.proUser23, varDetalleDatosProducto.proUser22, varDetalleDatosProducto.proUser21, _
                varDetalleDatosProducto.proUser20, varDetalleDatosProducto.proUser19, varDetalleDatosProducto.proUser18, varDetalleDatosProducto.proUser17, _
                varDetalleDatosProducto.proUser16, varDetalleDatosProducto.proUser15, varDetalleDatosProducto.proUser14, varDetalleDatosProducto.proUser13, _
                varDetalleDatosProducto.proUser12, varDetalleDatosProducto.proUser11, varDetalleDatosProducto.proUser10, varDetalleDatosProducto.proUser9, _
                varDetalleDatosProducto.proUser8, varDetalleDatosProducto.proUser7, varDetalleDatosProducto.proUser6, varDetalleDatosProducto.proUser5, _
                varDetalleDatosProducto.proUser4, varDetalleDatosProducto.proUser3, varDetalleDatosProducto.proUser2, varDetalleDatosProducto.proUser1, _
                varDetalleDatosProducto.proStatusId, varDetalleDatosProducto.proDetalleDatosProductoId, _
                varDetalleDatosProducto.proDatosProductoId
                
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varDetalleDatosProducto = Nothing
    
    MetConsultarxDatosProducto = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
