VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colNovedadDetalleDatosProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claNovedadDetalleDatosProducto"
Attribute VB_Ext_KEY = "Member0" ,"claNovedadDetalleDatosProducto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proSeleccionados As Integer

'Parametros para la modificacion de una columna completa
Public proDatosProductoId As String
Public proIncidentId As String
Public proCodigos As String
Public proProductNumber As String
Public proCampo As String
Public proValor As String
Public proTabla As String   ' 0:    CT_NovedadDetalleDatosProducto
                            ' 1:    CT_DetalleDatosProducto

Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proRecordStatus As String, _
                    proUser40 As String, proUser39 As String, proUser38 As String, _
                    proUser37 As String, proUser36 As String, proUser35 As String, _
                    proUser34 As String, proUser33 As String, proUser32 As String, _
                    proUser31 As String, proUser30 As String, proUser29 As String, _
                    proUser28 As String, proUser27 As String, proUser26 As String, _
                    proUser25 As String, proUser24 As String, proUser23 As String, _
                    proUser22 As String, proUser21 As String, proUser20 As String, _
                    proUser19 As String, proUser18 As String, proUser17 As String, _
                    proUser16 As String, proUser15 As String, proUser14 As String, _
                    proUser13 As String, proUser12 As String, proUser11 As String, _
                    proUser10 As String, proUser9 As String, proUser8 As String, _
                    proUser7 As String, proUser6 As String, proUser5 As String, _
                    proUser4 As String, proUser3 As String, proUser2 As String, _
                    proUser1 As String, proStatusId As String, proDetalleDatosProductoId As String, _
                    proIncidentId As String, proDatosProductoId As String, proTipoNovedadId As String, _
                    proNovedadDetalleDatosProductoId As String, proContadorNumeros As Integer, _
                    Optional sKey As String) As claNovedadDetalleDatosProducto
                    
    'create a new object
    Dim objNewMember As claNovedadDetalleDatosProducto
    Set objNewMember = New claNovedadDetalleDatosProducto


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
    objNewMember.proIncidentId = proIncidentId
    objNewMember.proDatosProductoId = proDatosProductoId
    objNewMember.proTipoNovedadId = proTipoNovedadId
    objNewMember.proNovedadDetalleDatosProductoId = proNovedadDetalleDatosProductoId
    objNewMember.proContadorNumeros = proContadorNumeros
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claNovedadDetalleDatosProducto
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

Public Function MetConsultarxIncidente(parDetalleDatosProducto As colDetalleDatosProducto) As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
    Dim proParametro As claParametro
    Set proParametro = New claParametro
    Set proParametro.proConexion = Me.proConexion
    proParametro.proAcronimo = "ValorSi"
    proParametro.FunGConsultar
    
    On Error GoTo ErrManager
    
    varComando = "Select    D.iNovedadDetalleDatosProductoId, D.iTipoNovedadId, " & _
                 "          D.iDatosProductoId,               D.iIncidentId,    " & _
                 "          D.iDetalleDatosProductoId,        D.chStatusId,     " & _
                 "          D.vchUser1,       D.vchUser2,       D.vchUser3,       " & _
                 "          D.vchUser4,       D.vchUser5,       D.vchUser6,       " & _
                 "          D.vchUser7,       D.vchUser8,       D.vchUser9,       " & _
                 "          D.vchUser10,      D.vchUser11,      D.vchUser12,      " & _
                 "          D.vchUser13,      D.vchUser14,      D.vchUser15,      " & _
                 "          D.vchUser16,      D.vchUser17,      D.vchUser18,      " & _
                 "          D.vchUser19,      D.vchUser20,      D.vchUser21,      " & _
                 "          D.vchUser22,      D.vchUser23,      D.vchUser24,      " & _
                 "          D.vchUser25,      D.vchUser26,      D.vchUser27,      " & _
                 "          D.vchUser28,      D.vchUser29,      D.vchUser30,      " & _
                 "          D.vchUser31,      D.vchUser32,      D.vchUser33,      " & _
                 "          D.vchUser34,      D.vchUser35,      D.vchUser36,      " & _
                 "          D.vchUser37,      D.vchUser38,      D.vchUser39,      " & _
                 "          D.vchUser40,      D.tiRecordStatus,                  " & _
                 "          IsNull(Count(N.iNovedadNumeroId),0) ContadorNumeros " & _
                 "From      CT_NovedadDetalleDatosProducto D                 " & _
                 "Left Outer Join CT_NOVEDADNUMEROS N " & _
                 "On D.iNovedadDetalleDatosProductoId = N.iNovedadDetalleDatosProductoId " & _
                 "Where     D.iDatosProductoId = " & Me.proDatosProductoId & " " & _
                 "And       D.iIncidentId      = " & Me.proIncidentId & " " & _
                 "And       D.tiRecordStatus   = 1 "
        varComando = varComando & "Group By D.iNovedadDetalleDatosProductoId, D.iTipoNovedadId, " & _
                 "          D.iDatosProductoId,               D.iIncidentId,    " & _
                 "          D.iDetalleDatosProductoId,        D.chStatusId,     " & _
                 "          D.vchUser1,       D.vchUser2,       D.vchUser3,       " & _
                 "          D.vchUser4,       D.vchUser5,       D.vchUser6,       " & _
                 "          D.vchUser7,       D.vchUser8,       D.vchUser9,       " & _
                 "          D.vchUser10,      D.vchUser11,      D.vchUser12,      " & _
                 "          D.vchUser13,      D.vchUser14,      D.vchUser15,      " & _
                 "          D.vchUser16,      D.vchUser17,      D.vchUser18,      " & _
                 "          D.vchUser19,      D.vchUser20,      D.vchUser21,      " & _
                 "          D.vchUser22,      D.vchUser23,      D.vchUser24,      " & _
                 "          D.vchUser25,      D.vchUser26,      D.vchUser27,      " & _
                 "          D.vchUser28,      D.vchUser29,      D.vchUser30,      " & _
                 "          D.vchUser31,      D.vchUser32,      D.vchUser33,      " & _
                 "          D.vchUser34,      D.vchUser35,      D.vchUser36,      " & _
                 "          D.vchUser37,      D.vchUser38,      D.vchUser39,      " & _
                 "          D.vchUser40,      D.tiRecordStatus                  "
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varNovedadDetalleDatosProducto = New claNovedadDetalleDatosProducto
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iNovedadDetalleDatosProductoId")) Then
            varNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId = ""
        Else
            varNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId = Trim(varResultado.Fields("iNovedadDetalleDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("iTipoNovedadId")) Then
            varNovedadDetalleDatosProducto.proTipoNovedadId = ""
        Else
            varNovedadDetalleDatosProducto.proTipoNovedadId = Trim(varResultado.Fields("iTipoNovedadId"))
        End If
        
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varNovedadDetalleDatosProducto.proDatosProductoId = ""
        Else
            varNovedadDetalleDatosProducto.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("iIncidentId")) Then
            varNovedadDetalleDatosProducto.proIncidentId = ""
        Else
            varNovedadDetalleDatosProducto.proIncidentId = Trim(varResultado.Fields("iIncidentId"))
        End If
        
        If IsNull(varResultado.Fields("iDetalleDatosProductoId")) Then
            varNovedadDetalleDatosProducto.proDetalleDatosProductoId = ""
        Else
            varNovedadDetalleDatosProducto.proDetalleDatosProductoId = Trim(varResultado.Fields("iDetalleDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("chStatusId")) Then
            varNovedadDetalleDatosProducto.proStatusId = ""
        Else
            varNovedadDetalleDatosProducto.proStatusId = Trim(varResultado.Fields("chStatusId"))
        End If
        
        If IsNull(varResultado.Fields("vchUser1")) Then
            varNovedadDetalleDatosProducto.proUser1 = ""
        Else
            varNovedadDetalleDatosProducto.proUser1 = Trim(varResultado.Fields("vchUser1"))
        End If
        
        If IsNull(varResultado.Fields("vchUser2")) Then
            varNovedadDetalleDatosProducto.proUser2 = ""
        Else
            varNovedadDetalleDatosProducto.proUser2 = Trim(varResultado.Fields("vchUser2"))
        End If
        
        If IsNull(varResultado.Fields("vchUser3")) Then
            varNovedadDetalleDatosProducto.proUser3 = ""
        Else
            varNovedadDetalleDatosProducto.proUser3 = Trim(varResultado.Fields("vchUser3"))
        End If
        
        If IsNull(varResultado.Fields("vchUser4")) Then
            varNovedadDetalleDatosProducto.proUser4 = ""
        Else
            varNovedadDetalleDatosProducto.proUser4 = Trim(varResultado.Fields("vchUser4"))
        End If
        
        If IsNull(varResultado.Fields("vchUser5")) Then
            varNovedadDetalleDatosProducto.proUser5 = ""
        Else
            varNovedadDetalleDatosProducto.proUser5 = Trim(varResultado.Fields("vchUser5"))
        End If
        
        If IsNull(varResultado.Fields("vchUser6")) Then
            varNovedadDetalleDatosProducto.proUser6 = ""
        Else
            varNovedadDetalleDatosProducto.proUser6 = Trim(varResultado.Fields("vchUser6"))
        End If
        
        If IsNull(varResultado.Fields("vchUser7")) Then
            varNovedadDetalleDatosProducto.proUser7 = ""
        Else
            varNovedadDetalleDatosProducto.proUser7 = Trim(varResultado.Fields("vchUser7"))
        End If
        
        If IsNull(varResultado.Fields("vchUser8")) Then
            varNovedadDetalleDatosProducto.proUser8 = ""
        Else
            varNovedadDetalleDatosProducto.proUser8 = Trim(varResultado.Fields("vchUser8"))
        End If
        
        If IsNull(varResultado.Fields("vchUser9")) Then
            varNovedadDetalleDatosProducto.proUser9 = ""
        Else
            varNovedadDetalleDatosProducto.proUser9 = Trim(varResultado.Fields("vchUser9"))
        End If
        
        If IsNull(varResultado.Fields("vchUser10")) Then
            varNovedadDetalleDatosProducto.proUser10 = ""
        Else
            varNovedadDetalleDatosProducto.proUser10 = Trim(varResultado.Fields("vchUser10"))
        End If
        
        If IsNull(varResultado.Fields("vchUser11")) Then
            varNovedadDetalleDatosProducto.proUser11 = ""
        Else
            varNovedadDetalleDatosProducto.proUser11 = Trim(varResultado.Fields("vchUser11"))
        End If
        
        If IsNull(varResultado.Fields("vchUser12")) Then
            varNovedadDetalleDatosProducto.proUser12 = ""
        Else
            varNovedadDetalleDatosProducto.proUser12 = Trim(varResultado.Fields("vchUser12"))
        End If
        
        If IsNull(varResultado.Fields("vchUser13")) Then
            varNovedadDetalleDatosProducto.proUser13 = ""
        Else
            varNovedadDetalleDatosProducto.proUser13 = Trim(varResultado.Fields("vchUser13"))
        End If
        
        If IsNull(varResultado.Fields("vchUser14")) Then
            varNovedadDetalleDatosProducto.proUser14 = ""
        Else
            varNovedadDetalleDatosProducto.proUser14 = Trim(varResultado.Fields("vchUser14"))
        End If
        
        If IsNull(varResultado.Fields("vchUser15")) Then
            varNovedadDetalleDatosProducto.proUser15 = ""
        Else
            varNovedadDetalleDatosProducto.proUser15 = Trim(varResultado.Fields("vchUser15"))
        End If
        
        If IsNull(varResultado.Fields("vchUser16")) Then
            varNovedadDetalleDatosProducto.proUser16 = ""
        Else
            varNovedadDetalleDatosProducto.proUser16 = Trim(varResultado.Fields("vchUser16"))
        End If
        
        If IsNull(varResultado.Fields("vchUser17")) Then
            varNovedadDetalleDatosProducto.proUser17 = ""
        Else
            varNovedadDetalleDatosProducto.proUser17 = Trim(varResultado.Fields("vchUser17"))
        End If
        
        If IsNull(varResultado.Fields("vchUser18")) Then
            varNovedadDetalleDatosProducto.proUser18 = ""
        Else
            varNovedadDetalleDatosProducto.proUser18 = Trim(varResultado.Fields("vchUser18"))
        End If
        
        If IsNull(varResultado.Fields("vchUser19")) Then
            varNovedadDetalleDatosProducto.proUser19 = ""
        Else
            varNovedadDetalleDatosProducto.proUser19 = Trim(varResultado.Fields("vchUser19"))
        End If
        
        If IsNull(varResultado.Fields("vchUser20")) Then
            varNovedadDetalleDatosProducto.proUser20 = ""
        Else
            varNovedadDetalleDatosProducto.proUser20 = Trim(varResultado.Fields("vchUser20"))
        End If
        
        If IsNull(varResultado.Fields("vchUser21")) Then
            varNovedadDetalleDatosProducto.proUser21 = ""
        Else
            varNovedadDetalleDatosProducto.proUser21 = Trim(varResultado.Fields("vchUser21"))
        End If
        
        If IsNull(varResultado.Fields("vchUser22")) Then
            varNovedadDetalleDatosProducto.proUser22 = ""
        Else
            varNovedadDetalleDatosProducto.proUser22 = Trim(varResultado.Fields("vchUser22"))
        End If
        
        If IsNull(varResultado.Fields("vchUser23")) Then
            varNovedadDetalleDatosProducto.proUser23 = ""
        Else
            varNovedadDetalleDatosProducto.proUser23 = Trim(varResultado.Fields("vchUser23"))
        End If
        
        If IsNull(varResultado.Fields("vchUser24")) Then
            varNovedadDetalleDatosProducto.proUser24 = ""
        Else
            varNovedadDetalleDatosProducto.proUser24 = Trim(varResultado.Fields("vchUser24"))
        End If
        
        If IsNull(varResultado.Fields("vchUser25")) Then
            varNovedadDetalleDatosProducto.proUser25 = ""
        Else
            varNovedadDetalleDatosProducto.proUser25 = Trim(varResultado.Fields("vchUser25"))
        End If
        
        If IsNull(varResultado.Fields("vchUser26")) Then
            varNovedadDetalleDatosProducto.proUser26 = ""
        Else
            varNovedadDetalleDatosProducto.proUser26 = Trim(varResultado.Fields("vchUser26"))
        End If
        
        If IsNull(varResultado.Fields("vchUser27")) Then
            varNovedadDetalleDatosProducto.proUser27 = ""
        Else
            varNovedadDetalleDatosProducto.proUser27 = Trim(varResultado.Fields("vchUser27"))
        End If
        
        If IsNull(varResultado.Fields("vchUser28")) Then
            varNovedadDetalleDatosProducto.proUser28 = ""
        Else
            varNovedadDetalleDatosProducto.proUser28 = Trim(varResultado.Fields("vchUser28"))
        End If
        
        If IsNull(varResultado.Fields("vchUser29")) Then
            varNovedadDetalleDatosProducto.proUser29 = ""
        Else
            varNovedadDetalleDatosProducto.proUser29 = Trim(varResultado.Fields("vchUser29"))
        End If
        
        If IsNull(varResultado.Fields("vchUser30")) Then
            varNovedadDetalleDatosProducto.proUser30 = ""
        Else
            varNovedadDetalleDatosProducto.proUser30 = Trim(varResultado.Fields("vchUser30"))
        End If
        
        If IsNull(varResultado.Fields("vchUser31")) Then
            varNovedadDetalleDatosProducto.proUser31 = ""
        Else
            varNovedadDetalleDatosProducto.proUser31 = Trim(varResultado.Fields("vchUser31"))
        End If
        
        If IsNull(varResultado.Fields("vchUser32")) Then
            varNovedadDetalleDatosProducto.proUser32 = ""
        Else
            varNovedadDetalleDatosProducto.proUser32 = Trim(varResultado.Fields("vchUser32"))
        End If
        
        If IsNull(varResultado.Fields("vchUser33")) Then
            varNovedadDetalleDatosProducto.proUser33 = ""
        Else
            varNovedadDetalleDatosProducto.proUser33 = Trim(varResultado.Fields("vchUser33"))
        End If
        
        If IsNull(varResultado.Fields("vchUser34")) Then
            varNovedadDetalleDatosProducto.proUser34 = ""
        Else
            varNovedadDetalleDatosProducto.proUser34 = Trim(varResultado.Fields("vchUser34"))
        End If
        
        If IsNull(varResultado.Fields("vchUser35")) Then
            varNovedadDetalleDatosProducto.proUser35 = ""
        Else
            varNovedadDetalleDatosProducto.proUser35 = Trim(varResultado.Fields("vchUser35"))
        End If
        
        If IsNull(varResultado.Fields("vchUser36")) Then
            varNovedadDetalleDatosProducto.proUser36 = ""
        Else
            varNovedadDetalleDatosProducto.proUser36 = Trim(varResultado.Fields("vchUser36"))
        End If
        
        If IsNull(varResultado.Fields("vchUser37")) Then
            varNovedadDetalleDatosProducto.proUser37 = ""
        Else
            varNovedadDetalleDatosProducto.proUser37 = Trim(varResultado.Fields("vchUser37"))
        End If
        
        If IsNull(varResultado.Fields("vchUser38")) Then
            varNovedadDetalleDatosProducto.proUser38 = ""
        Else
            varNovedadDetalleDatosProducto.proUser38 = Trim(varResultado.Fields("vchUser38"))
        End If
        
        If IsNull(varResultado.Fields("vchUser39")) Then
            varNovedadDetalleDatosProducto.proUser39 = ""
        Else
            varNovedadDetalleDatosProducto.proUser39 = Trim(varResultado.Fields("vchUser39"))
        End If
        
        If IsNull(varResultado.Fields("vchUser40")) Then
            varNovedadDetalleDatosProducto.proUser40 = ""
        Else
            varNovedadDetalleDatosProducto.proUser40 = Trim(varResultado.Fields("vchUser40"))
        End If
        
        If IsNull(varResultado.Fields("tiRecordStatus")) Then
            varNovedadDetalleDatosProducto.proRecordStatus = ""
        Else
            varNovedadDetalleDatosProducto.proRecordStatus = Trim(varResultado.Fields("tiRecordStatus"))
        End If
        
        varNovedadDetalleDatosProducto.proContadorNumeros = varResultado!ContadorNumeros
        
        With varNovedadDetalleDatosProducto
            Me.Add Me.proConexion, .proRecordStatus, _
                   .proUser40, .proUser39, .proUser38, .proUser37, _
                   .proUser36, .proUser35, .proUser34, .proUser33, _
                   .proUser32, .proUser31, .proUser30, .proUser29, _
                   .proUser28, .proUser27, .proUser26, .proUser25, _
                   .proUser24, .proUser23, .proUser22, .proUser21, _
                   .proUser20, .proUser19, .proUser18, .proUser17, _
                   .proUser16, .proUser15, .proUser14, .proUser13, _
                   .proUser12, .proUser11, .proUser10, .proUser9, _
                   .proUser8, .proUser7, .proUser6, .proUser5, _
                   .proUser4, .proUser3, .proUser2, .proUser1, _
                   .proStatusId, .proDetalleDatosProductoId, .proIncidentId, _
                   .proDatosProductoId, .proTipoNovedadId, _
                   .proNovedadDetalleDatosProductoId, .proContadorNumeros
            Dim varIndice As Integer
            If .proTipoNovedadId = "3" Then
                varIndice = parDetalleDatosProducto.IndexOf(.proDetalleDatosProductoId)
                parDetalleDatosProducto(varIndice).proEliminar = True
            End If
            If .proTipoNovedadId = "2" And varNovedadDetalleDatosProducto.proUser15 = proParametro.proValor Then
                varIndice = parDetalleDatosProducto.IndexOf(.proDetalleDatosProductoId)
                parDetalleDatosProducto(varIndice).proBackUp = True
            End If
            If .proTipoNovedadId = "2" Then
                varIndice = parDetalleDatosProducto.IndexOf(.proDetalleDatosProductoId)
                parDetalleDatosProducto(varIndice).proModificar = True
            End If
        End With
        
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varNovedadDetalleDatosProducto = Nothing
    
    MetConsultarxIncidente = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetActualizarColumna() As Boolean
    Dim varComando As ADODB.Command
    On Error GoTo ErrManager
    
    Set varComando = New ADODB.Command
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandType = adCmdStoredProc
    varComando.CommandTimeout = 30
    varComando.CommandText = "CP_ActualizarNovedadDetalleDatosProducto"
    
    varComando.Parameters.Item("@vchCodigos").Value = Me.proCodigos
    varComando.Parameters.Item("@chProductNumber").Value = Val(Me.proProductNumber)
    varComando.Parameters.Item("@chCampoPadre").Value = Me.proCampo
    varComando.Parameters.Item("@vchValor").Value = Me.proValor
    varComando.Parameters.Item("@vchTabla").Value = Me.proTabla
    varComando.Parameters.Item("@iDatosProductoId").Value = Me.proDatosProductoId
    varComando.Parameters.Item("@iIncidentId").Value = Me.proIncidentId
    
    varComando.Execute
    
    MetActualizarColumna = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Property Get IndexOf(vntValue As Long) As Long
    Dim i As Integer
    Dim j As Integer
    j = 0
    i = 1
    While i <= mCol.Count And j = 0
        Dim vardetalle As claNovedadDetalleDatosProducto
        Set vardetalle = mCol(i)
        If vardetalle.proNovedadDetalleDatosProductoId = vntValue Then
            j = i
        End If
        i = i + 1
    Wend
    IndexOf = j
End Property



