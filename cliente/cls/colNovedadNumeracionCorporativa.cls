VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colNovedadNumeracionCorporativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claNovedadNumeracionCorporativa"
Attribute VB_Ext_KEY = "Member0" ,"claNovedadNumeracionCorporativa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proDatosProductoId As String
Public proIncidentId As String

Public proSeleccionados As Integer

Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proMarcacion As String, _
                    proTipoNovedadId As Long, proIncidentId As Long, proDatosProductoId As Long, _
                    proVirtual As String, Optional sKey As String) As claNovedadNumeracionCorporativa
    'create a new object
    Dim objNewMember As claNovedadNumeracionCorporativa
    Set objNewMember = New claNovedadNumeracionCorporativa


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proSeleccion = "0"
    objNewMember.proMarcacion = proMarcacion
    objNewMember.proTipoNovedadId = proTipoNovedadId
    objNewMember.proIncidentId = proIncidentId
    objNewMember.proDatosProductoId = proDatosProductoId
    objNewMember.proVirtual = proVirtual 'Agregado por Carlos Castelblanco 2006/07/26
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claNovedadNumeracionCorporativa
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

Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varNovedadNumeracionCorporativa As claNovedadNumeracionCorporativa
    
    On Error GoTo ErrManager
    
    varComando = "Select    iDatosProductoId,               " & _
                 "          iIncidentId,                    " & _
                 "          iTipoNovedadId,                 " & _
                 "          vchMarcacion,                   " & _
                 "          chVirtual                       " & _
                 "From      CT_NovedadNumeracionCorporativa " & _
                 "Where     iDatosProductoId = " & Me.proDatosProductoId & " " & _
                 "And       iIncidentId      = " & Me.proIncidentId
                 'chVirtual Agregado por Carlos Castelblanco 2006/07/26
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varNovedadNumeracionCorporativa = New claNovedadNumeracionCorporativa
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varNovedadNumeracionCorporativa.proDatosProductoId = ""
        Else
            varNovedadNumeracionCorporativa.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("iIncidentId")) Then
            varNovedadNumeracionCorporativa.proIncidentId = ""
        Else
            varNovedadNumeracionCorporativa.proIncidentId = Trim(varResultado.Fields("iIncidentId"))
        End If
        
        If IsNull(varResultado.Fields("iTipoNovedadId")) Then
            varNovedadNumeracionCorporativa.proTipoNovedadId = ""
        Else
            varNovedadNumeracionCorporativa.proTipoNovedadId = Trim(varResultado.Fields("iTipoNovedadId"))
        End If
        
        If IsNull(varResultado.Fields("vchMarcacion")) Then
            varNovedadNumeracionCorporativa.proMarcacion = ""
        Else
            varNovedadNumeracionCorporativa.proMarcacion = Trim(varResultado.Fields("vchMarcacion"))
        End If
        
        'Agregado por Carlos Castelblanco 2006/07/26:
        If IsNull(varResultado.Fields("chVirtual")) Then
            varNovedadNumeracionCorporativa.proVirtual = ""
        Else
            varNovedadNumeracionCorporativa.proVirtual = Trim(varResultado.Fields("chVirtual"))
        End If
        
        
        Me.Add Me.proConexion, _
               varNovedadNumeracionCorporativa.proMarcacion, _
               varNovedadNumeracionCorporativa.proTipoNovedadId, _
               varNovedadNumeracionCorporativa.proIncidentId, _
               varNovedadNumeracionCorporativa.proDatosProductoId, _
               varNovedadNumeracionCorporativa.proVirtual
               
        varResultado.MoveNext
    Wend
    
    Set varNovedadNumeracionCorporativa = Nothing
    Set varResultado = Nothing
    
    MetConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
