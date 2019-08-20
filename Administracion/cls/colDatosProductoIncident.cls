VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colDatosProductoIncident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claDatosProductoIncident"
Attribute VB_Ext_KEY = "Member0" ,"claDatosProductoIncident"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proDatosProductoId As String

Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proDescripcion As String, _
                    proTipo As String, proCategoria As String, proFechaModificacion As String, _
                    proIncidentId As String, proDatosProductoId As String, _
                    Optional sKey As String) As claDatosProductoIncident
    'create a new object
    Dim objNewMember As claDatosProductoIncident
    Set objNewMember = New claDatosProductoIncident


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proDescripcion = proDescripcion
    objNewMember.proTipo = proTipo
    objNewMember.proCategoria = proCategoria
    objNewMember.proFechaModificacion = proFechaModificacion
    objNewMember.proIncidentId = proIncidentId
    objNewMember.proDatosProductoId = proDatosProductoId
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claDatosProductoIncident
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

Public Function MetConsultarxDatosProducto() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varDatosProductoIncident As claDatosProductoIncident
    
    On Error GoTo ErrManager
    
    varComando = "Select    DI.iDatosProductoId,                " & _
                 "          DI.iIncidentId,                     " & _
                 "          DI.dtFechaModificacion,             " & _
                 "          I.vchDesc1,                         " & _
                 "          R1.vchParameterDesc,                " & _
                 "          R2.vchParameterDesc                 " & _
                 "From      Incident                    I,      " & _
                 "          CT_DatosProductoIncident    DI,     " & _
                 "          Reference_Parameters        R1,     " & _
                 "          Reference_Parameters        R2      " & _
                 "Where     DI.iDatosProductoId = " & Me.proDatosProductoId & " " & _
                 "and       I.iIncidentTypeId = R1.iParameterId " & _
                 "and       I.iIncidentCategory = R2.iParameterId " & _
                 "and       I.iIncidentId = DI.iIncidentId      " & _
                 "Order by  DI.dtFechaModificacion Desc         "
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varDatosProductoIncident = New claDatosProductoIncident
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varDatosProductoIncident.proDatosProductoId = 0
        Else
            varDatosProductoIncident.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("iIncidentId")) Then
            varDatosProductoIncident.proIncidentId = 0
        Else
            varDatosProductoIncident.proIncidentId = Trim(varResultado.Fields("iIncidentId"))
        End If
        
        If IsNull(varResultado.Fields("dtFechaModificacion")) Then
            varDatosProductoIncident.proFechaModificacion = 0
        Else
            varDatosProductoIncident.proFechaModificacion = Trim(varResultado.Fields("dtFechaModificacion"))
        End If
        
        If IsNull(varResultado.Fields("vchDesc1")) Then
            varDatosProductoIncident.proDescripcion = 0
        Else
            varDatosProductoIncident.proDescripcion = Trim(varResultado.Fields("vchDesc1"))
        End If
        
        If IsNull(varResultado.Fields(4)) Then
            varDatosProductoIncident.proTipo = 0
        Else
            varDatosProductoIncident.proTipo = Trim(varResultado.Fields(4))
        End If
        
        If IsNull(varResultado.Fields(5)) Then
            varDatosProductoIncident.proCategoria = 0
        Else
            varDatosProductoIncident.proCategoria = Trim(varResultado.Fields(5))
        End If
        
        Me.Add Me.proConexion, _
                varDatosProductoIncident.proDescripcion, _
                varDatosProductoIncident.proTipo, _
                varDatosProductoIncident.proCategoria, _
                varDatosProductoIncident.proFechaModificacion, _
                varDatosProductoIncident.proIncidentId, _
                varDatosProductoIncident.proDatosProductoId
        
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varDatosProductoIncident = Nothing
    
    MetConsultarxDatosProducto = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
