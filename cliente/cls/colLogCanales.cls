VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colLogCanales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claClienteTelefonia"
Attribute VB_Ext_KEY = "Member0" ,"claClienteTelefonia"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************
'       DESCRIPCION: Clase que mantiene una coleccion de claLogCanales
'       y permite la busqueda en CT_LOG_CANALES
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Desarrollo inicial de la clase
'       VERSION:              1.0.000
'       REQUERIMIENTO:        3488
'       FECHA:                2009/07/30
'*******************************************************************
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public proConexion As ADODB.Connection

Public proCompanyId As String
Public proFechaDesde As String
Public proFechaHasta As String
Public proIncidentId As String
Public proSerialNumber As String
Public proTipoAsunto As String




Public Function Add(proConexion As ADODB.Connection, proIncidentId As String, proCompanyId As String, _
                    proIncidentCategory As String, proTipoLinea As String, proSerialNumber As String, _
                    proCanalesEnUso As Integer, proDiferenciaCanales As Integer, proCanalesCalculados As Integer, _
                    proUserId As String, proFechaNovedad As String, Optional sKey As String) As claLogCanales
    'create a new object
    Dim objNewMember As claLogCanales
    Set objNewMember = New claLogCanales


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proIncidentId = proIncidentId
    objNewMember.proCompanyId = proCompanyId
    objNewMember.proIncidentCategory = proIncidentCategory
    objNewMember.proTipoLinea = proTipoLinea
    objNewMember.proSerialNumber = proSerialNumber
    objNewMember.proCanalesEnUso = proCanalesEnUso
    objNewMember.proDiferenciaCanales = proDiferenciaCanales
    objNewMember.proCanalesCalculados = proCanalesCalculados
    objNewMember.proUserId = proUserId
    objNewMember.proFechaNovedad = proFechaNovedad
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Function MetConsultarLogCanales() As Boolean
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    Dim varLogCanales As claLogCanales
    
    On Error GoTo ErrManager
    
    Set varComando = New ADODB.Command
    
    varComando.CommandType = adCmdStoredProc
    varComando.ActiveConnection = Me.proConexion
    varComando.CommandText = "CRMCP_ConsultaLogCanales"
    varComando.CommandTimeout = 0
    
    varComando.Parameters.Item("@iCompanyId").Value = Me.proCompanyId
    varComando.Parameters.Item("@vchFechaDesde").Value = IIf(Me.proFechaDesde = "", Null, Me.proFechaDesde)
    varComando.Parameters.Item("@vchFechaHasta").Value = IIf(Me.proFechaHasta = "", Null, Me.proFechaHasta)
    varComando.Parameters.Item("@iIncidentId").Value = IIf(Me.proIncidentId = "", Null, Val(Me.proIncidentId))
    varComando.Parameters.Item("@vchSerialNumber").Value = IIf(Me.proSerialNumber = "", Null, Me.proSerialNumber)
    varComando.Parameters.Item("@iIncidentCategory").Value = IIf(Me.proTipoAsunto = "", Null, Val(Me.proTipoAsunto))
   
    Set varResultado = New ADODB.Recordset
    Set varResultado = varComando.Execute
    
    Set varLogCanales = New claLogCanales

    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iIncidentId")) Then
            varLogCanales.proIncidentId = ""
        Else
            varLogCanales.proIncidentId = varResultado.Fields("iIncidentId")
        End If
        
        If IsNull(varResultado.Fields("iCompanyId")) Then
            varLogCanales.proCompanyId = ""
        Else
            varLogCanales.proCompanyId = Trim(varResultado.Fields("iCompanyId"))
        End If
        
        If IsNull(varResultado.Fields("vchCategory")) Then
            varLogCanales.proIncidentCategory = ""
        Else
            varLogCanales.proIncidentCategory = Trim(varResultado.Fields("vchCategory"))
        End If
        
        If IsNull(varResultado.Fields("vchTipoLinea")) Then
            varLogCanales.proTipoLinea = ""
        Else
            varLogCanales.proTipoLinea = Trim(varResultado.Fields("vchTipoLinea"))
        End If
        
        If IsNull(varResultado.Fields("vchSerialNumber")) Then
            varLogCanales.proSerialNumber = ""
        Else
            varLogCanales.proSerialNumber = Trim(varResultado("vchSerialNumber"))
        End If
        
        If IsNull(varResultado.Fields("dtFechaNovedad")) Then
            varLogCanales.proFechaNovedad = ""
        Else
            varLogCanales.proFechaNovedad = Trim(varResultado.Fields("dtFechaNovedad"))
        End If
        
        If IsNull(varResultado.Fields("iCanalesEnUso")) Then
            varLogCanales.proCanalesEnUso = ""
        Else
            varLogCanales.proCanalesEnUso = varResultado.Fields("iCanalesEnUso")
        End If
        
        If IsNull(varResultado.Fields("iDiferenciaCanales")) Then
            varLogCanales.proDiferenciaCanales = ""
        Else
            varLogCanales.proDiferenciaCanales = varResultado("iDiferenciaCanales")
        End If
        
        If IsNull(varResultado.Fields("iCanalesCalculados")) Then
            varLogCanales.proCanalesCalculados = ""
        Else
            varLogCanales.proCanalesCalculados = varResultado("iCanalesCalculados")
        End If
        
        If IsNull(varResultado.Fields("chaUserId")) Then
            varLogCanales.proUserId = ""
        Else
            varLogCanales.proUserId = Trim(varResultado("chaUserId"))
        End If
        
        
        Me.Add Me.proConexion, _
               varLogCanales.proIncidentId, _
               varLogCanales.proCompanyId, _
               varLogCanales.proIncidentCategory, _
               varLogCanales.proTipoLinea, _
               varLogCanales.proSerialNumber, _
               varLogCanales.proCanalesEnUso, _
               varLogCanales.proDiferenciaCanales, _
               varLogCanales.proCanalesCalculados, _
               varLogCanales.proUserId, _
               varLogCanales.proFechaNovedad
              
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varLogCanales = Nothing
    
    MetConsultarLogCanales = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function




Public Property Get Item(vntIndexKey As Variant) As claLogCanales
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





