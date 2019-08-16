VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colRestriccionTabTel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"claRestriccionTabFacturacion"
Attribute VB_Ext_KEY = "Member0" ,"claRestriccionTabFacturacion"
'local variable to hold collection
Private mCol As Collection
'Conexión a la base
Public proConexion As ADODB.Connection
Function FunGValidarRestriccionesTab(parAsunto As String, _
                                                            parConexion As ADODB.Connection) As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto digitado ya fue utilizado en otra facturacion
'**************************'***********************************************************
'   AUTOR: Germán Fajardo G.
'   FECHA: 18/06/2004
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Verificar si el incidente ya esta ligado a alguna facturacion
    varComando = "SELECT COUNT(*) AS Cuenta" _
    & " FROM Incident INNER JOIN CT_RestriccionTabTelefonia ON Incident.iIncidentTypeId = CT_RestriccionTabTelefonia.iIncidentTypeId " _
    & " Where (Incident.iIncidentId= " & parAsunto & ")"
    varResultado.Open varComando, parConexion
    
    If varResultado.EOF Then
        FunGValidarRestriccionesTab = True
    Else
        If IsNull(varResultado.Fields(0)) Then
             FunGValidarRestriccionesTab = True
        Else
            If varResultado.Fields(0) > 0 Then
                MsgBox "El tipo de novedad no permite modificación sobre el Tab de Voz ", vbInformation, App.Title
                FunGValidarRestriccionesTab = False
            Else
                FunGValidarRestriccionesTab = True
            End If
        End If
    End If
    Set varResultado = Nothing
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function FunGConsulta() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varRestriccionTabFac As claRestriccionTabTel
    On Error GoTo ErrManager
    
    varComando = "Select    iIncidentTypeId " & _
                 "From      CT_RestriccionTabFac   "
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iFacturacionTarifaId")) Then
            varRestriccionTabFac.proiIncidentTypeId = ""
        Else
            varRestriccionTabFac.proiIncidentTypeId = Trim(varResultado.Fields("iFacturacionTarifaId"))
        End If
        Me.Add varResultados.Fields("iIncidentTypeId").Value, Me.proConexion
        'Siguiente registro del cursor
        varResultados.MoveNext
      Wend
    Set varResultado = Nothing
    Set varRestriccionTabFac = Nothing
    
    FunGConsulta = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Function Add(proiIncidentTypeId As String, proConexion As ADODB.Connection) As claRestriccionTabTel
    'create a new object
    Dim objNewMember As claRestriccionTabTel
    Set objNewMember = New claRestriccionTabTel
    'set the properties passed into the method
    objNewMember.proiIncidentTypeId = proiIncidentTypeId
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    Set objNewMember.colRestriccionTabFac = colRestriccionTabFac
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claRestriccionTabTel
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

