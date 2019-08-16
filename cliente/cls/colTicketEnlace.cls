VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colticketEnlace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"colticketsenlace"
Attribute VB_Ext_KEY = "Member0" ,"colticketsenlace"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Propiedad de conexión de la coleccion
Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection
Private Const TiposTickets = "106040, 102935"
'Variables para consulta
Public provchSerialNumber As String
Public proiIncidentId As String
Public proiIncidentTypeId As String


Public Function FunGConsulta() As Boolean
Dim varResultados As ADODB.Recordset
Dim varticketsenlace As claTicketEnlace
Dim varPosicion As Integer
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        Set varticketsenlace = New claTicketEnlace
        Set mCol = Nothing
        Set mCol = New Collection
            
        varResultados.Open " SELECT     Incident.iIncidentId, Incident.vchDesc1, Incident.chAssignedTo, reference_parameters.vchParameterDesc, Incident.dtInsertDate, " _
                  & " Incident.vchSerialNumber, Incident.iIncidentTypeId, Incident.iIncidentCategory, reference_parameters_1.vchParameterDesc AS Estado,  Incident.iStatusId " _
                  & " FROM         Incident INNER JOIN " _
                  & " reference_parameters ON Incident.iIncidentTypeId = reference_parameters.iParameterId INNER JOIN " _
                  & " reference_parameters reference_parameters_1 ON Incident.iStatusId = reference_parameters_1.iParameterId " _
                  & " WHERE     (Incident.iIncidentCategory = 2) AND (Incident.tiRecordStatus = 1) AND (Incident.vchUser9 = '" & provchSerialNumber & "') AND (Incident.iIncidentTypeId IN (" & TiposTickets & ")) ", proConexion
        While varResultados.EOF = False
                
                varticketsenlace.proiIncidentId = IIf(IsNull(varResultados("iIncidentId")), 0, varResultados.Fields("iIncidentId"))
                varticketsenlace.provchDesc1 = IIf(IsNull(varResultados("vchDesc1")), "", varResultados.Fields("vchDesc1"))
                varticketsenlace.prochAssignedTo = IIf(IsNull(varResultados("chAssignedTo")), "", varResultados.Fields("chAssignedTo"))
                varticketsenlace.provchParameterDesc = IIf(IsNull(varResultados("vchParameterDesc")), "", varResultados.Fields("vchParameterDesc"))
                varticketsenlace.prodtInsertDate = IIf(IsNull(varResultados("dtInsertDate")), Date, varResultados.Fields("dtInsertDate"))
                varticketsenlace.provchSerialNumber = IIf(IsNull(varResultados("vchSerialNumber")), "", varResultados.Fields("vchSerialNumber"))
                varticketsenlace.proiIncidentTypeId = IIf(IsNull(varResultados("iIncidentTypeId")), 0, varResultados.Fields("iIncidentTypeId"))
                varticketsenlace.proiIncidentCategory = IIf(IsNull(varResultados("iIncidentCategory")), 0, varResultados.Fields("iIncidentCategory"))
                varticketsenlace.proiStatusId = IIf(IsNull(varResultados("iStatusId")), 0, varResultados.Fields("iStatusId"))
                varticketsenlace.prosEstado = IIf(IsNull(varResultados("Estado")), "", Trim(varResultados.Fields("Estado")))
               
                Me.Add varticketsenlace.proiIncidentId, _
                varticketsenlace.provchDesc1, _
                varticketsenlace.prochAssignedTo, _
                varticketsenlace.provchParameterDesc, _
                varticketsenlace.prodtInsertDate, _
                varticketsenlace.provchSerialNumber, _
                varticketsenlace.proiIncidentTypeId, _
                varticketsenlace.proiIncidentCategory, _
                varticketsenlace.proiStatusId, _
                varticketsenlace.prosEstado, Me.proConexion
                varResultados.MoveNext
        Wend
        
        Set varResultados = Nothing
        Set varticketsenlace = Nothing
        
        FunGConsulta = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
    Resume
End Function
Public Function Add(proiIncidentId As Long, provchDesc1 As String, prochAssignedTo As String, provchParameterDesc As String, _
                                   prodtInsertDate As Date, provchSerialNumber As String, proiIncidentTypeId As Long, proiIncidentCategory As Long, proiStatusId As Long, prosEstado As String, _
                                   Optional proConexion As Variant, Optional sKey As String) As claTicketEnlace
    'create a new object
    Dim objNewMember As claTicketEnlace
    Set objNewMember = New claTicketEnlace

    'set the properties passed into the method
    objNewMember.proiIncidentId = proiIncidentId
    objNewMember.provchDesc1 = provchDesc1
    objNewMember.prochAssignedTo = prochAssignedTo
    objNewMember.provchParameterDesc = provchParameterDesc
    objNewMember.prodtInsertDate = prodtInsertDate
    objNewMember.provchSerialNumber = provchSerialNumber
    objNewMember.proiIncidentTypeId = proiIncidentTypeId
    objNewMember.proiIncidentCategory = proiIncidentCategory
    objNewMember.proiStatusId = proiStatusId
    objNewMember.prosEstado = prosEstado
    
    If IsMissing(proConexion) = False Then
        Set objNewMember.proConexion = proConexion
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claTicketEnlace
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
