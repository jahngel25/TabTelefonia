VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colProductMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claProductMaster"
Attribute VB_Ext_KEY = "Member0" ,"claProductMaster"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proIncidentId As String
'Conexion
Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proModifiedDate As String, proRecordStatus As String, proUpdateDate As String, proUpdateBy As String, proInsertDate As String, proInsertBy As String, proStatusId As String, proObsoleteDate As String, dtFirstShipDate As String, proOrderable As String, proIncidentProductNumber As String, proWarrantyExpDate As String, proWarrantyDays As String, proWarrantyCalls As String, proSupportType As String, proCost As String, proListPrice As String, proStockLevel As String, proUser10 As String, proUser9 As String, proUser8 As String, proUser7 As String, proUser6 As String, proUser5 As String, proUser4 As String, proUser3 As String, proUser2 As String, proUser1 As String, proDescription As String, proHierarchyId As String, proSecondaryId As String, proAlternateLangId As String, proSiteId As String, proProductNumber As String, Optional sKey As String) As claProductMaster
    'create a new object
    Dim objNewMember As claProductMaster
    Set objNewMember = New claProductMaster


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proModifiedDate = proModifiedDate
    objNewMember.proRecordStatus = proRecordStatus
    objNewMember.proUpdateDate = proUpdateDate
    objNewMember.proUpdateBy = proUpdateBy
    objNewMember.proInsertDate = proInsertDate
    objNewMember.proInsertBy = proInsertBy
    objNewMember.proStatusId = proStatusId
    objNewMember.proObsoleteDate = proObsoleteDate
    objNewMember.dtFirstShipDate = dtFirstShipDate
    objNewMember.proOrderable = proOrderable
    objNewMember.proIncidentProductNumber = proIncidentProductNumber
    objNewMember.proWarrantyExpDate = proWarrantyExpDate
    objNewMember.proWarrantyDays = proWarrantyDays
    objNewMember.proWarrantyCalls = proWarrantyCalls
    objNewMember.proSupportType = proSupportType
    objNewMember.proCost = proCost
    objNewMember.proListPrice = proListPrice
    objNewMember.proStockLevel = proStockLevel
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
    objNewMember.proDescription = proDescription
    objNewMember.proHierarchyId = proHierarchyId
    objNewMember.proSecondaryId = proSecondaryId
    objNewMember.proAlternateLangId = proAlternateLangId
    objNewMember.proSiteId = proSiteId
    objNewMember.proProductNumber = proProductNumber
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claProductMaster
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

Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varProducto As claProductMaster
    On Error GoTo ErrManager
    
    varComando = "SELECT    PM.chProductNumber,     PM.iSiteId,                                 " & _
                 "          PM.iAlternateLangId,    PM.vchSecondaryId,                          " & _
                 "          PM.iHierarchyId,                                                    " & _
                 "          PH.vchHierarchyDesc + ' - ' + PM.vchDescription as vchDescription,  " & _
                 "          PM.vchUser1,            PM.vchUser2,                                " & _
                 "          PM.vchUser3,            PM.vchUser4,                                " & _
                 "          PM.vchUser5,            PM.vchUser6,                                " & _
                 "          PM.vchUser7,            PM.vchUser8,                                " & _
                 "          PM.vchUser9,            PM.vchUser10,                               " & _
                 "          PM.iStockLevel,         PM.mListPrice,                              " & _
                 "          PM.mCost,               PM.vchSupportType,                          " & _
                 "          PM.siWarrantyCalls,     PM.siWarrantyDays,                          " & _
                 "          PM.dtWarrantyExpDate,   PM.vchIncidentProductNumber,                " & _
                 "          PM.bOrderable,          PM.dtFirstShipDate,                         " & _
                 "          PM.dtObsoleteDate,      PM.iStatusId,                               " & _
                 "          PM.chInsertBy,          PM.dtInsertDate,                            " & _
                 "          PM.chUpdateBy,          PM.dtUpdateDate,                            " & _
                 "          PM.tiRecordStatus,      PM.dtModifiedDate                           " & _
                 "From      Product_Master PM,      ProductHierarchy PH                         " & _
                 "Where     PH.iHierarchyId     = PM.iHierarchyId                               " & _
                 "AND       PM.tiRecordStatus   = 1 ORDER BY vchDescription                     "

    Set varResultado = New ADODB.Recordset
    Set varProducto = New claProductMaster
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chProductNumber")) Then
            varProducto.proProductNumber = ""
        Else
            varProducto.proProductNumber = Trim(varResultado.Fields("chProductNumber"))
        End If
    
        If IsNull(varResultado.Fields("iSiteId")) Then
            varProducto.proSiteId = ""
        Else
            varProducto.proSiteId = Trim(varResultado.Fields("iSiteId"))
        End If
        
        If IsNull(varResultado.Fields("iAlternateLangId")) Then
            varProducto.proAlternateLangId = ""
        Else
            varProducto.proAlternateLangId = Trim(varResultado.Fields("iAlternateLangId"))
        End If
        
        If IsNull(varResultado.Fields("vchSecondaryId")) Then
            varProducto.proSecondaryId = ""
        Else
            varProducto.proSecondaryId = Trim(varResultado.Fields("vchSecondaryId"))
        End If
        
        If IsNull(varResultado.Fields("iHierarchyId")) Then
            varProducto.proHierarchyId = ""
        Else
            varProducto.proHierarchyId = Trim(varResultado.Fields("iHierarchyId"))
        End If
        
        If IsNull(varResultado.Fields("vchDescription")) Then
            varProducto.proDescription = ""
        Else
            varProducto.proDescription = Trim(varResultado.Fields("vchDescription"))
        End If
        
        If IsNull(varResultado.Fields("vchUser1")) Then
            varProducto.proUser1 = ""
        Else
            varProducto.proUser1 = Trim(varResultado.Fields("vchUser1"))
        End If
        
        If IsNull(varResultado.Fields("vchUser2")) Then
            varProducto.proUser2 = ""
        Else
            varProducto.proUser2 = Trim(varResultado.Fields("vchUser2"))
        End If
        
        If IsNull(varResultado.Fields("vchUser3")) Then
            varProducto.proUser3 = ""
        Else
            varProducto.proUser3 = Trim(varResultado.Fields("vchUser3"))
        End If
        
        If IsNull(varResultado.Fields("vchUser4")) Then
            varProducto.proUser4 = ""
        Else
            varProducto.proUser4 = Trim(varResultado.Fields("vchUser4"))
        End If
        
        If IsNull(varResultado.Fields("vchUser5")) Then
            varProducto.proUser5 = ""
        Else
            varProducto.proUser5 = Trim(varResultado.Fields("vchUser5"))
        End If
        
        If IsNull(varResultado.Fields("vchUser6")) Then
            varProducto.proUser6 = ""
        Else
            varProducto.proUser6 = Trim(varResultado.Fields("vchUser6"))
        End If
        
        If IsNull(varResultado.Fields("vchUser7")) Then
            varProducto.proUser7 = ""
        Else
            varProducto.proUser7 = Trim(varResultado.Fields("vchUser7"))
        End If
        
        If IsNull(varResultado.Fields("vchUser8")) Then
            varProducto.proUser8 = ""
        Else
            varProducto.proUser8 = Trim(varResultado.Fields("vchUser8"))
        End If
        
        If IsNull(varResultado.Fields("vchUser9")) Then
            varProducto.proUser9 = ""
        Else
            varProducto.proUser9 = Trim(varResultado.Fields("vchUser9"))
        End If
        
        If IsNull(varResultado.Fields("vchUser10")) Then
            varProducto.proUser10 = ""
        Else
            varProducto.proUser10 = Trim(varResultado.Fields("vchUser10"))
        End If
        
        If IsNull(varResultado.Fields("iStockLevel")) Then
            varProducto.proStockLevel = ""
        Else
            varProducto.proStockLevel = Trim(varResultado.Fields("iStockLevel"))
        End If
        
        If IsNull(varResultado.Fields("mListPrice")) Then
            varProducto.proListPrice = ""
        Else
            varProducto.proListPrice = Trim(varResultado.Fields("mListPrice"))
        End If
        
        If IsNull(varResultado.Fields("mCost")) Then
            varProducto.proCost = ""
        Else
            varProducto.proCost = Trim(varResultado.Fields("mCost"))
        End If
        
        If IsNull(varResultado.Fields("vchSupportType")) Then
            varProducto.proSupportType = ""
        Else
            varProducto.proSupportType = Trim(varResultado.Fields("vchSupportType"))
        End If
        
        If IsNull(varResultado.Fields("siWarrantyCalls")) Then
            varProducto.proWarrantyCalls = ""
        Else
            varProducto.proWarrantyCalls = Trim(varResultado.Fields("siWarrantyCalls"))
        End If
        
        If IsNull(varResultado.Fields("siWarrantyDays")) Then
            varProducto.proWarrantyDays = ""
        Else
            varProducto.proWarrantyDays = Trim(varResultado.Fields("siWarrantyDays"))
        End If
        
        If IsNull(varResultado.Fields("dtWarrantyExpDate")) Then
            varProducto.proWarrantyExpDate = ""
        Else
            varProducto.proWarrantyExpDate = Trim(varResultado.Fields("dtWarrantyExpDate"))
        End If
        
        If IsNull(varResultado.Fields("vchIncidentProductNumber")) Then
            varProducto.proIncidentProductNumber = ""
        Else
            varProducto.proIncidentProductNumber = Trim(varResultado.Fields("vchIncidentProductNumber"))
        End If
        
        If IsNull(varResultado.Fields("bOrderable")) Then
            varProducto.proOrderable = ""
        Else
            varProducto.proOrderable = Trim(varResultado.Fields("bOrderable"))
        End If
        
        If IsNull(varResultado.Fields("dtFirstShipDate")) Then
            varProducto.dtFirstShipDate = ""
        Else
            varProducto.dtFirstShipDate = Trim(varResultado.Fields("dtFirstShipDate"))
        End If
        
        If IsNull(varResultado.Fields("dtObsoleteDate")) Then
            varProducto.proObsoleteDate = ""
        Else
            varProducto.proObsoleteDate = Trim(varResultado.Fields("dtObsoleteDate"))
        End If
        
        If IsNull(varResultado.Fields("iStatusId")) Then
            varProducto.proStatusId = ""
        Else
            varProducto.proStatusId = Trim(varResultado.Fields("iStatusId"))
        End If
        
        If IsNull(varResultado.Fields("chInsertBy")) Then
            varProducto.proInsertBy = ""
        Else
            varProducto.proInsertBy = Trim(varResultado.Fields("chInsertBy"))
        End If
        
        If IsNull(varResultado.Fields("dtInsertDate")) Then
            varProducto.proInsertDate = ""
        Else
            varProducto.proInsertDate = Trim(varResultado.Fields("dtInsertDate"))
        End If
        
        If IsNull(varResultado.Fields("chUpdateBy")) Then
            varProducto.proUpdateBy = ""
        Else
            varProducto.proUpdateBy = Trim(varResultado.Fields("chUpdateBy"))
        End If
        
        If IsNull(varResultado.Fields("dtUpdateDate")) Then
            varProducto.proUpdateDate = ""
        Else
            varProducto.proUpdateDate = Trim(varResultado.Fields("dtUpdateDate"))
        End If
        
        If IsNull(varResultado.Fields("tiRecordStatus")) Then
            varProducto.proRecordStatus = ""
        Else
            varProducto.proRecordStatus = Trim(varResultado.Fields("tiRecordStatus"))
        End If
        
        If IsNull(varResultado.Fields("dtModifiedDate")) Then
            varProducto.proModifiedDate = ""
        Else
            varProducto.proModifiedDate = Trim(varResultado.Fields("dtModifiedDate"))
        End If
        
        Me.Add Me.proConexion, varProducto.proModifiedDate, _
                        varProducto.proRecordStatus, varProducto.proUpdateDate, _
                        varProducto.proUpdateBy, varProducto.proInsertDate, _
                        varProducto.proInsertBy, varProducto.proStatusId, _
                        varProducto.proObsoleteDate, varProducto.dtFirstShipDate, _
                        varProducto.proOrderable, varProducto.proIncidentProductNumber, _
                        varProducto.proWarrantyExpDate, varProducto.proWarrantyDays, _
                        varProducto.proWarrantyCalls, varProducto.proSupportType, _
                        varProducto.proCost, varProducto.proListPrice, _
                        varProducto.proStockLevel, varProducto.proUser10, _
                        varProducto.proUser9, varProducto.proUser8, _
                        varProducto.proUser7, varProducto.proUser6, _
                        varProducto.proUser5, varProducto.proUser4, _
                        varProducto.proUser3, varProducto.proUser2, _
                        varProducto.proUser1, varProducto.proDescription, _
                        varProducto.proHierarchyId, varProducto.proSecondaryId, _
                        varProducto.proAlternateLangId, varProducto.proSiteId, _
                        varProducto.proProductNumber
                        
        varResultado.MoveNext
    Wend
    Set varProducto = Nothing
    Set varResultado = Nothing
    
    MetConsultar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarxIncidente() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varProducto As claProductMaster
    On Error GoTo ErrManager
    
    varComando = "SELECT    PM.chProductNumber,    PM.iSiteId,                  " & _
                 "          PM.iAlternateLangId,   PM.vchSecondaryId,           " & _
                 "          PM.iHierarchyId,                                    " & _
                 "          PH.vchHierarchyDesc + ' - ' + PM.vchDescription as vchDescription,  " & _
                 "          PM.vchUser1,           PM.vchUser2,                 " & _
                 "          PM.vchUser3,           PM.vchUser4,                 " & _
                 "          PM.vchUser5,           PM.vchUser6,                 " & _
                 "          PM.vchUser7,           PM.vchUser8,                 " & _
                 "          PM.vchUser9,           PM.vchUser10,                " & _
                 "          PM.iStockLevel,        PM.mListPrice,               " & _
                 "          PM.mCost,              PM.vchSupportType,           " & _
                 "          PM.siWarrantyCalls,    PM.siWarrantyDays,           " & _
                 "          PM.dtWarrantyExpDate,  PM.vchIncidentProductNumber, " & _
                 "          PM.bOrderable,         PM.dtFirstShipDate,          " & _
                 "          PM.dtObsoleteDate,     PM.iStatusId,                " & _
                 "          PM.chInsertBy,         PM.dtInsertDate,             " & _
                 "          PM.chUpdateBy,         PM.dtUpdateDate,             " & _
                 "          PM.tiRecordStatus,     PM.dtModifiedDate            " & _
                 "FROM      ReferenceDefinition RD, Reference_Fields    RF,     " & _
                 "          Product_Master      PM, Incident I, ProductHierarchy PH " & _
                 "WHERE     RD.iReferenceId     = RF.iReferenceId               " & _
                 "AND       RD.iParameterId     = I.vchUser5                    " & _
                 "AND       RF.chFieldName      = 'Incident.Sales.User5'        " & _
                 "AND       PM.chProductNumber  = RD.vchExtraData               " & _
                 "AND       I.iIncidentId       = " & Me.proIncidentId & " And PH.iHierarchyId     = PM.iHierarchyId "


    Set varResultado = New ADODB.Recordset
    Set varProducto = New claProductMaster
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chProductNumber")) Then
            varProducto.proProductNumber = ""
        Else
            varProducto.proProductNumber = Trim(varResultado.Fields("chProductNumber"))
        End If
    
        If IsNull(varResultado.Fields("iSiteId")) Then
            varProducto.proSiteId = ""
        Else
            varProducto.proSiteId = Trim(varResultado.Fields("iSiteId"))
        End If
        
        If IsNull(varResultado.Fields("iAlternateLangId")) Then
            varProducto.proAlternateLangId = ""
        Else
            varProducto.proAlternateLangId = Trim(varResultado.Fields("iAlternateLangId"))
        End If
        
        If IsNull(varResultado.Fields("vchSecondaryId")) Then
            varProducto.proSecondaryId = ""
        Else
            varProducto.proSecondaryId = Trim(varResultado.Fields("vchSecondaryId"))
        End If
        
        If IsNull(varResultado.Fields("iHierarchyId")) Then
            varProducto.proHierarchyId = ""
        Else
            varProducto.proHierarchyId = Trim(varResultado.Fields("iHierarchyId"))
        End If
        
        If IsNull(varResultado.Fields("vchDescription")) Then
            varProducto.proDescription = ""
        Else
            varProducto.proDescription = Trim(varResultado.Fields("vchDescription"))
        End If
        
        If IsNull(varResultado.Fields("vchUser1")) Then
            varProducto.proUser1 = ""
        Else
            varProducto.proUser1 = Trim(varResultado.Fields("vchUser1"))
        End If
        
        If IsNull(varResultado.Fields("vchUser2")) Then
            varProducto.proUser2 = ""
        Else
            varProducto.proUser2 = Trim(varResultado.Fields("vchUser2"))
        End If
        
        If IsNull(varResultado.Fields("vchUser3")) Then
            varProducto.proUser3 = ""
        Else
            varProducto.proUser3 = Trim(varResultado.Fields("vchUser3"))
        End If
        
        If IsNull(varResultado.Fields("vchUser4")) Then
            varProducto.proUser4 = ""
        Else
            varProducto.proUser4 = Trim(varResultado.Fields("vchUser4"))
        End If
        
        If IsNull(varResultado.Fields("vchUser5")) Then
            varProducto.proUser5 = ""
        Else
            varProducto.proUser5 = Trim(varResultado.Fields("vchUser5"))
        End If
        
        If IsNull(varResultado.Fields("vchUser6")) Then
            varProducto.proUser6 = ""
        Else
            varProducto.proUser6 = Trim(varResultado.Fields("vchUser6"))
        End If
        
        If IsNull(varResultado.Fields("vchUser7")) Then
            varProducto.proUser7 = ""
        Else
            varProducto.proUser7 = Trim(varResultado.Fields("vchUser7"))
        End If
        
        If IsNull(varResultado.Fields("vchUser8")) Then
            varProducto.proUser8 = ""
        Else
            varProducto.proUser8 = Trim(varResultado.Fields("vchUser8"))
        End If
        
        If IsNull(varResultado.Fields("vchUser9")) Then
            varProducto.proUser9 = ""
        Else
            varProducto.proUser9 = Trim(varResultado.Fields("vchUser9"))
        End If
        
        If IsNull(varResultado.Fields("vchUser10")) Then
            varProducto.proUser10 = ""
        Else
            varProducto.proUser10 = Trim(varResultado.Fields("vchUser10"))
        End If
        
        If IsNull(varResultado.Fields("iStockLevel")) Then
            varProducto.proStockLevel = ""
        Else
            varProducto.proStockLevel = Trim(varResultado.Fields("iStockLevel"))
        End If
        
        If IsNull(varResultado.Fields("mListPrice")) Then
            varProducto.proListPrice = ""
        Else
            varProducto.proListPrice = Trim(varResultado.Fields("mListPrice"))
        End If
        
        If IsNull(varResultado.Fields("mCost")) Then
            varProducto.proCost = ""
        Else
            varProducto.proCost = Trim(varResultado.Fields("mCost"))
        End If
        
        If IsNull(varResultado.Fields("vchSupportType")) Then
            varProducto.proSupportType = ""
        Else
            varProducto.proSupportType = Trim(varResultado.Fields("vchSupportType"))
        End If
        
        If IsNull(varResultado.Fields("siWarrantyCalls")) Then
            varProducto.proWarrantyCalls = ""
        Else
            varProducto.proWarrantyCalls = Trim(varResultado.Fields("siWarrantyCalls"))
        End If
        
        If IsNull(varResultado.Fields("siWarrantyDays")) Then
            varProducto.proWarrantyDays = ""
        Else
            varProducto.proWarrantyDays = Trim(varResultado.Fields("siWarrantyDays"))
        End If
        
        If IsNull(varResultado.Fields("dtWarrantyExpDate")) Then
            varProducto.proWarrantyExpDate = ""
        Else
            varProducto.proWarrantyExpDate = Trim(varResultado.Fields("dtWarrantyExpDate"))
        End If
        
        If IsNull(varResultado.Fields("vchIncidentProductNumber")) Then
            varProducto.proIncidentProductNumber = ""
        Else
            varProducto.proIncidentProductNumber = Trim(varResultado.Fields("vchIncidentProductNumber"))
        End If
        
        If IsNull(varResultado.Fields("bOrderable")) Then
            varProducto.proOrderable = ""
        Else
            varProducto.proOrderable = Trim(varResultado.Fields("bOrderable"))
        End If
        
        If IsNull(varResultado.Fields("dtFirstShipDate")) Then
            varProducto.dtFirstShipDate = ""
        Else
            varProducto.dtFirstShipDate = Trim(varResultado.Fields("dtFirstShipDate"))
        End If
        
        If IsNull(varResultado.Fields("dtObsoleteDate")) Then
            varProducto.proObsoleteDate = ""
        Else
            varProducto.proObsoleteDate = Trim(varResultado.Fields("dtObsoleteDate"))
        End If
        
        If IsNull(varResultado.Fields("iStatusId")) Then
            varProducto.proStatusId = ""
        Else
            varProducto.proStatusId = Trim(varResultado.Fields("iStatusId"))
        End If
        
        If IsNull(varResultado.Fields("chInsertBy")) Then
            varProducto.proInsertBy = ""
        Else
            varProducto.proInsertBy = Trim(varResultado.Fields("chInsertBy"))
        End If
        
        If IsNull(varResultado.Fields("dtInsertDate")) Then
            varProducto.proInsertDate = ""
        Else
            varProducto.proInsertDate = Trim(varResultado.Fields("dtInsertDate"))
        End If
        
        If IsNull(varResultado.Fields("chUpdateBy")) Then
            varProducto.proUpdateBy = ""
        Else
            varProducto.proUpdateBy = Trim(varResultado.Fields("chUpdateBy"))
        End If
        
        If IsNull(varResultado.Fields("dtUpdateDate")) Then
            varProducto.proUpdateDate = ""
        Else
            varProducto.proUpdateDate = Trim(varResultado.Fields("dtUpdateDate"))
        End If
        
        If IsNull(varResultado.Fields("tiRecordStatus")) Then
            varProducto.proRecordStatus = ""
        Else
            varProducto.proRecordStatus = Trim(varResultado.Fields("tiRecordStatus"))
        End If
        
        If IsNull(varResultado.Fields("dtModifiedDate")) Then
            varProducto.proModifiedDate = ""
        Else
            varProducto.proModifiedDate = Trim(varResultado.Fields("dtModifiedDate"))
        End If
        
        Me.Add Me.proConexion, varProducto.proModifiedDate, _
                        varProducto.proRecordStatus, varProducto.proUpdateDate, _
                        varProducto.proUpdateBy, varProducto.proInsertDate, _
                        varProducto.proInsertBy, varProducto.proStatusId, _
                        varProducto.proObsoleteDate, varProducto.dtFirstShipDate, _
                        varProducto.proOrderable, varProducto.proIncidentProductNumber, _
                        varProducto.proWarrantyExpDate, varProducto.proWarrantyDays, _
                        varProducto.proWarrantyCalls, varProducto.proSupportType, _
                        varProducto.proCost, varProducto.proListPrice, _
                        varProducto.proStockLevel, varProducto.proUser10, _
                        varProducto.proUser9, varProducto.proUser8, _
                        varProducto.proUser7, varProducto.proUser6, _
                        varProducto.proUser5, varProducto.proUser4, _
                        varProducto.proUser3, varProducto.proUser2, _
                        varProducto.proUser1, varProducto.proDescription, _
                        varProducto.proHierarchyId, varProducto.proSecondaryId, _
                        varProducto.proAlternateLangId, varProducto.proSiteId, _
                        varProducto.proProductNumber
                        
        varResultado.MoveNext
    Wend
    
    Set varProducto = Nothing
    Set varResultado = Nothing
    
    MetConsultarxIncidente = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


