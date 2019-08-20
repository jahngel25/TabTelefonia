VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colPlanNumeracionEnCurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claPlanNumeracionEnCurso"
Attribute VB_Ext_KEY = "Member0" ,"claPlanNumeracionEnCurso"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const EstadoCerrada = "101465"
Const EstadoCancelada = "103679"
Const TelefoniaLocal = "1810"

'local variable to hold collection
Private mCol As Collection

Public proCliente As String

Public proConexion As ADODB.Connection


Public Function Add(proConexion As ADODB.Connection, proMarcacion As String, proIncidentId As String, _
                    proCategoria As String, proDatosProductoId As String, _
                    proVirtual As String, Optional sKey As String) As claPlanNumeracionEnCurso
    'create a new object
    Dim objNewMember As claPlanNumeracionEnCurso
    Set objNewMember = New claPlanNumeracionEnCurso

    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proMarcacion = proMarcacion
    objNewMember.proIncidentId = proIncidentId
    objNewMember.proCategoria = proCategoria
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

Public Property Get Item(vntIndexKey As Variant) As claPlanNumeracionEnCurso
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


Public Function MetConsultaEnCurso() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varPlanNumeracionEnCurso As claPlanNumeracionEnCurso
    On Error GoTo ErrManager
    
    varComando = "Select    NN.iDatosProductoId,                            " & _
                 "          'Categoria' =   case I.iIncidentCategory        " & _
                 "                              when    2 then  'Atencion'  " & _
                 "                              when    3 then  'Venta'     " & _
                 "                          end,                            " & _
                 "          NN.iIncidentId,                                 " & _
                 "          NN.vchMarcacion,                                " & _
                 "          NN.chVirtual                                    " & _
                 "From      Incident                            I,          " & _
                 "          CT_NovedadNumeracionCorporativa     NN,         " & _
                 "          CT_DatosProductoIncident            DP,         " & _
                 "          CT_OperacionOnyx                    OO,         " & _
                 "          ReferenceDefinition                 RD          "
                 ' NN.chVirtual Agregado por Carlos Castelblanco 2006/07/26
                 
    varComando = varComando & _
                 "Where     RD.vchExtraData = '" & TelefoniaLocal & "'      " & _
                 "And       I.vchUser5  = RD.iParameterId                   " & _
                 "And       I.iIncidentCategory in (2, 3)                   " & _
                 "And       I.iOwnerId = " & Me.proCliente & "              " & _
                 "And       DP.iIncidentId  = I.iIncidentId                 " & _
                 "And       I.iStatusId not in (" & EstadoCerrado & ")      " & _
                 "And       I.tiRecordStatus = 1                            " & _
                 "And       NN.iDatosProductoId = DP.iDatosProductoId       " & _
                 "And       NN.iIncidentId  = DP.iIncidentId                " & _
                 "And       NN.iTipoNovedadId = 1                           " & _
                 "And       OO.iIncidentCategory = I.iIncidentCategory      " & _
                 "And       OO.iincidentTypeId = I.iIncidentTypeId          " & _
                 "And       OO.iTipoNovedadId = 1                           "
    
    varComando = varComando & _
                 "And RD.iReferenceId = 541                                 " & _
                 "And (OO.chSeccion = 'C' or OO.chSeccion is Null)          " & _
                 "And I.iIncidentId not in(                                 " & _
                 "          Select  convert(int, vchuser7)                  " & _
                 "          From    Incident                                " & _
                 "          Where   iIncidentCategory = 1                   " & _
                 "          And     vchUser7    = I.iIncidentId             " & _
                 "          And     tiRecordStatus  = 1                     " & _
                 "          And     iStatusId   in (" & EstadoCerrada & ",  " & _
                 "                                  " & EstadoCancelada & ")) " & _
                 "Order By NN.iIncidentId                                     "
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varPlanNumeracionEnCurso = New claPlanNumeracionEnCurso
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varPlanNumeracionEnCurso.proDatosProductoId = ""
        Else
            varPlanNumeracionEnCurso.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("Categoria")) Then
            varPlanNumeracionEnCurso.proCategoria = ""
        Else
            varPlanNumeracionEnCurso.proCategoria = Trim(varResultado.Fields("Categoria"))
        End If
        
        If IsNull(varResultado.Fields("iIncidentId")) Then
            varPlanNumeracionEnCurso.proIncidentId = ""
        Else
            varPlanNumeracionEnCurso.proIncidentId = Trim(varResultado.Fields("iIncidentId"))
        End If
        
        If IsNull(varResultado.Fields("vchMarcacion")) Then
            varPlanNumeracionEnCurso.proMarcacion = ""
        Else
            varPlanNumeracionEnCurso.proMarcacion = Trim(varResultado.Fields("vchMarcacion"))
        End If
        
        'Agregado por Carlos Castelblanco 20006/07/26:
        If IsNull(varResultado.Fields("chVirtual")) Then
            varPlanNumeracionEnCurso.proVirtual = ""
        Else
            varPlanNumeracionEnCurso.proVirtual = Trim(varResultado.Fields("chVirtual"))
        End If
        
        Me.Add Me.proConexion, _
               varPlanNumeracionEnCurso.proMarcacion, _
               varPlanNumeracionEnCurso.proIncidentId, _
               varPlanNumeracionEnCurso.proCategoria, _
               varPlanNumeracionEnCurso.proDatosProductoId, _
               varPlanNumeracionEnCurso.proVirtual
               
               
               
        varResultado.MoveNext
    Wend
    
    Set varPlanNumeracionEnCurso = Nothing
    Set varResultado = Nothing
    
    MetConsultaEnCurso = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function




