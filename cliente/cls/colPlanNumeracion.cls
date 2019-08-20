VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colPlanNumeracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claPlanNumeracion"
Attribute VB_Ext_KEY = "Member0" ,"claPlanNumeracion"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proCliente As String

Const EstadoActivado = "103519"
Const EstadoCerrado = "104"

Const TelefoniaLocal = "1810"

Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proMarcacion As String, _
                    proDatosProductoId As String, proAlias As String, proSerialNumber As String, _
                    proProductId As String, proVirtual As String, Optional sKey As String) _
                    As claPlanNumeracion
    'create a new object
    Dim objNewMember As claPlanNumeracion
    Set objNewMember = New claPlanNumeracion


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proMarcacion = proMarcacion
    objNewMember.proDatosProductoId = proDatosProductoId
    objNewMember.proAlias = proAlias
    objNewMember.proSerialNumber = proSerialNumber
    objNewMember.proProductId = proProductId
    objNewMember.proVirtual = proVirtual
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claPlanNumeracion
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

Public Function MetConsultaActuales() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varPlanNumeracion As claPlanNumeracion
    On Error GoTo ErrManager
    
    varComando = "Select    CP.iProductId,                                          " & _
                 "          CP.vchSerialNumber,                                     " & _
                 "          'Alias' = ( Select  vchDataValue                        " & _
                 "                      From    ExpansionData                       " & _
                 "                      Where   iSystemId   = CP.iProductId         " & _
                 "                      And     chFieldName = 'Product.Expand19'    " & _
                 "                      And     tiRecordStatus  = 1),               " & _
                 "          DP.iDatosProductoId,                                    " & _
                 "          NC.vchMarcacion,                                        " & _
                 "          NC.chVirtual                                            " & _
                 "From      CT_DatosProducto            DP,                         " & _
                 "          CT_NumeracionCorporativa    NC,                         " & _
                 "          Customer_Product CP                                     " & _
                 "Where     CP.iOwnerId         = " & Me.proCliente & "             " & _
                 "And       CP.chProductNumber  = '" & TelefoniaLocal & "'          " & _
                 "And       CP.iStatusId        = " & EstadoActivado & "            " & _
                 "And       DP.iProductId       = CP.iProductId                     " & _
                 "And       NC.iDatosProductoId = DP.iDatosProductoId               " & _
                 "Order by  CP.vchSerialNumber, Alias, DP.iDatosProductoId          "
                 'chVirtual Agregado por Carlos Castelblanco 2006/07/26
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varPlanNumeracion = New claPlanNumeracion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iProductId")) Then
            varPlanNumeracion.proProductId = ""
        Else
            varPlanNumeracion.proProductId = Trim(varResultado.Fields("iProductId"))
        End If
        
        If IsNull(varResultado.Fields("vchSerialNumber")) Then
            varPlanNumeracion.proSerialNumber = ""
        Else
            varPlanNumeracion.proSerialNumber = Trim(varResultado.Fields("vchSerialNumber"))
        End If
        
        If IsNull(varResultado.Fields("Alias")) Then
            varPlanNumeracion.proAlias = ""
        Else
            varPlanNumeracion.proAlias = Trim(varResultado.Fields("Alias"))
        End If
        
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varPlanNumeracion.proDatosProductoId = ""
        Else
            varPlanNumeracion.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("vchMarcacion")) Then
            varPlanNumeracion.proMarcacion = ""
        Else
            varPlanNumeracion.proMarcacion = Trim(varResultado.Fields("vchMarcacion"))
        End If
        
        'Agregado por Carlos Castelblanco 2006/07/26:
        If IsNull(varResultado.Fields("chVirtual")) Then
            varPlanNumeracion.proVirtual = ""
        Else
            varPlanNumeracion.proVirtual = Trim(varResultado.Fields("chVirtual"))
        End If
        
        Me.Add Me.proConexion, _
               varPlanNumeracion.proMarcacion, _
               varPlanNumeracion.proDatosProductoId, _
               varPlanNumeracion.proAlias, _
               varPlanNumeracion.proSerialNumber, _
               varPlanNumeracion.proProductId, _
               varPlanNumeracion.proVirtual
        'varPlanNumeracion.proVirtual Agregado por Carlos Castelblanco 2006/07/26
               
        varResultado.MoveNext
    Wend
    
    Set varPlanNumeracion = Nothing
    Set varResultado = Nothing
    
    MetConsultaActuales = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

