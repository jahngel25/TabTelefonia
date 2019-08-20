VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colNumeracionCorporativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claNumeracionCorporativa"
Attribute VB_Ext_KEY = "Member0" ,"claNumeracionCorporativa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proDatosProductoId   As String

Public proConexion As ADODB.Connection

Public proSeleccionados As Integer

Public Function Add(proConexion As ADODB.Connection, proMarcacion As String, _
                    proDatosProductoId As Long, proVirtual As String, Optional sKey As String _
                    ) As claNumeracionCorporativa
    'create a new object
    Dim objNewMember As claNumeracionCorporativa
    Set objNewMember = New claNumeracionCorporativa


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proSeleccion = "0"
    objNewMember.proMarcacion = proMarcacion
    objNewMember.proDatosProductoId = proDatosProductoId
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

Public Property Get Item(vntIndexKey As Variant) As claNumeracionCorporativa
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
    proSeleccionados = 0
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Public Function FunGConsulta() As Boolean
    Dim varResultado As ADODB.Recordset
    Dim varNumeracioncorporativa As claNumeracionCorporativa
    Dim varPosicion As Integer
    
    On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
              
        Set mCol = Nothing
        Set mCol = New Collection
        
        varResultado.Open " SELECT DISTINCT " & _
                          "         iDatosProductoId, " & _
                          "         vchMarcacion, " & _
                          "         chVirtual" & _
                          " FROM    CT_Numeracioncorporativa " & _
                          " WHERE   iDatosProductoId =  " & Me.proDatosProductoId, Me.proConexion
                          'chVirtual Agregado por Carlos Castelblanco 2006/07/26
                          
        Set varNumeracioncorporativa = New claNumeracionCorporativa
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
            
            varNumeracioncorporativa.proDatosProductoId = IIf(IsNull(varResultado("iDatosProductoId")), "", varResultado.Fields("iDatosProductoId"))
            
            varNumeracioncorporativa.proMarcacion = IIf(IsNull(varResultado("vchMarcacion")), "", varResultado.Fields("vchMarcacion"))
            'Agregado por Carlos Castelblanco 2006/07/26:
            varNumeracioncorporativa.proVirtual = IIf(IsNull(varResultado("chVirtual")), "", varResultado.Fields("chVirtual"))
            
            Me.Add Me.proConexion, _
                   varNumeracioncorporativa.proMarcacion, _
                   varNumeracioncorporativa.proDatosProductoId, _
                   varNumeracioncorporativa.proVirtual
                   
                       
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        Set varNumeracioncorporativa = Nothing
        Set varResultado = Nothing
        
        FunGConsulta = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

