VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colProductosRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claProductosRelacionados"
Attribute VB_Ext_KEY = "Member0" ,"claProductosRelacionados"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proProductNumber As String, provchDescription As String, Optional sKey As String) As claProductosRelacionados
    'create a new object
    Dim objNewMember As claProductosRelacionados
    Set objNewMember = New claProductosRelacionados


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proProductNumber = proProductNumber
    objNewMember.provchDescription = provchDescription
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claProductosRelacionados
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

Public Function FunGConsulta() As Boolean
'****************************************************************
' OBJETIVO: Realizar la consulta de todas los productos relacionados
'****************************************************************
' AUTOR: Hernan Botache
'Fecha 02/09/2004
'****************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open "SELECT C.chProductNumber, " & _
                            " PH.vchHierarchyDesc + ' - ' + PM.vchDescription as vchDescription " & _
                            " FROM CT_PRODUCTOSRELACIONADOS C " & _
                            "inner join Product_Master PM on C.chProductNumber = PM.chProductNumber,      ProductHierarchy PH " & _
                            " Where     PH.iHierarchyId     = PM.iHierarchyId " & _
                            "order by PM.vchDescription", Me.proConexion
                            
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add Me.proConexion, varResultados("chProductNumber").Value, varResultados("vchDescription").Value
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        FunGConsulta = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function FunGConsultaNoRelacionados() As Boolean
'****************************************************************
' OBJETIVO: Realizar la consulta de todas los productos no relacionados
'****************************************************************
' AUTOR: Hernan Botache
'Fecha 02/09/2004
'****************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open " SELECT    PM.chProductNumber,      PH.vchHierarchyDesc + ' - ' + PM.vchDescription as vchDescription " & _
                            " From      Product_Master PM,      ProductHierarchy PH " & _
                            " Where     PH.iHierarchyId     = PM.iHierarchyId " & _
                            " AND       PM.tiRecordStatus   = 1  " & _
                            " AND PM.chProductNumber " & _
                            " not in (select distinct chProductNumber from CT_ProductosRelacionados) " & _
                            " ORDER BY vchDescription ", Me.proConexion
            
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add Me.proConexion, Trim(varResultados("chProductNumber").Value), Trim(varResultados("vchDescription").Value)
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        FunGConsultaNoRelacionados = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function
Public Function FunGConsultaxNombre(vchNombre As String) As Boolean
'****************************************************************
' OBJETIVO: Realizar la consulta de todas los productos no relacionados
'****************************************************************
' AUTOR: Hernan Botache
'Fecha 02/09/2004
'****************************************************************
Dim varResultados As ADODB.Recordset
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset
        varResultados.Open " SELECT    PM.chProductNumber,      PH.vchHierarchyDesc + ' - ' + PM.vchDescription as vchDescription " & _
                            " From      Product_Master PM,      ProductHierarchy PH " & _
                            " Where     PH.iHierarchyId     = PM.iHierarchyId " & _
                            " AND       PM.tiRecordStatus   = 1  and (PH.vchHierarchyDesc like '%" & vchNombre & "%' or PM.vchDescription like '%" & vchNombre & "%')" & _
                            " AND PM.chProductNumber " & _
                            " not in (select distinct chProductNumber from CT_ProductosRelacionados) " & _
                            " ORDER BY vchDescription ", Me.proConexion
            
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultados.EOF = False
            Me.Add Me.proConexion, Trim(varResultados("chProductNumber").Value), Trim(varResultados("vchDescription").Value)
                    
            'Siguiente registro del cursor
            varResultados.MoveNext
        Wend
                
        FunGConsultaxNombre = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function


