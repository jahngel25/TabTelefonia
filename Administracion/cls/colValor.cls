VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claValor"
Attribute VB_Ext_KEY = "Member0" ,"claValor"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variable de Conexión
Public proConexion As ADODB.Connection

Public proProductMaster As String
Public proCampo As String

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proRecordStatus As String, proValorDesc As String, proValorId As String, Optional sKey As String) As claValor
    'create a new object
    Dim objNewMember As claValor
    Set objNewMember = New claValor


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proRecordStatus = proRecordStatus
    objNewMember.proValorDesc = proValorDesc
    objNewMember.proValorId = proValorId
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claValor
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
'****************************************************************
' OBJETIVO: Consultar los valores parametrizados en el sistema
'****************************************************************
' AUTOR: Gustavo Gavilan
' FECHA: 03/03/2004
'****************************************************************

    Dim varResultados As ADODB.Recordset
    Dim varComando As String
    Dim varValor As claValor
    On Error GoTo ErrorManager


    varComando = "SELECT top 200   iValorId,       " & _
                 "          vchValorDesc,   " & _
                 "          tiRecordStatus  " & _
                 "FROM      CT_VALORDATOS   "
                 
    'Trae los datos de la base
    Set varResultados = New ADODB.Recordset
        
    varResultados.Open varComando, Me.proConexion
        
    Set varValor = New claValor
    'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
    While varResultados.EOF = False
        
        'Codigo
        If IsNull(varResultados.Fields("iValorId")) Then
            varValor.proValorId = 0
        Else
            varValor.proValorId = varResultados.Fields("iValorId")
        End If
        
        'Descripcion
        If IsNull(varResultados.Fields("vchValorDesc")) Then
            varValor.proValorDesc = ""
        Else
            varValor.proValorDesc = varResultados.Fields("vchValorDesc")
        End If
        
        'Estado
        If IsNull(varResultados.Fields("tiRecordStatus")) Then
            varValor.proRecordStatus = 0
        Else
            varValor.proRecordStatus = varResultados.Fields("tiRecordStatus")
        End If
        
        Me.Add Me.proConexion, _
                varValor.proRecordStatus, _
                varValor.proValorDesc, _
                varValor.proValorId
                    
        'Siguiente registro del cursor
        varResultados.MoveNext
    Wend
                
        Set varResultados = Nothing
        Set varValor = Nothing
        
        MetConsultar = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function MetConsultarSemejantes(proValor As String) As Boolean
'****************************************************************
' OBJETIVO: Consultar los valores parametrizados en el sistema
'****************************************************************
' AUTOR: Gustavo Gavilan
' FECHA: 03/03/2004
'****************************************************************

    Dim varResultados As ADODB.Recordset
    Dim varComando As String
    Dim varValor As claValor
    Dim i As Integer
    On Error GoTo ErrorManager
    For i = 1 To Me.Count
        Me.Remove 1
    Next
    varComando = "SELECT top 200   iValorId,       " & _
                 "          vchValorDesc,   " & _
                 "          tiRecordStatus  " & _
                 "FROM      CT_VALORDATOS   "
    If Trim(proValor) <> "" Then
        varComando = varComando & " WHERE (vchValorDesc LIKE '%" & proValor & "%') " _
                                                      & " ORDER BY vchValorDesc "
    End If
                 
    'Trae los datos de la base
    Set varResultados = New ADODB.Recordset
        
    varResultados.Open varComando, Me.proConexion
        
    Set varValor = New claValor
    'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
    While varResultados.EOF = False
        
        'Codigo
        If IsNull(varResultados.Fields("iValorId")) Then
            varValor.proValorId = 0
        Else
            varValor.proValorId = varResultados.Fields("iValorId")
        End If
        
        'Descripcion
        If IsNull(varResultados.Fields("vchValorDesc")) Then
            varValor.proValorDesc = ""
        Else
            varValor.proValorDesc = varResultados.Fields("vchValorDesc")
        End If
        
        'Estado
        If IsNull(varResultados.Fields("tiRecordStatus")) Then
            varValor.proRecordStatus = 0
        Else
            varValor.proRecordStatus = varResultados.Fields("tiRecordStatus")
        End If
        
        Me.Add Me.proConexion, _
                varValor.proRecordStatus, _
                varValor.proValorDesc, _
                varValor.proValorId
                    
        'Siguiente registro del cursor
        varResultados.MoveNext
    Wend
                
        Set varResultados = Nothing
        Set varValor = Nothing
        
        MetConsultarSemejantes = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
        MetConsultarSemejantes = False
End Function


Public Function MetConsultarSinAsignar() As Boolean
'****************************************************************
' OBJETIVO: Consultar los valores parametrizados en el sistema
'****************************************************************
' AUTOR: Gustavo Gavilan
' FECHA: 03/03/2004
'****************************************************************

    Dim varResultados As ADODB.Recordset
    Dim varComando As String
    Dim varValor As claValor
    On Error GoTo ErrorManager


    varComando = "SELECT    iValorId,                   " & _
                 "          vchValorDesc,               " & _
                 "          tiRecordStatus              " & _
                 "FROM      CT_ValorDatos               " & _
                 "WHERE     iValorId not in (   SELECT  D.iValorId                  " & _
                 "                              FROM    CT_ValoresCampoXProducto V, " & _
                 "                                      CT_ValorDatos            D  " & _
                 "                              WHERE   V.chProductNumber   = '" & Me.proProductMaster & "' " & _
                 "                              AND     V.chCampo           = '" & Me.proCampo & "' " & _
                 "                              AND     D.iValorId = V.iValorId)    " & _
                 "AND       tiRecordStatus = 1"
                 
    'Trae los datos de la base
    Set varResultados = New ADODB.Recordset
        
    varResultados.Open varComando, Me.proConexion
        
    Set varValor = New claValor
    'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
    While varResultados.EOF = False
        
        'Codigo
        If IsNull(varResultados.Fields("iValorId")) Then
            varValor.proValorId = 0
        Else
            varValor.proValorId = varResultados.Fields("iValorId")
        End If
        
        'Descripcion
        If IsNull(varResultados.Fields("vchValorDesc")) Then
            varValor.proValorDesc = ""
        Else
            varValor.proValorDesc = varResultados.Fields("vchValorDesc")
        End If
        
        'Estado
        If IsNull(varResultados.Fields("tiRecordStatus")) Then
            varValor.proRecordStatus = 0
        Else
            varValor.proRecordStatus = varResultados.Fields("tiRecordStatus")
        End If
        
        Me.Add Me.proConexion, _
                varValor.proRecordStatus, _
                varValor.proValorDesc, _
                varValor.proValorId
                    
        'Siguiente registro del cursor
        varResultados.MoveNext
    Wend
                
        Set varResultados = Nothing
        Set varValor = Nothing
        
        MetConsultarSinAsignar = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function


Public Function MetConsultarAsignados() As Boolean
'****************************************************************
' OBJETIVO: Consultar los valores parametrizados en el sistema
'****************************************************************
' AUTOR: Gustavo Gavilan
' FECHA: 03/03/2004
'****************************************************************

    Dim varResultados As ADODB.Recordset
    Dim varComando As String
    Dim varValor As claValor
    On Error GoTo ErrorManager


    varComando = "SELECT    D.iValorId,                     " & _
                 "          D.vchValorDesc,                 " & _
                 "          D.tiRecordStatus                " & _
                 "FROM      CT_ValoresCampoXProducto    V,  " & _
                 "          CT_ValorDatos               D   " & _
                 "WHERE     V.chProductNumber = '" & Me.proProductMaster & "' " & _
                 "AND       V.chCampo = '" & Me.proCampo & "' " & _
                 "AND       D.iValorId = V.iValorId         "
                 
    'Trae los datos de la base
    Set varResultados = New ADODB.Recordset
        
    varResultados.Open varComando, Me.proConexion
        
    Set varValor = New claValor
    'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
    While varResultados.EOF = False
        
        'Codigo
        If IsNull(varResultados.Fields("iValorId")) Then
            varValor.proValorId = 0
        Else
            varValor.proValorId = varResultados.Fields("iValorId")
        End If
        
        'Descripcion
        If IsNull(varResultados.Fields("vchValorDesc")) Then
            varValor.proValorDesc = ""
        Else
            varValor.proValorDesc = varResultados.Fields("vchValorDesc")
        End If
        
        'Estado
        If IsNull(varResultados.Fields("tiRecordStatus")) Then
            varValor.proRecordStatus = 0
        Else
            varValor.proRecordStatus = varResultados.Fields("tiRecordStatus")
        End If
        
        Me.Add Me.proConexion, _
                varValor.proRecordStatus, _
                varValor.proValorDesc, _
                varValor.proValorId
                    
        'Siguiente registro del cursor
        varResultados.MoveNext
    Wend
        
        Set varResultados = Nothing
        Set varValor = Nothing

        MetConsultarAsignados = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

