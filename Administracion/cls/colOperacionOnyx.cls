VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colOperacionOnyx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claOperacionOnyx"
Attribute VB_Ext_KEY = "Member0" ,"claOperacionOnyx"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'local variable to hold collection
Private mCol As Collection

Public proConexion As ADODB.Connection

Public proIncidente As String
Public proTipoIncidente As String
Public proCategoriaIncidente As String
Public proTipoNovedad As String
Public Function FunGEliminar(parItem As Integer) As Boolean
'****************************************************************
' OBJETIVO: Realizar la eliminación de una operación
'****************************************************************
' PARITEM:  Item a eliminar
'****************************************************************
' AUTOR: Germán A. Fajardo I&T
' FECHA: 27/07/2004
'****************************************************************
On Error GoTo ErrorManager
      
    
    
        Set Me.Item(parItem).proConexion = Me.proConexion
        FunGEliminar = Me.Item(parItem).MetEliminar
        
        If FunGEliminar = True Then
            'Elimina el segmento de la coleccion
            Me.Remove parItem
        End If
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Public Function Add(proConexion As ADODB.Connection, proIncidentCategory As String, proTipoNovedadId As String, proIncidentTypeId As String, proTipoSeccionId As String, Optional sKey As String) As claOperacionOnyx
    'create a new object
    Dim objNewMember As claOperacionOnyx
    Set objNewMember = New claOperacionOnyx


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proIncidentCategory = proIncidentCategory
    objNewMember.proTipoNovedadId = proTipoNovedadId
    objNewMember.proIncidentTypeId = proIncidentTypeId
    objNewMember.proTipoSeccionId = proTipoSeccionId
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claOperacionOnyx
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

Public Function MetConsultarxTipoCategoria() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varOperacionOnyx As claOperacionOnyx
    On Error GoTo ErrManager
    
    If Trim(Me.proTipoIncidente) = "" Or Trim(Me.proCategoriaIncidente) = "" Then
        varComando = "Select    iIncidentTypeId,    " & _
                     "          iIncidentCategory   " & _
                     "From      Incident            " & _
                     "Where     iIncidentId =       " & Me.proIncidente
                
        Set varResultado = New ADODB.Recordset
        
        varResultado.Open varComando, Me.proConexion
        
        If Not varResultado.EOF Then
            If IsNull(varResultado.Fields("iIncidentTypeId")) Then
                Me.proTipoIncidente = 0
            Else
                Me.proTipoIncidente = varResultado.Fields("iIncidentTypeId")
            End If
            
            If IsNull(varResultado.Fields("iIncidentCategory")) Then
                Me.proCategoriaIncidente = 0
            Else
                Me.proCategoriaIncidente = varResultado.Fields("iIncidentCategory")
            End If
        Else
            MetConsultarxTipoCategoria = False
        End If
        
        Set varResultado = Nothing
    End If
        
    varComando = "Select    iIncidentTypeId,    " & _
                 "          iTipoNovedadId,     " & _
                 "          iIncidentCategory,   " & _
                 "          chSeccion  " & _
                 "From      CT_OperacionOnyx    " & _
                 "Where     iIncidentTypeId =   " & Me.proTipoIncidente & " " & _
                 "And       iIncidentCategory = " & Me.proCategoriaIncidente
                    
    Set varResultado = New ADODB.Recordset
    
    Set varOperacionOnyx = New claOperacionOnyx
    
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        'Tipo Incidente
        If IsNull(varResultado.Fields("iIncidentTypeId")) Then
            varOperacionOnyx.proIncidentTypeId = 0
        Else
            varOperacionOnyx.proIncidentTypeId = varResultado.Fields("iIncidentTypeId")
        End If
        
        'Tipo Novedad
        If IsNull(varResultado.Fields("iTipoNovedadId")) Then
            varOperacionOnyx.proTipoNovedadId = 0
        Else
            varOperacionOnyx.proTipoNovedadId = varResultado.Fields("iTipoNovedadId")
        End If

        'Categoria del incidente
        If IsNull(varResultado.Fields("iIncidentCategory")) Then
            varOperacionOnyx.proIncidentCategory = 0
        Else
            varOperacionOnyx.proIncidentCategory = varResultado.Fields("iIncidentCategory")
        End If
         'Seccion
        If IsNull(varResultado.Fields("chSeccion")) Then
            varOperacionOnyx.proTipoSeccionId = "*"
        Else
            varOperacionOnyx.proTipoSeccionId = varResultado.Fields("chSeccion")
        End If
        Me.Add Me.proConexion, _
                varOperacionOnyx.proIncidentCategory, _
                varOperacionOnyx.proTipoNovedadId, _
                varOperacionOnyx.proIncidentTypeId, _
                varOperacionOnyx.proTipoSeccionId
        
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varOperacionOnyx = Nothing
    
    MetConsultarxTipoCategoria = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
    
Public Function MetConsultarTodos() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varOperacionOnyx As claOperacionOnyx
    On Error GoTo ErrManager
    varComando = "Select    iIncidentTypeId,    " & _
                 "          iTipoNovedadId,     " & _
                 "          iIncidentCategory,   " & _
                 "          chSeccion   " & _
                 " From      CT_OperacionOnyx   " & _
                 " ORDER By iIncidentTypeId,    iTipoNovedadId, iIncidentCategory "
    Set varResultado = New ADODB.Recordset
    Set varOperacionOnyx = New claOperacionOnyx
    Set varOperacionOnyx.proConexion = Me.proConexion
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        'Tipo Incidente
        If IsNull(varResultado.Fields("iIncidentTypeId")) Then
            varOperacionOnyx.proIncidentTypeId = 0
        Else
            varOperacionOnyx.proIncidentTypeId = varResultado.Fields("iIncidentTypeId")
        End If
        
        'Tipo Novedad
        If IsNull(varResultado.Fields("iTipoNovedadId")) Then
            varOperacionOnyx.proTipoNovedadId = 0
        Else
            varOperacionOnyx.proTipoNovedadId = varResultado.Fields("iTipoNovedadId")
        End If

        'Categoria del incidente
        If IsNull(varResultado.Fields("iIncidentCategory")) Then
            varOperacionOnyx.proIncidentCategory = 0
        Else
            varOperacionOnyx.proIncidentCategory = varResultado.Fields("iIncidentCategory")
        End If
        
        'Seccion
        If IsNull(varResultado.Fields("chSeccion")) Then
            varOperacionOnyx.proTipoSeccionId = "*"
        Else
            varOperacionOnyx.proTipoSeccionId = varResultado.Fields("chSeccion")
        End If
        Me.Add Me.proConexion, _
                varOperacionOnyx.proIncidentCategory, _
                varOperacionOnyx.proTipoNovedadId, _
                varOperacionOnyx.proIncidentTypeId, _
                varOperacionOnyx.proTipoSeccionId
                
        
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varOperacionOnyx = Nothing
    
    MetConsultarTodos = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

