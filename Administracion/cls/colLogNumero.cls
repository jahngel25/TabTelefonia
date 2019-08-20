VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colLogNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claLogNumero"
Attribute VB_Ext_KEY = "Member0" ,"claLogNumero"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proFecha As String

Public proNumeroInicial As String
Public proNumeroFinal As String

Public proFiltrarxRango As String
Public proVerUltimaEjecucion As String

Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proFecha As String, proUsuario As String, proMensaje As String, proNumero As String, proRegionCode As String, Optional sKey As String) As claLogNumero
    'create a new object
    Dim objNewMember As claLogNumero
    Set objNewMember = New claLogNumero


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proFecha = proFecha
    objNewMember.proUsuario = proUsuario
    objNewMember.proMensaje = proMensaje
    objNewMember.proNumero = proNumero
    objNewMember.proRegionCode = proRegionCode
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claLogNumero
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

Public Function MetConsultarxFecha() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varLogNumero As claLogNumero
    On Error GoTo ErrManager
    
    varComando = "Select    top 32000           " & _
                 "          chRegionCode,       " & _
                 "          vchNumero,          " & _
                 "          vchMensaje,         " & _
                 "          chUpdateBy,         " & _
                 "          dtUpdateDate        " & _
                 "From      CT_LogNumeros       "
                 
    If Me.proVerUltimaEjecucion = "1" And Me.proFiltrarxRango = "1" Then
        varComando = varComando & "Where    dtUpdateDate = convert( datetime, '" & Me.proFecha & "')" & _
                                  "And      convert( int, vchNumero) >= " & Me.proNumeroInicial & " " & _
                                  "And      convert( int, vchNumero) <= " & Me.proNumeroFinal
    End If
    
    If Me.proVerUltimaEjecucion = "1" And Me.proFiltrarxRango = "0" Then
        varComando = varComando & "Where    convert( datetime,  dtUpdateDate)  = convert( datetime, '" & Me.proFecha & "')"
    End If
    
    If Me.proVerUltimaEjecucion = "0" And Me.proFiltrarxRango = "1" Then
        varComando = varComando & "Where    convert( int, vchNumero) >= " & Me.proNumeroInicial & " " & _
                                  "And      convert( int, vchNumero) <= " & Me.proNumeroFinal
    End If
    
    varComando = varComando & " Order by chRegionCode, dtUpdateDate, convert(int, vchNumero)"
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varLogNumero = New claLogNumero
    Set varLogNumero.proConexion = Me.proConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varLogNumero.proRegionCode = ""
        Else
            varLogNumero.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("vchNumero")) Then
            varLogNumero.proNumero = ""
        Else
            varLogNumero.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        If IsNull(varResultado.Fields("vchMensaje")) Then
            varLogNumero.proMensaje = ""
        Else
            varLogNumero.proMensaje = Trim(varResultado.Fields("vchMensaje"))
        End If
        
        If IsNull(varResultado.Fields("chUpdateBy")) Then
            varLogNumero.proUsuario = ""
        Else
            varLogNumero.proUsuario = Trim(varResultado.Fields("chUpdateBy"))
        End If
        
        If IsNull(varResultado.Fields("dtUpdateDate")) Then
            varLogNumero.proFecha = ""
        Else
            varLogNumero.proFecha = Trim(varResultado.Fields("dtUpdateDate"))
        End If
        
        Me.Add varLogNumero.proConexion, _
               varLogNumero.proFecha, _
               varLogNumero.proUsuario, _
               varLogNumero.proMensaje, _
               varLogNumero.proNumero, _
               varLogNumero.proRegionCode
               
        varResultado.MoveNext
    Wend
    
    Set varLogNumero = Nothing
    Set varResultado = Nothing
    
    MetConsultarxFecha = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
