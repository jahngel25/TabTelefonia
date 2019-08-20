VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colServiciosxNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claServiciosxNumero"
Attribute VB_Ext_KEY = "Member0" ,"claServiciosxNumero"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proConexion As ADODB.Connection
Public proDatosProductoId As String
Public proServicioID As String
Public proRegionCode As String
Public proNumero As String


'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proFechaActivacion As String, _
                    proNumero As String, proRegionName As String, _
                    proRegionCode As String, proDatosProductoId As String, _
                    proNombreServicio As String, proServicioSuplementarioId As String, _
                    Optional sKey As String) As claServiciosxNumero

    'create a new object
    Dim objNewMember As claServiciosxNumero
    Set objNewMember = New claServiciosxNumero


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    objNewMember.proFechaActivacion = proFechaActivacion
    objNewMember.proNumero = proNumero
    objNewMember.proRegionName = proRegionName
    objNewMember.proRegionCode = proRegionCode
    objNewMember.proDatosProductoId = proDatosProductoId
    objNewMember.proNombreServicio = proNombreServicio
    objNewMember.proServicioSuplementarioId = proServicioSuplementarioId
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claServiciosxNumero
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
Public Function MetConsultarxServicio() As Boolean
   Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varServicioxNumero As claServiciosxNumero
    On Error GoTo ErrManager
    Set varResultado = New ADODB.Recordset
    varComando = "SELECT IsNull(Count(iServicioSuplementarioId),0) Cuantos " & _
                                       " FROM ct_serviciosxnumero " & _
                                       " WHERE iServicioSuplementarioId =  " & Me.proServicioID & _
                                       " AND iDatosProductoId =  " & Me.proDatosProductoId & _
                                       " AND chRegionCode =  '" & Me.proRegionCode & _
                                       "' AND vchNumero =  '" & Me.proNumero & "'"
     varResultado.Open varComando, Me.proConexion
    If varResultado.EOF Then
        MetConsultarxServicio = False
        Exit Function
    Else
        If varResultado.Fields("Cuantos") = 0 Then
            MetConsultarxServicio = False
            Exit Function
        End If
    End If
    Set varResultado = Nothing
    Set varServicioxNumero = Nothing
    MetConsultarxServicio = True
    
    Exit Function
ErrManager:
    MetConsultarxServicio = False
    SubGMuestraError
End Function
Public Function MetConsultar() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varServicioxNumero As claServiciosxNumero
    On Error GoTo ErrManager
    
 varComando = "Select SN.iServicioSuplementarioId, " & _
                            " SS.vchNombreServicio, " & _
                            " SN.iDatosProductoId, " & _
                            " SN.chRegionCode, " & _
                            " R.chRegionName, " & _
                            " SN.vchNumero, " & _
                            " SN.dtFechaActivacion " & _
                            " From CT_ServiciosxNumero SN, " & _
                            " CT_ServiciosSuplementarios SS, " & _
                            " Region R " & _
                            " Where SS.iServicioSuplementarioId = SN.iServicioSuplementarioId " & _
                            " And R.chRegionCode = SN.chRegionCode " & _
                            " And SN.iDatosProductoId = " & Me.proDatosProductoId & " " & _
                            "Order by SN.iDatosProductoId, SN.chRegionCode, SN.vchNumero, SN.iServicioSuplementarioId "
    
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varServicioxNumero = New claServiciosxNumero
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iServicioSuplementarioId")) Then
            varServicioxNumero.proServicioSuplementarioId = ""
        Else
            varServicioxNumero.proServicioSuplementarioId = Trim(varResultado.Fields("iServicioSuplementarioId"))
        End If
        
        If IsNull(varResultado.Fields("vchNombreServicio")) Then
            varServicioxNumero.proNombreServicio = ""
        Else
            varServicioxNumero.proNombreServicio = Trim(varResultado.Fields("vchNombreServicio"))
        End If
        
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varServicioxNumero.proDatosProductoId = ""
        Else
            varServicioxNumero.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varServicioxNumero.proRegionCode = ""
        Else
            varServicioxNumero.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("chRegionName")) Then
            varServicioxNumero.proRegionName = ""
        Else
            varServicioxNumero.proRegionName = Trim(varResultado.Fields("chRegionName"))
        End If
        
        If IsNull(varResultado.Fields("vchNumero")) Then
            varServicioxNumero.proNumero = ""
        Else
            varServicioxNumero.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        If IsNull(varResultado.Fields("dtFechaActivacion")) Then
            varServicioxNumero.proFechaActivacion = ""
        Else
            varServicioxNumero.proFechaActivacion = Trim(varResultado.Fields("dtFechaActivacion"))
        End If
        
        Me.Add Me.proConexion, _
               varServicioxNumero.proFechaActivacion, _
               varServicioxNumero.proNumero, _
               varServicioxNumero.proRegionName, _
               varServicioxNumero.proRegionCode, _
               varServicioxNumero.proDatosProductoId, _
               varServicioxNumero.proNombreServicio, _
               varServicioxNumero.proServicioSuplementarioId
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varServicioxNumero = Nothing
    
    MetConsultar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function
