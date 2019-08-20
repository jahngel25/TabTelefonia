VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colserviciosxreserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"colServiciosxReserva"
Attribute VB_Ext_KEY = "Member0" ,"colSserviciosxReserva"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Procedure : colserviciosxreserva
' DateTime  : 13/10/2004
' Author    :
' Propósito   :
'---------------------------------------------------------------------------------------
'
Option Explicit

'Propiedad de conexión de la coleccion
Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

'Variables para consulta
Public proNovedadNumeroId As String
Public proServicioSuplementarioId As String
Public proTipoNovedadId As String
Public proNumero As String
Public proRegionCode As String

Public Function MetConsultarxServicio() As Boolean
   Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varServicioxNumero As claServiciosxNumero
    On Error GoTo ErrManager
    Set varResultado = New ADODB.Recordset
    varComando = " SELECT ISNULL(COUNT(CT_NOVEDADNUMEROS.vchNumero), 0) AS Cuantos " _
                                & " FROM CT_SERVICIOSXRESERVA INNER JOIN " _
                                & " CT_NOVEDADNUMEROS ON CT_SERVICIOSXRESERVA.iNovedadNumeroId = CT_NOVEDADNUMEROS.iNovedadNumeroId " _
                                & " WHERE (CT_NOVEDADNUMEROS.chRegionCode = '" & proRegionCode & "') AND (CT_NOVEDADNUMEROS.iNovedadNumeroId = " & proNovedadNumeroId & ") AND " _
                                & " (CT_SERVICIOSXRESERVA.iServicioSuplementarioId = " & proServicioSuplementarioId & ") AND (CT_NOVEDADNUMEROS.vchNumero = '" & proNumero & "') "
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
Public Function FunGConsulta() As Boolean
Dim varResultado As ADODB.Recordset
Dim varServiciosxreserva As claServiciosxReserva
Dim varPosicion As Integer
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
              
        
        varResultado.Open "SELECT   iNovedadNumeroId,           " & _
                          "         iServicioSuplementarioId,   " & _
                          "         iTipoNovedadId              " & _
                          "FROM     CT_ServiciosxReserva        " & _
                          "WHERE    iNovedadNumeroId         =  " & Me.proNovedadNumeroId, Me.proConexion
                          
        Set varServiciosxreserva = New claServiciosxReserva
        
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
                varServiciosxreserva.proNovedadNumeroId = IIf(IsNull(varResultado("iNovedadNumeroId")), "", varResultado.Fields("iNovedadNumeroId"))
                varServiciosxreserva.proServicioSuplementarioId = IIf(IsNull(varResultado("iServicioSuplementarioId")), "", varResultado.Fields("iServicioSuplementarioId"))
                varServiciosxreserva.proTipoNovedadId = IIf(IsNull(varResultado("iTipoNovedadId")), "", varResultado.Fields("iTipoNovedadId"))
                
                Me.Add varServiciosxreserva.proNovedadNumeroId, _
                        varServiciosxreserva.proServicioSuplementarioId, _
                        varServiciosxreserva.proTipoNovedadId, _
                        Me.proConexion
                        
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        Set varResultado = Nothing
        
        FunGConsulta = True
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function Add(proNovedadNumeroId As Long, proServicioSuplementarioId As Long, proTipoNovedadId As Long, Optional proConexion As Variant, Optional sKey As String) As claServiciosxReserva _
    'create a new object
    Dim objNewMember As claServiciosxReserva
    Set objNewMember = New claServiciosxReserva

    'set the properties passed into the method
    objNewMember.proNovedadNumeroId = proNovedadNumeroId
    objNewMember.proServicioSuplementarioId = proServicioSuplementarioId
    objNewMember.proTipoNovedadId = proTipoNovedadId
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

Public Property Get Item(vntIndexKey As Variant) As claServiciosxReserva
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
