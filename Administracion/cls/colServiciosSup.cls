VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colServiciosSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"colServiciosSup"
Attribute VB_Ext_KEY = "Member0" ,"colServiciosSup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Propiedad de conexión de la coleccion
Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

'Variables para consulta
Public proiServicioSuplementarioId As String
Public provchNombreServicio As String
Public prochProductNumber As String
Public prochTipoServicio As String
Public User As String
Public Incident As String
Public Client As String

Public Function FunGConsulta() As Boolean
Dim varResultado As ADODB.Recordset
Dim varServiciosSup As claServiciosSup
Dim varPosicion As Integer
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
        Set varServiciosSup = New claServiciosSup
        Set mCol = Nothing
        Set mCol = New Collection
        varResultado.Open "SELECT DISTINCT " & _
                                       "iServicioSuplementarioId, " & _
                                       "vchNombreServicio, " & _
                                       "chProductNumber, " & _
                                       "chTipoServicio " & _
                                       " FROM ct_serviciossuplementarios " & _
                                       " WHERE chProductNumber =  '" & Me.prochProductNumber & "'", Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
                varServiciosSup.proiServicioSuplementarioId = IIf(IsNull(varResultado("iServicioSuplementarioId")), "", varResultado.Fields("iServicioSuplementarioId"))
                varServiciosSup.provchNombreServicio = IIf(IsNull(varResultado("vchNombreServicio")), "", varResultado.Fields("vchNombreServicio"))
                varServiciosSup.prochProductNumber = IIf(IsNull(varResultado("chProductNumber")), "", varResultado.Fields("chProductNumber"))
                varServiciosSup.prochTipoServicio = IIf(IsNull(varResultado("chTipoServicio")), "", varResultado.Fields("chTipoServicio"))
                Me.Add varServiciosSup.proiServicioSuplementarioId, _
                varServiciosSup.provchNombreServicio, _
                varServiciosSup.prochProductNumber, _
                varServiciosSup.prochTipoServicio, _
                 Me.proConexion
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        FunGConsulta = True
        varResultado.Close
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Function FunGConsultaTodos() As Boolean
Dim varResultado As ADODB.Recordset
Dim varServiciosSup As claServiciosSup
Dim varPosicion As Integer
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        UserName = Me.User
        IncidentId = Me.Incident
        ClientId = Me.Client
        Set varResultado = New ADODB.Recordset
        Set varServiciosSup = New claServiciosSup
        Set mCol = Nothing
        Set mCol = New Collection
        varResultado.Open "SELECT DISTINCT " & _
                                       "iServicioSuplementarioId, " & _
                                       "vchNombreServicio, " & _
                                       "chProductNumber, " & _
                                       "chTipoServicio " & _
                                       " FROM ct_serviciossuplementarios ", Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
                varServiciosSup.proiServicioSuplementarioId = IIf(IsNull(varResultado("iServicioSuplementarioId")), "", varResultado.Fields("iServicioSuplementarioId"))
                varServiciosSup.provchNombreServicio = IIf(IsNull(varResultado("vchNombreServicio")), "", varResultado.Fields("vchNombreServicio"))
                varServiciosSup.prochProductNumber = IIf(IsNull(varResultado("chProductNumber")), "", varResultado.Fields("chProductNumber"))
                varServiciosSup.prochTipoServicio = IIf(IsNull(varResultado("chTipoServicio")), "", varResultado.Fields("chTipoServicio"))
                Me.Add varServiciosSup.proiServicioSuplementarioId, _
                varServiciosSup.provchNombreServicio, _
                varServiciosSup.prochProductNumber, _
                varServiciosSup.prochTipoServicio, _
                 Me.proConexion
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        FunGConsultaTodos = True
        varResultado.Close
        Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function Add(proiServicioSuplementarioId As String, provchNombreServicio As String, prochProductNumber As String, prochTipoServicio As String, Optional proConexion As Variant, Optional sKey As String) As claServiciosSup _
    'create a new object
    Dim objNewMember As claServiciosSup
    Set objNewMember = New claServiciosSup

    'set the properties passed into the method
    objNewMember.proiServicioSuplementarioId = proiServicioSuplementarioId
    objNewMember.provchNombreServicio = provchNombreServicio
    objNewMember.prochProductNumber = prochProductNumber
    objNewMember.prochTipoServicio = prochTipoServicio
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

Public Property Get Item(vntIndexKey As Variant) As claServiciosSup
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

