VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colticketxdetalledatosproducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"colticketxdetalledatosproducto"
Attribute VB_Ext_KEY = "Member0" ,"colticketxdetalledatosproducto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Propiedad de conexión de la coleccion
Public proConexion As ADODB.Connection

'local variable to hold collection
Private mCol As Collection

'Variables para consulta
Public proiIncidentId As String
Public proiDatosProductoId As String
Public proiDetalleDatosProductoId As String

Public Function FunGConsulta() As Boolean
Dim varResultado As ADODB.Recordset
Dim varct_ticketxdetalledatosproducto As claticketxdetalledatosproducto
Dim varPosicion As Integer
On Error GoTo ErrorManager

   
        'Trae los datos de la base
        Set varResultado = New ADODB.Recordset
              
        Set mCol = Nothing
        Set mCol = New Collection
        Set varct_ticketxdetalledatosproducto = New claticketxdetalledatosproducto
        varResultado.Open "SELECT DISTINCT " & _
                                        "   iIncidentId, iDatosProductoId, iDetalleDatosProductoId    " & _
                                        "  FROM Ct_TicketxDetalleDatosProducto " & _
                                        " WHERE iIncidentId =  " & Me.proiIncidentId, Me.proConexion
        'Almacena los datos de la base en las propiedades, previa verificacion de no nulidad
        While varResultado.EOF = False
        
                If IsNull(varResultado.Fields("iIncidentId")) Then
                    varct_ticketxdetalledatosproducto.proiIncidentId = 0
                Else
                    varct_ticketxdetalledatosproducto.proiIncidentId = Trim(varResultado.Fields("iIncidentId"))
                End If
                If IsNull(varResultado.Fields("iDatosProductoId")) Then
                    varct_ticketxdetalledatosproducto.proiDatosProductoId = 0
                Else
                    varct_ticketxdetalledatosproducto.proiDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
                End If
                If IsNull(varResultado.Fields("iDetalleDatosProductoId")) Then
                    varct_ticketxdetalledatosproducto.proiDetalleDatosProductoId = 0
                Else
                    varct_ticketxdetalledatosproducto.proiDetalleDatosProductoId = Trim(varResultado.Fields("iDetalleDatosProductoId"))
                End If
                Me.Add varct_ticketxdetalledatosproducto.proiIncidentId, _
                varct_ticketxdetalledatosproducto.proiDatosProductoId, _
                varct_ticketxdetalledatosproducto.proiDetalleDatosProductoId, _
                 Me.proConexion
            'Siguiente registro del cursor
            varResultado.MoveNext
        Wend
        
        Set varResultado = Nothing
        Set varct_ticketxdetalledatosproducto = Nothing
        
        FunGConsulta = True
        Exit Function
    
ErrorManager:
    SubGMuestraError

End Function
Public Function Add(proiIncidentId As Long, proiDatosProductoId As Long, proiDetalleDatosProductoId As Long, Optional proConexion As Variant, Optional sKey As String) As claticketxdetalledatosproducto _
    'create a new object
    Dim objNewMember As claticketxdetalledatosproducto
    Set objNewMember = New claticketxdetalledatosproducto

    'set the properties passed into the method
    objNewMember.proiIncidentId = proiIncidentId
    objNewMember.proiDatosProductoId = proiDatosProductoId
    objNewMember.proiDetalleDatosProductoId = proiDetalleDatosProductoId
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
Public Property Get Item(vntIndexKey As Variant) As claticketxdetalledatosproducto
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
