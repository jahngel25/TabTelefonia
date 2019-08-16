VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTipoLineaEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(proNovedadDetalleDatosProductoId As String, proUser1 As String, _
                    proUser15 As String, proContadorNumeros As Integer, proNovedad As Boolean, Optional sKey As String) As claTipoLineaEdicion
    Dim objNewMember As New claTipoLineaEdicion
    'set the properties passed into the method
    objNewMember.proNovedadDetalleDatosProductoId = proNovedadDetalleDatosProductoId
    objNewMember.proUser1 = proUser1
    objNewMember.proUser15 = proUser15
    objNewMember.proContadorNumeros = proContadorNumeros
    objNewMember.proNovedad = proNovedad
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As claTipoLineaEdicion
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


Public Property Get IndexOf(vntValue As Long) As Long
    Dim i As Integer
    Dim j As Integer
    j = 0
    i = 1
    While i <= mCol.Count And j = 0
        Dim vardetalle As claTipoLineaEdicion
        Set vardetalle = mCol(i)
        If vardetalle.proNovedadDetalleDatosProductoId = vntValue Then
            j = i
        End If
        i = i + 1
    Wend
    IndexOf = j
End Property




