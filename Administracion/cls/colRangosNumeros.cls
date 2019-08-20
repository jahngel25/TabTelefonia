VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colRangosNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claRangosNumeros"
Attribute VB_Ext_KEY = "Member0" ,"claRangosNumeros"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : colRangosNumerosMódulo de clase
' Fecha  : 12/10/2004 09:20
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Consultar y administrar los posibles rangos de disponibilidad de números
'---------------------------------------------------------------------------------------
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public proRegionCode  As String
Public proEstadoNumero  As String
Public proNumeroMenor  As String
Public proNumeroMayor  As String

Public proConexion As ADODB.Connection



Public Function FunGConsulta() As Boolean
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    Dim varRangosNumeros As claRangosNumeros
   On Error GoTo ErrorManager
        
    Set varComando = New ADODB.Command
    Set varComando.ActiveConnection = Me.proConexion
    varComando.CommandType = adCmdStoredProc
    varComando.CommandText = "CP_DevolverRangosNumeros"
    
    varComando.Parameters.Item("@RegionCode").Value = Me.proRegionCode
    varComando.Parameters.Item("@Estado").Value = Me.proEstadoNumero
    varComando.Parameters.Item("@NumeroMenor").Value = Me.proNumeroMenor
    varComando.Parameters.Item("@NumeroMayor").Value = Me.proNumeroMayor
    varComando.CommandTimeout = 0
    Set varResultado = New ADODB.Recordset
    
    Set varResultado = varComando.Execute
    
    Set varRangosNumeros = New claRangosNumeros
    
   
    While Not varResultado.EOF
        
        varRangosNumeros.proInicio = IIf(IsNull(varResultado.Fields("rango_inicio")), "0", varResultado.Fields("rango_inicio"))
        varRangosNumeros.proFin = IIf(IsNull(varResultado.Fields("rango_fin")), "0", varResultado.Fields("rango_fin"))
        varRangosNumeros.proCuantos = varRangosNumeros.proFin - varRangosNumeros.proInicio
        Me.Add Me.proConexion, _
                varRangosNumeros.proInicio, _
                varRangosNumeros.proFin, _
                varRangosNumeros.proCuantos
        varResultado.MoveNext
    Wend
    
    Set varComando = Nothing
    Set varRangosNumeros = Nothing
    Set varResultado = Nothing
    
    FunGConsulta = True
    Exit Function
ErrorManager:
    SubGMuestraError

End Function


Public Function Add(proConexion As ADODB.Connection, proInicio As String, proFin As String, proCuantos As String, Optional sKey As String) As claRangosNumeros
    'create a new object
    Dim objNewMember As claRangosNumeros
    Set objNewMember = New claRangosNumeros

    Set objNewMember.proConexion = proConexion
    objNewMember.proCuantos = proCuantos
    objNewMember.proFin = proFin
    objNewMember.proInicio = proInicio
    

    'set the properties passed into the method
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claRangosNumeros
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

