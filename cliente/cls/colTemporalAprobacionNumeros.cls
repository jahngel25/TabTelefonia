VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTemporalAprobacionNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************************
'   Colección de Datos para manipular los datos temporales en la aprobación de números
'   Autor: Fernando González
'   Fecha: 16-Nov-2006
'******************************************************************************************************
Option Explicit
Public temUserID As String
Public temRegionCode As String
Public temRegionname As String
Public temNumero As String
Public temEstadonumero As String
Public temDescripcionestado As String
Public temClasificacion As String
Public temUpdateby As String
Public temUseridaprobador As String
Public temFechaaprobacion As String
Public temUpdatedate As String
Public temIncidentid As String
Public temCompanyid As String
Public temCompanyname As String

Public varConexion As ADODB.Connection

Public temSeleccionados As Integer

'local variable to hold collection
Private mCol As Collection

Public Function Add(temConexion As ADODB.Connection _
                    , temUserID As String _
                    , temRegionCode As String _
                    , temRegionname As String _
                    , temNumero As String _
                    , temEstadonumero As String _
                    , temDescripcionestado As String _
                    , temClasificacion As String _
                    , temUpdateby As String _
                    , temUseridaprobador As String _
                    , temFechaaprobacion As String _
                    , temUpdatedate As String _
                    , temIncidentid As String _
                    , temCompanyid As String _
                    , temCompanyname As String _
                    , Optional sKey As String) As claTemporalAprobacionNumeros
    'create a new object
    Dim objNewMember As claTemporalAprobacionNumeros
    Set objNewMember = New claTemporalAprobacionNumeros

    'set the properties passed into the method
    If IsObject(temConexion) Then
        Set objNewMember.temConexion = temConexion
    Else
        objNewMember.temConexion = temConexion
    End If
        
    objNewMember.temUserID = temUserID
     objNewMember.temRegionCode = temRegionCode
     objNewMember.temRegionname = temRegionname
     objNewMember.temNumero = temNumero
     objNewMember.temEstadonumero = temEstadonumero
     objNewMember.temDescripcionestado = temDescripcionestado
     objNewMember.temClasificacion = temClasificacion
     objNewMember.temUpdateby = temUpdateby
     objNewMember.temUseridaprobador = temUseridaprobador
     objNewMember.temFechaaprobacion = temFechaaprobacion
     objNewMember.temUpdatedate = temUpdatedate
     objNewMember.temIncidentid = temIncidentid
     objNewMember.temCompanyid = temCompanyid
     objNewMember.temCompanyname = temCompanyname
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claTemporalAprobacionNumeros
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
    Me.temSeleccionados = 0
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Function MetConsultarTemporalNumeros(UserId As String) As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claTemporalAprobacionNumeros
    On Error GoTo ErrManager
        
    varComando = "exec CP_TempAprobarNumeros 0, '" & UserId & "'"
    
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.varConexion
    
    Set varNumero = New claTemporalAprobacionNumeros
    Set varNumero.temConexion = Me.varConexion
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("chUserId")) Then
            varNumero.temUserID = ""
        Else
            varNumero.temUserID = Trim(varResultado.Fields("chUserId"))
        End If
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varNumero.temRegionCode = ""
        Else
            varNumero.temRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        If IsNull(varResultado.Fields("chregionname")) Then
            varNumero.temRegionname = ""
        Else
            varNumero.temRegionname = Trim(varResultado.Fields("chregionname"))
        End If
        If IsNull(varResultado.Fields("vchnumero")) Then
            varNumero.temNumero = ""
        Else
            varNumero.temNumero = Trim(varResultado.Fields("vchnumero"))
        End If
        If IsNull(varResultado.Fields("chestadonumero")) Then
            varNumero.temEstadonumero = ""
        Else
            varNumero.temEstadonumero = Trim(varResultado.Fields("chestadonumero"))
        End If
        If IsNull(varResultado.Fields("vchdescripcionestado")) Then
            varNumero.temDescripcionestado = ""
        Else
            varNumero.temDescripcionestado = Trim(varResultado.Fields("vchdescripcionestado"))
        End If
        If IsNull(varResultado.Fields("Clasificacion")) Then
            varNumero.temClasificacion = ""
        Else
            varNumero.temClasificacion = Trim(varResultado.Fields("Clasificacion"))
        End If
        If IsNull(varResultado.Fields("chupdateby")) Then
            varNumero.temUpdateby = ""
        Else
            varNumero.temUpdateby = Trim(varResultado.Fields("chupdateby"))
        End If
        If IsNull(varResultado.Fields("chupdateby")) Then
            varNumero.temUpdateby = ""
        Else
            varNumero.temUpdateby = Trim(varResultado.Fields("chupdateby"))
        End If
        If IsNull(varResultado.Fields("chuseridaprobador")) Then
            varNumero.temUseridaprobador = ""
        Else
            varNumero.temUseridaprobador = Trim(varResultado.Fields("chuseridaprobador"))
        End If
        If IsNull(varResultado.Fields("dtfechaaprobacion")) Then
            varNumero.temFechaaprobacion = ""
        Else
            varNumero.temFechaaprobacion = Trim(varResultado.Fields("dtfechaaprobacion"))
        End If
        If IsNull(varResultado.Fields("dtupdatedate")) Then
            varNumero.temUpdatedate = ""
        Else
            varNumero.temUpdatedate = Trim(varResultado.Fields("dtupdatedate"))
        End If
        If IsNull(varResultado.Fields("iincidentid")) Then
            varNumero.temIncidentid = ""
        Else
            varNumero.temIncidentid = Trim(varResultado.Fields("iincidentid"))
        End If
        If IsNull(varResultado.Fields("iCompanyId")) Then
            varNumero.temCompanyid = ""
        Else
            varNumero.temCompanyid = Trim(varResultado.Fields("iCompanyId"))
        End If
        If IsNull(varResultado.Fields("vchcompanyname")) Then
            varNumero.temCompanyname = ""
        Else
            varNumero.temCompanyname = Trim(varResultado.Fields("vchcompanyname"))
        End If
        Me.Add varNumero.temConexion, _
                varNumero.temUserID, _
                varNumero.temRegionCode, _
                varNumero.temRegionname, _
                varNumero.temNumero, _
                varNumero.temEstadonumero, _
                varNumero.temDescripcionestado, _
                varNumero.temClasificacion, _
                varNumero.temUpdateby, _
                varNumero.temUseridaprobador, _
                varNumero.temFechaaprobacion, _
                varNumero.temUpdatedate, _
                varNumero.temIncidentid, _
                varNumero.temCompanyid, _
                varNumero.temCompanyname

        varResultado.MoveNext
    Wend

    
    Set varNumero = Nothing
    Set varResultado = Nothing
    
    MetConsultarTemporalNumeros = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function
Public Function MetInsertarTemporalNumeros(UserId As String, chRegionCode As String, chregionname As String, _
                    vchnumero As String, chestadonumero As String, vchdescripcionestado As String, _
                    Clasificacion As String, chupdateby As String, chuseridaprobador As String, _
                    iincidentid As String, icompanyid As String, vchcompanyname As String) As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claTemporalAprobacionNumeros
    On Error GoTo ErrManager
    
    varComando = "exec CP_TempAprobarNumeros 1, '" & UserId & "','" _
                    & vchnumero & "','" _
                    & chRegionCode & "','" _
                    & chregionname & "','" _
                    & chestadonumero & "','" _
                    & vchdescripcionestado & "','" _
                    & Clasificacion & "','" _
                    & chupdateby & "','" _
                    & chuseridaprobador & "'," _
                    & iincidentid & "," _
                    & icompanyid & ",'" _
                    & vchcompanyname & "'"
    
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.varConexion

    Set varNumero = Nothing
    Set varResultado = Nothing
    
    MetInsertarTemporalNumeros = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetBorrarTemporalNumeros(UserId As String, vchnumero As String) As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claTemporalAprobacionNumeros
    On Error GoTo ErrManager
    
    varComando = "exec CP_TempAprobarNumeros 2, '" & UserId & "','" & vchnumero & "'"
    
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.varConexion

    Set varNumero = Nothing
    Set varResultado = Nothing
    
    MetBorrarTemporalNumeros = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetDefinitivaTemporalNumeros(Operacion As String, UserId As String) As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varNumero As claTemporalAprobacionNumeros
    On Error GoTo ErrManager
    
    varComando = "exec CP_AprobarNumeros " & Operacion & ",'" & UserId & "'"
    
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.varConexion

    Set varNumero = Nothing
    Set varResultado = Nothing
    
    MetDefinitivaTemporalNumeros = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

