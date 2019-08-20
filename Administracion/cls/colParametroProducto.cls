VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colParametroProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claParametroProducto"
Attribute VB_Ext_KEY = "Member0" ,"claParametroProducto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***********************************************************************
'* Descripcion:
'*
'*
'*
'*
'*
'* Parametros:
'*
'*
'*
'*
'*
'*
'*
'**************************************************************************
'**********************************************************************
' MODIFICADO POR :      CARLOS ALBERTO BARRERA
' DESCRIPCION CAMBIO:   Se pasa como parametro la propiedad del id del cliente
' VERSION: 1.0.100
' FECHA: SEPTIEMBRE 7/2009
'****************************************************************

Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proProductNumber As String
Public proCampoPadre As String

'Conexion a la base
Public proConexion As ADODB.Connection

'variable local para capturar el id del cliente
Private variClienteId As Long

Public Function Add(proProductNumber As String, proCampo As String, _
                    proEtiqueta As String, proTipo As String, _
                    proTamaño As Long, proIDInterfase As Boolean, _
                    proPosicionInterfase As Long, proMascara As String, _
                    proObligatorioVenta As Boolean, proObligatorioAtencion As Boolean, _
                    proObligatorioOT As Boolean, _
                    proEditableVenta As Boolean, proEditableAtencion As Boolean, _
                    proEditableOT As Boolean, _
                    proValidarRepetidos As String, _
                    proCampoPadre As String, ProOrden As Long, _
                    proConexion As ADODB.Connection, _
                    Optional proValores As colValoresCampoProducto, _
                    Optional sKey As String) As claParametroProducto
                    
    'create a new object
    Dim objNewMember As claParametroProducto
    Set objNewMember = New claParametroProducto


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proValidarRepetidos = proValidarRepetidos
    objNewMember.proObligatorioOT = IIf(proObligatorioOT, "True", "False")
    objNewMember.proObligatorioAtencion = IIf(proObligatorioAtencion, "True", "False")
    objNewMember.proObligatorioVenta = IIf(proObligatorioVenta, "True", "False")
    
    objNewMember.proEditableOT = IIf(proEditableOT, "True", "False")
    objNewMember.proEditableAtencion = IIf(proEditableAtencion, "True", "False")
    objNewMember.proEditableVenta = IIf(proEditableVenta, "True", "False")
    
    
    objNewMember.proMascara = proMascara
    objNewMember.proPosicionInterfase = proPosicionInterfase
    objNewMember.proIDInterfase = IIf(proIDInterfase, "True", "False")
    objNewMember.proTamaño = proTamaño
    objNewMember.proTipo = proTipo
    objNewMember.proEtiqueta = proEtiqueta
    objNewMember.proCampo = proCampo
    objNewMember.proCampoPadre = proCampoPadre
    objNewMember.ProOrden = ProOrden
    objNewMember.proProductNumber = proProductNumber
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

     'set the properties passed into the method
    If Not IsMissing(proValores) Then
        If IsObject(proValores) Then
            Set objNewMember.proValores = proValores
        Else
            objNewMember.proValores = proValores
        End If
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claParametroProducto
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
'"--------------------------------------------------------------------------------------------
'"<summary> Método que consulta los valores de acuerdo con el id de un cliente
'"</summary>
'"<remark>MetConsultarValores(Optional pariClienteId As Long = 0)<para><version>1.0.100</version><cambio>Se adiciono para pasar como parametro el id del cliente</cambio><fecha>30/06/2009</fecha></remark>
'"</remark>
'"<param name = "pariClienteId">Id del cliente</param>
'"<returns> Booleano
'"<example> <code>if (MetConsultarValores(proiClienteId)) then </code></example>
'"<history>
'"</history>
'"----------------------------------------------------------------------------------------------
Public Function metConsultarxProducto(Optional pariClienteId As Long = 0) As Boolean '1.0.100 Se pasa como parametro el id del cliente
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varParametro As claParametroProducto
    On Error GoTo ErrManager
            'Elimina todos los elementos de la colección
    While Me.Count
            Me.Remove 1
    Wend
    varComando = "SELECT    chProductNumber,        " & _
                                       "chCampo, " & _
                                       "vchEtiqueta, " & _
                                       "chTipo, " & _
                                       "iTamaño, " & _
                                       "bIDInterfase, " & _
                                       "iPosicionInterfase, " & _
                                       "vchMascara, " & _
                                       "bObligatorioVenta, " & _
                                       "bObligatorioAtencion, " & _
                                       "bObligatorioOT, " & _
                                       "bEditableVenta, " & _
                                       "bEditableAtencion, " & _
                                       "bEditableOT, " & _
                                       "cValidarRepetidos, " & _
                                       "chCampoPadre, " & _
                                       "iOrden " & _
                 "FROM      CT_ParametrosXProducto  " & _
                 "WHERE     chProductNumber = '" & Me.proProductNumber & "' " & _
                 " ORDER BY iOrden"
    
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.proConexion
    
    
    While Not varResultado.EOF
        Set varParametro = New claParametroProducto
        varParametro.proProductNumber = IIf(IsNull(varResultado("chProductNumber")), "", varResultado.Fields("chProductNumber"))
        varParametro.proCampo = IIf(IsNull(varResultado("chCampo")), "", varResultado.Fields("chCampo"))
        varParametro.proEtiqueta = IIf(IsNull(varResultado("vchEtiqueta")), "", varResultado.Fields("vchEtiqueta"))
        varParametro.proTipo = IIf(IsNull(varResultado("chTipo")), "", varResultado.Fields("chTipo"))
        varParametro.proTamaño = IIf(IsNull(varResultado("iTamaño")), "", varResultado.Fields("iTamaño"))
        varParametro.proIDInterfase = IIf(IsNull(varResultado("bIDInterfase")), "False", varResultado.Fields("bIDInterfase"))
        varParametro.proPosicionInterfase = IIf(IsNull(varResultado("iPosicionInterfase")), "", varResultado.Fields("iPosicionInterfase"))
        varParametro.proMascara = IIf(IsNull(varResultado("vchMascara")), "", varResultado.Fields("vchMascara"))
        varParametro.proObligatorioVenta = IIf(IsNull(varResultado("bObligatorioVenta")), "False", varResultado.Fields("bObligatorioVenta"))
        varParametro.proObligatorioAtencion = IIf(IsNull(varResultado("bObligatorioAtencion")), "False", varResultado.Fields("bObligatorioAtencion"))
        varParametro.proObligatorioOT = IIf(IsNull(varResultado("bObligatorioOT")), "False", varResultado.Fields("bObligatorioOT"))
        varParametro.proEditableVenta = IIf(IsNull(varResultado("bEditableVenta")), "False", varResultado.Fields("bEditableVenta"))
        varParametro.proEditableAtencion = IIf(IsNull(varResultado("bEditableAtencion")), "False", varResultado.Fields("bEditableAtencion"))
        varParametro.proEditableOT = IIf(IsNull(varResultado("bEditableOT")), "False", varResultado.Fields("bEditableOT"))
        varParametro.proValidarRepetidos = IIf(IsNull(varResultado("cValidarRepetidos")), "", varResultado.Fields("cValidarRepetidos"))
        varParametro.proCampoPadre = IIf(IsNull(varResultado("chCampoPadre")), "", varResultado.Fields("chCampoPadre"))
        varParametro.ProOrden = IIf(IsNull(varResultado("iOrden")), 0, varResultado.Fields("iOrden"))
        
        'Cargar los valores del parametro en el caso de ser una lista
        If varParametro.proTipo = "L" And Trim(varParametro.proCampoPadre) = "" Then
            Set varParametro.proConexion = Me.proConexion
            If Not varParametro.MetConsultarValores(pariClienteId) Then '1.0.100 Se pasa como parametro el id del cliente
                MsgBox "Error al consultar los valores.", vbCritical, App.Title
            End If
        Else
            Set varParametro.proValores = New colValoresCampoProducto
        End If
        
        Me.Add varParametro.proProductNumber, _
                varParametro.proCampo, _
                varParametro.proEtiqueta, _
                varParametro.proTipo, _
                varParametro.proTamaño, _
                varParametro.proIDInterfase, _
                varParametro.proPosicionInterfase, _
                varParametro.proMascara, _
                varParametro.proObligatorioVenta, _
                varParametro.proObligatorioAtencion, _
                varParametro.proObligatorioOT, _
                varParametro.proEditableVenta, _
                varParametro.proEditableAtencion, _
                varParametro.proEditableOT, _
                varParametro.proValidarRepetidos, _
                varParametro.proCampoPadre, _
                varParametro.ProOrden, _
                Me.proConexion, _
                varParametro.proValores
                
        Set varParametro = Nothing
        varResultado.MoveNext
    Wend
    
    Set varParametro = Nothing
    Set varResultado = Nothing
    
    metConsultarxProducto = True
    Exit Function

ErrManager:
    SubGMuestraError

End Function
Public Function metConsultarxProductoyCampo() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varParametro As claParametroProducto
    
    On Error GoTo ErrManager
            'Elimina todos los elementos de la colección
    While Me.Count
            Me.Remove 1
    Wend
    varComando = "SELECT    chProductNumber,        " & _
                                       "chCampo, " & _
                                       "vchEtiqueta, " & _
                                       "chTipo, " & _
                                       "iTamaño, " & _
                                       "bIDInterfase, " & _
                                       "iPosicionInterfase, " & _
                                       "vchMascara, " & _
                                       "bObligatorioVenta, " & _
                                       "bObligatorioAtencion, " & _
                                       "bObligatorioOT, " & _
                                       "bEditableVenta, " & _
                                       "bEditableAtencion, " & _
                                       "bEditableOT, " & _
                                       "cValidarRepetidos, " & _
                                       "chCampoPadre, " & _
                                       "iOrden " & _
                 "FROM      CT_ParametrosXProducto  " & _
                 "WHERE     chProductNumber = '" & Me.proProductNumber & "' " & _
                 " AND chCampoPadre='" & proCampoPadre & "'" & _
                 " ORDER BY iOrden"
    
    Set varResultado = New ADODB.Recordset
    varResultado.Open varComando, Me.proConexion
    
    While Not varResultado.EOF
        Set varParametro = New claParametroProducto
    
        varParametro.proProductNumber = IIf(IsNull(varResultado("chProductNumber")), "", varResultado.Fields("chProductNumber"))
        varParametro.proCampo = IIf(IsNull(varResultado("chCampo")), "", varResultado.Fields("chCampo"))
        varParametro.proEtiqueta = IIf(IsNull(varResultado("vchEtiqueta")), "", varResultado.Fields("vchEtiqueta"))
        varParametro.proTipo = IIf(IsNull(varResultado("chTipo")), "", varResultado.Fields("chTipo"))
        varParametro.proTamaño = IIf(IsNull(varResultado("iTamaño")), "", varResultado.Fields("iTamaño"))
        varParametro.proIDInterfase = IIf(IsNull(varResultado("bIDInterfase")), "False", varResultado.Fields("bIDInterfase"))
        varParametro.proPosicionInterfase = IIf(IsNull(varResultado("iPosicionInterfase")), "", varResultado.Fields("iPosicionInterfase"))
        varParametro.proMascara = IIf(IsNull(varResultado("vchMascara")), "", varResultado.Fields("vchMascara"))
        varParametro.proObligatorioVenta = IIf(IsNull(varResultado("bObligatorioVenta")), "False", varResultado.Fields("bObligatorioVenta"))
        varParametro.proObligatorioAtencion = IIf(IsNull(varResultado("bObligatorioAtencion")), "False", varResultado.Fields("bObligatorioAtencion"))
        varParametro.proObligatorioOT = IIf(IsNull(varResultado("bObligatorioOT")), "False", varResultado.Fields("bObligatorioOT"))
        varParametro.proEditableVenta = IIf(IsNull(varResultado("beditableVenta")), "False", varResultado.Fields("bEditableVenta"))
        varParametro.proEditableAtencion = IIf(IsNull(varResultado("beditableAtencion")), "False", varResultado.Fields("bEditableAtencion"))
        varParametro.proEditableOT = IIf(IsNull(varResultado("beditableOT")), "False", varResultado.Fields("bEditableOT"))
        varParametro.proValidarRepetidos = IIf(IsNull(varResultado("cValidarRepetidos")), "", varResultado.Fields("cValidarRepetidos"))
        varParametro.proCampoPadre = IIf(IsNull(varResultado("chCampoPadre")), "", varResultado.Fields("chCampoPadre"))
        varParametro.ProOrden = IIf(IsNull(varResultado("iOrden")), 0, varResultado.Fields("iOrden"))
        
        'Cargar los valores del parametro en el caso de ser una lista
        If varParametro.proTipo = "L" And Trim(varParametro.proCampoPadre) = "" Then
            Set varParametro.proConexion = Me.proConexion
            If Not varParametro.MetConsultarValores(variClienteId) Then '1.0.100
                MsgBox "Error al consultar los valores.", vbCritical, App.Title
            End If
        Else
            Set varParametro.proValores = New colValoresCampoProducto
        End If
        
        
        Me.Add varParametro.proProductNumber, _
                varParametro.proCampo, _
                varParametro.proEtiqueta, _
                varParametro.proTipo, _
                varParametro.proTamaño, _
                varParametro.proIDInterfase, _
                varParametro.proPosicionInterfase, _
                varParametro.proMascara, _
                varParametro.proObligatorioVenta, _
                varParametro.proObligatorioAtencion, _
                varParametro.proObligatorioOT, _
                varParametro.proEditableVenta, _
                varParametro.proEditableAtencion, _
                varParametro.proEditableOT, _
                varParametro.proValidarRepetidos, _
                varParametro.proCampoPadre, _
                varParametro.ProOrden, _
                Me.proConexion, _
                varParametro.proValores
                
        Set varParametro = Nothing
        
        varResultado.MoveNext
    Wend
    
    Set varParametro = Nothing
    Set varResultado = Nothing
    
    metConsultarxProductoyCampo = True
    Exit Function

ErrManager:
    SubGMuestraError

End Function


