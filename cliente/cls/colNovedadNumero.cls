VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colNovedadNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claNovedadNumero"
Attribute VB_Ext_KEY = "Member0" ,"claNovedadNumero"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proConexion As ADODB.Connection

Public proDatosProductoId As String
Public proSeleccionados As Integer

'Datos consulta
Public proNumero As String
Public proRegionCode As String
Public proIncidentId As String
Public proNovedadNumeroId As String 'parametro de salida
Public proTipoNovedadId As String 'parametro de salida
'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proFechaLiberacion As String, _
                    proFechaReserva As String, proTipoNovedadId As String, _
                    proIncidentId As String, proNumero As String, _
                    proDatosProductoId As String, _
                    proRegionName As String, proRegionCode As String, _
                    proNovedadNumeroId As String, _
                    proClasificacionDescripcion As String, _
                    proTipoLinea As Long, _
                    proTipoLineaAnterior As Long, _
                    proPublicar As String, _
                    parasociaNovedad As String, _
                    Optional sKey As String) As claNovedadNumero
                    'proClasificacionDescripcion As String, Agregado por Carlos Castelblanco 2006/07/28
                    
    'create a new object
    Dim objNewMember As claNovedadNumero
    Set objNewMember = New claNovedadNumero


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proSeleccion = "0"
    objNewMember.proFechaLiberacion = proFechaLiberacion
    objNewMember.proFechaReserva = proFechaReserva
    objNewMember.proTipoNovedadId = proTipoNovedadId
    objNewMember.proIncidentId = proIncidentId
    objNewMember.proNumero = proNumero
    objNewMember.proDatosProductoId = proDatosProductoId
    objNewMember.proRegionName = proRegionName
    objNewMember.proRegionCode = proRegionCode
    objNewMember.proNovedadNumeroId = proNovedadNumeroId
    objNewMember.proClasificacionDescripcion = proClasificacionDescripcion 'Agregado por Carlos Castelblanco 2006/07/28
    objNewMember.proTipoLinea = proTipoLinea
    objNewMember.proPublicar = proPublicar
    objNewMember.proTipoLineaAnterior = proTipoLineaAnterior
    objNewMember.proAsociaNovedad = parasociaNovedad
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claNovedadNumero
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
    Me.proSeleccionados = 0
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
    varComando = "SELECT   iNovedadNumeroId,  ISNULL(COUNT(iNovedadNumeroId), 0) AS Cuantos, iTipoNovedadId " & _
                            " From CT_NOVEDADNUMEROS " & _
                            " WHERE     " & _
                            "  iDatosProductoId =  " & Me.proDatosProductoId & _
                            " AND iIncidentId      = " & Me.proIncidentId & _
                            " AND chRegionCode =  '" & Me.proRegionCode & _
                            "' AND vchNumero =  '" & Me.proNumero & "' group by  iNovedadNumeroId, iTipoNovedadId "
     
     varResultado.Open varComando, Me.proConexion
    If varResultado.EOF Then
        MetConsultarxServicio = False
        Exit Function
    Else
        If varResultado.Fields("Cuantos") = 0 Then
            MetConsultarxServicio = False
            Exit Function
        Else
            Me.proNovedadNumeroId = varResultado.Fields("iNovedadNumeroId")
            Me.proTipoNovedadId = varResultado.Fields("iTipoNovedadId")
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
Public Function MetConsultar(parDetalleDatosProducto As colDetalleDatosProducto, Optional parModificarContador As Boolean = True) As Boolean
    Dim varComando As String
    Dim varComandoClasificacion As String
    Dim varResultadoClasificacion  As ADODB.Recordset
    Dim varNovedadNumero As claNovedadNumero
    Dim i As Integer
    On Error GoTo ErrManager
    Dim varResultado As New ADODB.Recordset
    Dim varCommand As New ADODB.Command
    
    'Limpia la coleccion
    If Me.Count > 0 Then
        For i = Me.Count To 1 Step -1
            Me.Remove (i)
        Next i
    End If
    
    'Llamar al procedimiento de consulta en la tabla CT_NOVEDADNUMEROS
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_CONSULTARNOVEDADNUMEROS"
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IDATOSPRODUCTOID", adInteger, adParamInput, 0, proDatosProductoId)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IINCIDENTID", adInteger, adParamInput, 0, proIncidentId)
    varResultado.Open varCommand

    While Not varResultado.EOF
        Set varNovedadNumero = New claNovedadNumero
        
        If IsNull(varResultado.Fields("iNovedadNumeroId")) Then
            varNovedadNumero.proNovedadNumeroId = ""
        Else
            varNovedadNumero.proNovedadNumeroId = Trim(varResultado.Fields("iNovedadNumeroId"))
        End If
        
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varNovedadNumero.proRegionCode = ""
        Else
            varNovedadNumero.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("chRegionName")) Then
            varNovedadNumero.proRegionName = ""
        Else
            varNovedadNumero.proRegionName = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("vchNumero")) Then
            varNovedadNumero.proNumero = ""
        Else
            varNovedadNumero.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varNovedadNumero.proDatosProductoId = ""
        Else
            varNovedadNumero.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("iIncidentId")) Then
            varNovedadNumero.proIncidentId = ""
        Else
            varNovedadNumero.proIncidentId = Trim(varResultado.Fields("iIncidentId"))
        End If
        
        If IsNull(varResultado.Fields("iTipoNovedadId")) Then
            varNovedadNumero.proTipoNovedadId = ""
        Else
            varNovedadNumero.proTipoNovedadId = Trim(varResultado.Fields("iTipoNovedadId"))
        End If
        
        If IsNull(varResultado.Fields("dtFechaReserva")) Then
            varNovedadNumero.proFechaReserva = ""
        Else
            If Trim(varResultado.Fields("dtFechaReserva")) = "01/01/1900" Then
                varNovedadNumero.proFechaReserva = ""
            Else
                varNovedadNumero.proFechaReserva = Trim(varResultado.Fields("dtFechaReserva"))
            End If
        End If

        If IsNull(varResultado.Fields("dtFechaLiberacion")) Then
            varNovedadNumero.proFechaLiberacion = ""
        Else
            If Trim(varResultado.Fields("dtFechaLiberacion")) = "01/01/1900" Then
                varNovedadNumero.proFechaLiberacion = ""
            Else
                varNovedadNumero.proFechaLiberacion = Trim(varResultado.Fields("dtFechaLiberacion"))
            End If
        End If
        varNovedadNumero.proAsociaNovedad = varResultado!chAsocianovedad
        varNovedadNumero.proTipoLinea = varResultado!INOVEDADDETALLEDATOSPRODUCTOID
        varNovedadNumero.proTipoLineaAnterior = IIf(IsNull(varResultado!IDETALLEDATOSPRODUCTOANTERIORID), 0, varResultado!IDETALLEDATOSPRODUCTOANTERIORID)
        varNovedadNumero.proPublicar = varResultado!CHPUBLICAR
        If varNovedadNumero.proTipoLineaAnterior <> 0 And parModificarContador Then
            Dim varIndice As Integer
            varIndice = parDetalleDatosProducto.IndexOf(varNovedadNumero.proTipoLineaAnterior)
            parDetalleDatosProducto.Item(varIndice).FunGDisminuirContador
        End If
        'Comando agregado por Carlos Castelblanco 2006/07/28
        'PLAN DE EJECUCIÓN # 1
        varComandoClasificacion = "select distinct                                          " & _
                    "    ct.vchclasificacion vchClasificaciones                             " & _
                    "From                                                                   " & _
                    "    ct_clasificacionnumero c                                           " & _
                    "left outer join                                                        " & _
                    "    ct_clasificacion ct on c.iclasificacionid = ct.iclasificacionid    " & _
                    "where c.vchnumero  = '" & varNovedadNumero.proNumero & "'              " & _
                    "    and c.chRegionCode = '" & varNovedadNumero.proRegionCode & "'      "
        
        Set varResultadoClasificacion = New ADODB.Recordset 'Agregado por Carlos Castelblanco 2006/07/28
        
        varResultadoClasificacion.Open varComandoClasificacion, Me.proConexion 'Agregado por Carlos Castelblanco 2006/07/28
        
        'Set varNovedadNumero = New claNovedadNumero 'Agregado por Carlos Castelblanco 2006/07/28
        
        'While Agregado por Carlos Castelblanco 2006/07/28
        While Not varResultadoClasificacion.EOF
            If IsNull(varResultadoClasificacion.Fields("vchClasificaciones")) Then
                varNovedadNumero.proClasificacionDescripcion = ""
            Else
                If varNovedadNumero.proClasificacionDescripcion <> "" Then
                    varNovedadNumero.proClasificacionDescripcion = varNovedadNumero.proClasificacionDescripcion & ", "
                End If
                varNovedadNumero.proClasificacionDescripcion = varNovedadNumero.proClasificacionDescripcion & Trim(varResultadoClasificacion.Fields("vchClasificaciones"))
            End If
            varResultadoClasificacion.MoveNext
        Wend
        
        varResultadoClasificacion.Close 'Agregado por Carlos Castelblanco 2006/07/28
        Set varResultadoClasificacion = Nothing
       
        Me.Add Me.proConexion, _
               varNovedadNumero.proFechaLiberacion, _
               varNovedadNumero.proFechaReserva, _
               varNovedadNumero.proTipoNovedadId, _
               varNovedadNumero.proIncidentId, _
               varNovedadNumero.proNumero, _
               varNovedadNumero.proDatosProductoId, _
               varNovedadNumero.proRegionName, _
               varNovedadNumero.proRegionCode, _
               varNovedadNumero.proNovedadNumeroId, _
               varNovedadNumero.proClasificacionDescripcion, _
               varNovedadNumero.proTipoLinea, _
               varNovedadNumero.proTipoLineaAnterior, _
               varNovedadNumero.proPublicar, _
               varNovedadNumero.proAsociaNovedad
        Set varNovedadNumero = Nothing
               
        varResultado.MoveNext
    Wend
    
    Set varResultado = Nothing
    Set varNovedadNumero = Nothing
    MetConsultar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


