VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colDatosProductoNumero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claDatosProductoNumero"
Attribute VB_Ext_KEY = "Member0" ,"claDatosProductoNumero"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public proDatosProductoId As String
Public proNumero As String
Public proRegionCode As String
Public proSeleccionados As String

Public proConexion As ADODB.Connection

Public Function Add(proConexion As ADODB.Connection, proFechaAsignacion As String, proNumero As String, _
                    proRegionName As String, proRegionCode As String, proDatosProductoId As String, _
                    proClasificacionDescripcion As String, proTipoLinea As Long, proPublicar As String, Optional sKey As String) As claDatosProductoNumero
                    'proClasificacionDescripcion As String, Agregado por Carlos Castelblanco 2006/07/28
                    
    'create a new object
    Dim objNewMember As claDatosProductoNumero
    Set objNewMember = New claDatosProductoNumero


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proSeleccion = "0"
    objNewMember.proFechaAsignacion = proFechaAsignacion
    objNewMember.proNumero = proNumero
    objNewMember.proRegionName = proRegionName
    objNewMember.proRegionCode = proRegionCode
    objNewMember.proDatosProductoId = proDatosProductoId
    objNewMember.proClasificacionDescripcion = proClasificacionDescripcion 'Agregado por Carlos Castelblanco 2006/07/28
    objNewMember.proTipoLinea = proTipoLinea
    objNewMember.proPublicar = proPublicar
        
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claDatosProductoNumero
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


Public Function MetConsultar(parDetalleDatosProducto As colDetalleDatosProducto, Optional parModificarContador As Boolean = True) As Boolean
    Dim varComando As String
    Dim varComandoClasificacion As String
    Dim varResultadoClasificacion As ADODB.Recordset
    Dim varDatosProductoNumero As claDatosProductoNumero
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

    'Llamar al procedimiento de consulta en la tabla CT_DatosProductoNumero
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_CONSULTARDATOSPRODUCTONUMERO"
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IDATOSPRODUCTOID", adInteger, adParamInput, 0, Me.proDatosProductoId)
    varResultado.Open varCommand
    
    While Not varResultado.EOF
        Set varDatosProductoNumero = New claDatosProductoNumero
        varDatosProductoNumero.proDatosProductoId = Me.proDatosProductoId
        
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varDatosProductoNumero.proRegionCode = ""
        Else
            varDatosProductoNumero.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("chRegionName")) Then
            varDatosProductoNumero.proRegionName = ""
        Else
            varDatosProductoNumero.proRegionName = Trim(varResultado.Fields("chRegionName"))
        End If
        
        If IsNull(varResultado.Fields("vchNumero")) Then
            varDatosProductoNumero.proNumero = "0"
        Else
            varDatosProductoNumero.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        If IsNull(varResultado.Fields("dtFechaAsignacion")) Then
            varDatosProductoNumero.proFechaAsignacion = ""
        Else
            varDatosProductoNumero.proFechaAsignacion = Trim(varResultado.Fields("dtFechaAsignacion"))
        End If
        
        varDatosProductoNumero.proTipoLinea = IIf(IsNull(varResultado!iDetalleDatosProductoId), 0, varResultado!iDetalleDatosProductoId)
        varDatosProductoNumero.proPublicar = varResultado!CHPUBLICAR

        'Comando agregado por Carlos Castelblanco 2006/07/28
        'PLAN DE EJECUCIÓN #1
        varComandoClasificacion = "select distinct                                      " & _
                "    ct.vchclasificacion vchClasificaciones                             " & _
                "From                                                                   " & _
                "    ct_clasificacionnumero c                                           " & _
                "left outer join                                                        " & _
                "    ct_clasificacion ct on c.iclasificacionid = ct.iclasificacionid    " & _
                "where c.vchnumero  = '" & varDatosProductoNumero.proNumero & "'        " & _
                "    and c.chRegionCode = '" & varDatosProductoNumero.proRegionCode & "'"
 
        Set varResultadoClasificacion = New ADODB.Recordset 'Agregado por Carlos Castelblanco 2006/07/28
    
        varResultadoClasificacion.Open varComandoClasificacion, Me.proConexion 'Agregado por Carlos Castelblanco 2006/07/28
    
        'While Agregado por Carlos Castelblanco 2006/07/28
        While Not varResultadoClasificacion.EOF
            If IsNull(varResultadoClasificacion.Fields("vchClasificaciones")) Then
                varDatosProductoNumero.proClasificacionDescripcion = ""
            Else
                If varDatosProductoNumero.proClasificacionDescripcion <> "" Then
                    varDatosProductoNumero.proClasificacionDescripcion = varDatosProductoNumero.proClasificacionDescripcion & ", "
                End If
                varDatosProductoNumero.proClasificacionDescripcion = varDatosProductoNumero.proClasificacionDescripcion & Trim(varResultadoClasificacion.Fields("vchClasificaciones"))
            End If
            varResultadoClasificacion.MoveNext
        Wend
    
        varResultadoClasificacion.Close 'Agregado por Carlos Castelblanco 2006/07/28
        Set varResultadoClasificacion = Nothing
        
        'vardatosproductonumero.proClasificacionDescripcion =
        Me.Add Me.proConexion, _
            varDatosProductoNumero.proFechaAsignacion, _
            varDatosProductoNumero.proNumero, _
            varDatosProductoNumero.proRegionName, _
            varDatosProductoNumero.proRegionCode, _
            varDatosProductoNumero.proDatosProductoId, _
            varDatosProductoNumero.proClasificacionDescripcion, _
            varDatosProductoNumero.proTipoLinea, _
            varDatosProductoNumero.proPublicar
        Dim varIndice As Integer
        varIndice = parDetalleDatosProducto.IndexOf(varDatosProductoNumero.proTipoLinea)
        If varIndice <> 0 And parModificarContador Then
          parDetalleDatosProducto(varIndice).FunGAumentarContador
        End If
        Set varDatosProductoNumero = Nothing
               
        varResultado.MoveNext
    Wend
       
    varResultado.Close
    Set varResultado = Nothing
    
    MetConsultar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetConsultarExistenciaNumero() As Boolean
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim varDatosProductoNumero As claDatosProductoNumero
    Dim i As Integer
    On Error GoTo ErrManager
    'limpia la coleccion
    'Limpia la coleccion
    If Me.Count > 0 Then
        For i = Me.Count To 1 Step -1
            Me.Remove (i)
        Next i
    End If
    
    varComando = "Select    D.iDatosProductoId,                     " & _
                 "          D.chRegionCode,                         " & _
                 "          R.chRegionName,                         " & _
                 "          D.vchNumero,                            " & _
                 "          D.dtFechaAsignacion,                     " & _
                 "          D.IDETALLEDATOSPRODUCTOID," & _
                 "          D.CHPUBLICAR " & _
                 "From      CT_DatosProductoNumero  D,              " & _
                 "          Region                  R               " & _
                 "Where     D.chRegionCode      = '" & Me.proRegionCode & "' " & _
                 " AND      D.vchNumero         = '" & Me.proNumero & "' " & _
                 " And      D.iDatosProductoId  = " & Me.proDatosProductoId & _
                 " AND      R.chRegionCode      = D.chRegionCode    "
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    Set varDatosProductoNumero = New claDatosProductoNumero
    
    While Not varResultado.EOF
        If IsNull(varResultado.Fields("iDatosProductoId")) Then
            varDatosProductoNumero.proDatosProductoId = "0"
        Else
            varDatosProductoNumero.proDatosProductoId = Trim(varResultado.Fields("iDatosProductoId"))
        End If
        
        If IsNull(varResultado.Fields("chRegionCode")) Then
            varDatosProductoNumero.proRegionCode = ""
        Else
            varDatosProductoNumero.proRegionCode = Trim(varResultado.Fields("chRegionCode"))
        End If
        
        If IsNull(varResultado.Fields("chRegionName")) Then
            varDatosProductoNumero.proRegionName = ""
        Else
            varDatosProductoNumero.proRegionName = Trim(varResultado.Fields("chRegionName"))
        End If
        
        If IsNull(varResultado.Fields("vchNumero")) Then
            varDatosProductoNumero.proNumero = "0"
        Else
            varDatosProductoNumero.proNumero = Trim(varResultado.Fields("vchNumero"))
        End If
        
        If IsNull(varResultado.Fields("dtFechaAsignacion")) Then
            varDatosProductoNumero.proFechaAsignacion = ""
        Else
            varDatosProductoNumero.proFechaAsignacion = Trim(varResultado.Fields("dtFechaAsignacion"))
        End If

        varDatosProductoNumero.proTipoLinea = varResultado!iDetalleDatosProductoId
        varDatosProductoNumero.proPublicar = varResultado!CHPUBLICAR
        Me.Add Me.proConexion, _
               varDatosProductoNumero.proFechaAsignacion, _
               varDatosProductoNumero.proNumero, _
               varDatosProductoNumero.proRegionName, _
               varDatosProductoNumero.proRegionCode, _
               varDatosProductoNumero.proDatosProductoId, _
               varDatosProductoNumero.proClasificacionDescripcion, _
               varDatosProductoNumero.proTipoLinea, _
               varDatosProductoNumero.proPublicar
               
        varResultado.MoveNext
    Wend
    
    Set varDatosProductoNumero = Nothing
    Set varResultado = Nothing
    
    MetConsultarExistenciaNumero = True
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


