VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proUsuario As String
Public proIncidentId As String
Public proIncidentCategory As iCategoriaIncidente
Public proIncidentType As iTipoIncidente
Public proStatusId As iEstado
Public proOTId As String
Public proStatusOTId As iEstado
Public proCompanyId As String
Public proDatosProductoId As String
Public proProductId As String

'Propiedad que indica que al usuario la validación del proceso siempre le permitirá full access
Public proNoValidar As Boolean

'Propiedad de tipo de acceso
Public proAcceso As iTipoAcceso

'Indica las categorias existentes de lead
Public Enum iCategoriaIncidente
        IncidenteAtencion = 2
        IncidenteVenta = 3
        IncidenteOT = 1
End Enum

'Indica los estados posibles para los incidentes
Public Enum iEstado
        'Estados para la Venta
        EstadoGeneracionContactos = 101813
        EstadoVisita = 102881
        EstadoSolicitudFactibilidad = 101606
        EstadoAnalisisComercial = 102948
        EstadoPresentacionPropuesta = 101610
        EstadoAceptacionFormal = 103675
        EstadoAnalisisFinanciero = 103854
        EstadoFirmaContrato = 102059
        EstadoCerrado = 104
        EstadoCancelada = 103679
        EstadoSuspendida = 102158
        
        'Estados para la Atención
        EstadoRecibido = 102983
        EstadoAtencion = 101598
        EstadoDiagnostico = 102990
        EstadoSolucionCliente = 101662
        EstadoCancelado = 103520
        
        'Estados para la OT
        EstadoGenerada = 101463
        EstadoEnEjecucion = 102380
        EstadoTerminacion = 101601
        EstadoSuspension = 101456
        EstadoCerrada = 101465
        EstadoReclamacion = 103884
End Enum

'Indica los tipos de atención
Public Enum iTipoIncidente
        'Tipos para Atenciones
        TipoReporteTecnico = 102935
        TipoProblemaFacturacion = 102987
        TipoProblemaVentas = 103877
        TipoConsultaClientes = 103883
        TipoReconexion = 101462
        TipoDesconexionTemporal = 101457
        TipoDesinstalacion = 103687
        TipoTraslado = 101460
        TipoMantenimiento = 101458
        TipoCambioVelocidad = 103688
        TipoPruebaAmpliacionCanal = 103960
        TipoRenovacionContrato = 103894
        TipoCambioTipoAcceso = 103919
        TipoCambioEquipos = 103920
        TipoPrueba5Dias = 103961
        TipoCambioServicio = 103962
        TipoRevisionFacturacion = 104015
        TipoReconexionPorPago = 104598
        
        'Tipos para OT
        TipoFactibilidad = 101480
        TipoFactibilidadRegional = 103963
End Enum

'Propiedad que indica el tipo de acceso a la facturación
Public Enum iTipoAcceso
        'Cuando hay acceso total todos los conceptos pueden ser insertados
        AccesoTotal = 0
        'Cuando hay acceso restringido sólo algunos conceptos pueden ser empleados
        'y sólo algunos usuarios pueden usarlos
        AccesoRestringido = 1
        'Cuando hay acceso denegado, nadie puede ingresar ningún concepto
        AccesoDenegado = 2
End Enum

'Conexion a la base de datos
Public proConexion As ADODB.Connection

'Colecciones que indican quienes y que es posible usar
Public proColUsuariosPermitidos As colUsuario
'Public proColConceptosPermitidos As colCobro
'****************************************************************
' El proceso que verifica esta clase es el siguiente:
'
'   - Luego de que el Lead está cerrado, el núcleo de facturación sólo puede ser modificado
'   para ingresar la obra civil por parte de provisioning
'
'   - Para el caso de atenciones, Contratos revisa la atención y luego de que cerrada no
'      debe poderse modificar para los siguientes tipos:
'           -  Renovación de Contrato
'           -  Revisión de Facturación
'           -  Desinstalación
'
'   - Para los tipos de Atención:
'           -  Cambio de configuración
'           -  Traslado
'           -  Reconexión
'           -  Suspensión
'           -  Mantenimiento
'           -  Cambio de Velocidad
'****************************************************************


Public Function MetValidaPermisos() As Boolean

Dim varCuenta As Integer
Dim varEncontro As Boolean
On Error GoTo ErrorManager

        Screen.MousePointer = vbHourglass

        'acceso total por defecto
        Me.proAcceso = AccesoTotal
        MetValidaPermisos = True

        'Verifica si el usuario es de los que saltan la validación
        If Me.proNoValidar Then
                Screen.MousePointer = vbDefault
                Exit Function
        End If

        'Consulta los datos del incidente para tomar decisiones
        If Me.MetConsultaDatosIncidente = False Then
                MsgBox "No fue posible consultar los datos necesarios para validar el proceso", vbInformation, App.Title
                'Deniega el acceso
                Me.proAcceso = AccesoDenegado
                Exit Function
        End If


        'Verifica que tipo de incidente es
        If Me.proIncidentCategory = IncidenteVenta Then
                If Me.proStatusId = EstadoCancelada Then
                    MsgBox "Su usuario no puede modificar la información en este punto del proceso.", vbInformation, App.Title
                    Me.proAcceso = AccesoDenegado
                Else
                    If Me.proStatusId = EstadoCerrado Then
                            'Si la OT no está cerrada, es posible modificar la facturacion
                            If Me.proStatusOTId <> EstadoCerrada And Me.proStatusOTId <> EstadoCancelada Then
                                'Evaluar que se debe hacer
                            Else 'Si la OT Esta cerrada no se puede modificar el incidente
                                MsgBox "Su usuario no puede modificar la información en este punto del proceso.", vbInformation, App.Title
                                    Me.proAcceso = AccesoDenegado
                            End If
                    End If
                End If
        ElseIf Me.proIncidentCategory = IncidenteAtencion Then
                If Me.proStatusId = EstadoCancelado Then
                    MsgBox "Su usuario no puede modificar la información en este punto del proceso.", vbInformation, App.Title
                    Me.proAcceso = AccesoDenegado
                Else
                    If Me.proStatusOTId <> 0 Then
                        If Me.proIncidentType = TipoRenovacionContrato Or Me.proIncidentType = TipoRevisionFacturacion Or Me.proIncidentType = TipoDesinstalacion Then
                                If Me.proStatusId = EstadoCerrado Then
                                    MsgBox "No es posible modificar la información en este punto del proceso. Debe indicar un nuevo asunto.", vbInformation, App.Title
                                    Me.proAcceso = AccesoDenegado
                                End If
                        Else
                                'Si la OT no está cerrada, es posible modificar la facturacion
                                If Me.proStatusOTId <> EstadoCerrada And Me.proStatusId <> EstadoCancelada Then
                                    'Evaluar que se debe hacer
                                Else 'Si la OT Esta cerrada no se puede modificar el incidente
                                        MsgBox "No es posible modificar la información en este punto del proceso. Debe indicar un nuevo asunto.", vbInformation, App.Title
                                        Me.proAcceso = AccesoDenegado
                                End If 'Fin de estado de la OT
                        End If 'FIn de incidentes de renovacion, revision, desinstalacion
                    ElseIf Me.proStatusOTId = 0 Then ' FNP 20060914
                    If Me.proStatusId = EstadoCerrado Then
                        If MetConsultaGeneracionOT = True Then
                            MsgBox "Su usuario no puede modificar la información en este punto del proceso.", vbInformation, App.Title
                            Me.proAcceso = AccesoDenegado
                        End If
                    End If
                    End If
            End If
        End If ' Fin de categorias de incidente

        Screen.MousePointer = vbDefault

        MetValidaPermisos = True
        Exit Function

ErrorManager:
        Screen.MousePointer = vbDefault
        SubGMuestraError
End Function
Function MetConsultaGeneracionOT() As Boolean
' Objetivo: Creada para validar que las Atenciones tengan definidas la generación automatica de OT
' Fecha : 2006/09/14
' Autor: Fredhling Nisperuza Perneth
Dim varResultado As ADODB.Recordset

On Error GoTo ErrorManager
        'Busca si la atención genera una OT
        Set varResultado = New ADODB.Recordset
        varResultado.Open "SELECT 1 " & _
                          " FROM ct_herenciaonyx H " & _
                          " WHERE H.iStatusId = 104 AND H.iIncidentTypeId = " & Me.proIncidentType, Me.proConexion
        
        If varResultado.EOF = False Then
                MetConsultaGeneracionOT = False
        Else
                MetConsultaGeneracionOT = True
        End If
        Exit Function
ErrorManager:
        SubGMuestraError

End Function
Function MetConsultaDatosIncidente() As Boolean
    Dim varTipoOTFactibilidad As iTipoIncidente
    Dim varTipoOTFactibilidadRegional As iTipoIncidente
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrorManager

        'Busca el tipo, el estado y la categoria del incidente que esta permitiendo editar la
        'facturacion
        Set varResultado = New ADODB.Recordset
        varResultado.Open "SELECT I.iStatusId, I.iIncidentCategory, I.iIncidentTypeId " & _
                                    " FROM Incident I " & _
                                    " WHERE I.iIncidentId = " & Me.proIncidentId, Me.proConexion
        
        If varResultado.EOF = False Then
                Me.proStatusId = varResultado("iStatusId").Value
                Me.proIncidentCategory = varResultado("iIncidentCategory").Value
                Me.proIncidentType = varResultado("iIncidentTypeId").Value
        End If
        
       
        'Busca la OT y el estado de la OT
        Me.proStatusOTId = 0
        Me.proOTId = ""
        
        varTipoOTFactibilidad = TipoFactibilidad
        varTipoOTFactibilidadRegional = TipoFactibilidadRegional
        
        Set varResultado = New ADODB.Recordset
        varResultado.Open "SELECT I.iStatusId, I.iIncidentId " & _
                                    " FROM Incident I " & _
                                    " WHERE I.vchUser7 = " & Me.proIncidentId & _
                                    "And    I.iIncidentCategory = 1 " & _
                                    " AND iIncidentTypeId not in (" & varTipoOTFactibilidad & ", " & _
                                     varTipoOTFactibilidadRegional & ") " & _
                                     " AND tiRecordStatus = 1  Order by I.iIncidentId Desc ", Me.proConexion
        
        If varResultado.EOF = False Then
                Me.proStatusOTId = varResultado("iStatusId").Value
                Me.proOTId = varResultado("iIncidentId").Value
        End If
        
        MetConsultaDatosIncidente = True
        Exit Function
ErrorManager:
        SubGMuestraError
End Function

Public Function MetValidarAsunto() As Boolean
'**************************'***********************************************************
'   OBJETIVO:     valida que el asunto digitado no exista en otra facturacion, sea
'                 un asunto de ventas o de atencion, que pertenezca al usuario
'                 activo
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    On Error GoTo ErrManager
    
    If MetValidarExistenciaAsunto() Then
        If MetValidarClienteAsunto() Then
            If MetValidarCategoriaAsunto() Then
                If Not MetValidarOTCerrada() Then
                    MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
                    MetValidarAsunto = False
                    Exit Function
                End If
            Else
                MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
                MetValidarAsunto = False
                Exit Function
            End If
        Else
            MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
            MetValidarAsunto = False
            Exit Function
        End If
    Else
        MsgBox "No se puede editar la información del Asunto.", vbInformation, App.Title
        MetValidarAsunto = False
        Exit Function
    End If
    
    MetValidarAsunto = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Function MetValidarExistenciaAsunto() As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto digitado ya fue utilizado en otro Tab de Voz
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 16/03/2004
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Verificar si el incidente ya esta ligado a algun tab de voz
    varComando = "SELECT COUNT(1) " & _
                 "FROM   CT_DatosProductoIncident " & _
                 "WHERE  iIncidentId = " & Me.proIncidentId
            
    varResultado.Open varComando, Me.proConexion
    
    If varResultado.EOF Then
        MetValidarExistenciaAsunto = True
    Else
        If IsNull(varResultado.Fields(0)) Then
             MetValidarExistenciaAsunto = True
        Else
            If varResultado.Fields(0) <> 0 Then
                MsgBox "El asunto digitado ya fue utilizado en otro TAB de Datos por Servicio. ", vbInformation, App.Title
                MetValidarExistenciaAsunto = False
            Else
                MetValidarExistenciaAsunto = True
            End If
        End If
    End If
    
    Set varResultado = Nothing
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Function MetValidarClienteAsunto() As Boolean
        
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto pertenece al cliente actual
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 16/03/2004
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
        
    'Verificar si el asunto pertenece al cliente actual y cual es la categoria del incidente
    varComando = "SELECT    iOwnerId      " & _
                 "FROM      Incident      " & _
                 "WHERE     iIncidentId = " & Me.proIncidentId
            
    varResultado.Open varComando, Me.proConexion
    
    If varResultado.EOF Then
        MsgBox "El Asunto Digitado no existe. ", vbCritical, App.Title
        MetValidarClienteAsunto = False
    Else
        If IsNull(varResultado.Fields("iOwnerId")) Then
            MsgBox "El Asunto No tiene Cliente Asociado. ", vbCritical, App.Title
            MetValidarClienteAsunto = False
        Else
            If varResultado.Fields("iOwnerId") <> CLng(Trim(Me.proCompanyId)) Then
                MsgBox "El Asunto no pertece a este Cliente. ", vbCritical, App.Title
                MetValidarClienteAsunto = False
            Else
                 MetValidarClienteAsunto = True
            End If
        End If
    End If
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Function MetValidarCategoriaAsunto() As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Validar la Categoria del asunto digitado
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    Dim Producto As Double
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
        
    varComando = "SELECT    iIncidentCategory   " & _
                 "FROM      Incident            " & _
                 "WHERE     iIncidentId =       " & Me.proIncidentId
            
    varResultado.Open varComando, Me.proConexion
    
    If varResultado.EOF Then
        MsgBox "El Asunto Digitado no existe. ", vbCritical, App.Title
        MetValidarCategoriaAsunto = False
        Exit Function
    Else
        If IsNull(varResultado.Fields("iIncidentCategory")) Then
            MsgBox "El Asunto No tiene Categoria Asociada. ", vbCritical, App.Title
            MetValidarCategoriaAsunto = False
            Exit Function
        Else
            If varResultado.Fields("iIncidentCategory") = 3 Then
                If Val(Me.proDatosProductoId) <> 0 Then
                    MsgBox "Este TAB de Datos por Servicio ya tiene ligado un Lead. Debe crear un " & Chr(13) & _
                            "nuevo Item en el TAB de Datos o seleccionar el ticket que lo modifique .", vbInformation, App.Title
                    MetValidarCategoriaAsunto = False
                    Exit Function
                End If
            End If
                    
            If Val(varResultado.Fields("iIncidentCategory")) <> 2 And Val(varResultado.Fields("iIncidentCategory")) <> 3 Then
                MsgBox "El TAB de Datos por Servicio solo puede estar ligado a incidentes de ventas o de atención.", vbInformation, App.Title
                MetValidarCategoriaAsunto = False
                Exit Function
            End If
                    
            If Val(varResultado.Fields("iIncidentCategory")) = 2 Then
                If Val(Trim(Me.proDatosProductoId)) = 0 Then
                    MsgBox "Solo se puede generar un TAB de datos de servicio nuevo por un incidente de ventas.", vbInformation, App.Title
                    MetValidarCategoriaAsunto = False
                    Exit Function
                Else
                    Producto = MetBuscarProductoAsunto()
                    If Producto <> Trim(Me.proProductId) Then
                        MsgBox "El incidente de atencion digitado no pertenece al servicio de esta facturacion.", vbInformation, App.Title
                        MetValidarCategoriaAsunto = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    Set varResultado = Nothing
    MetValidarCategoriaAsunto = True
    Exit Function
ErrManager:
    MetValidarCategoriaAsunto = False
    SubGMuestraError
End Function

Public Function MetBuscarProductoAsunto() As String

'**************************'***********************************************************
'   OBJETIVO:  Busca el producto ligado para un asunto de atencion
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
        
    varComando = "Select    CP.iProductId           " & _
                 "From      Customer_Product CP,    " & _
                 "          Incident I              " & _
                 "Where     I.iIncidentId   =       " & Val(Me.proIncidentId) & " " & _
                 "And       I .vchUser9     = CP.vchSerialNumber "
                            
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields("iProductId").Value) Then
            MetBuscarProductoAsunto = ""
        Else
            MetBuscarProductoAsunto = Trim(varResultado.Fields("iProductId").Value)
        End If
    End If
    
    Set varResultado = Nothing
    
    Exit Function
ErrManager:
    SubGMuestraError
End Function


Public Function MetValidarOTCerrada(Optional varMostrarMensajes As Boolean = True) As Boolean
                                                            
'**************************'***********************************************************
'   OBJETIVO:  Valida si el asunto digitado tiene la ot de instalacion cerrada o no
'**************************'***********************************************************
'   AUTOR: Gustavo Gavilan
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    'Verificar si el incidente ya esta ligado a alguna facturacion
    varComando = "SELECT    iStatusId " & _
                 "FROM      Incident " & _
                 "WHERE     vchUser7 = " & Val(Trim(Me.proIncidentId)) & _
                 " AND      iIncidentTypeId <> 101480" & _
                 " AND      iIncidentCategory = 1 " & _
                 " AND       tiRecordStatus = 1"
            
    varResultado.Open varComando, Me.proConexion
    
    If varResultado.EOF Then
        MetValidarOTCerrada = True
    Else
        If IsNull(varResultado.Fields(0)) Or varResultado.Fields(0) = 0 Then
             If varMostrarMensajes Then MsgBox "La Ot no tiene estado definido.", vbCritical, App.Title
             MetValidarOTCerrada = False
        Else
            Select Case varResultado.Fields(0)
                Case 101456
                    If varMostrarMensajes Then MsgBox "La OT del asunto se encuentra suspendida.", vbInformation, App.Title
                    MetValidarOTCerrada = False
                Case 101465
                    If varMostrarMensajes Then MsgBox "La OT del asunto se encuentra cerrada.", vbInformation, App.Title
                    MetValidarOTCerrada = False
                Case 103679
                    If varMostrarMensajes Then MsgBox "La OT del asunto se encuentra cancelada.", vbInformation, App.Title
                    MetValidarOTCerrada = False
                Case Else
                    MetValidarOTCerrada = True
            End Select
        End If
    End If
    
    Set varResultado = Nothing
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetValidarProducto() As String
'**************************'***********************************************************
'   OBJETIVO:  Valida si el producto pertence al cliente
'**************************'***********************************************************
'   AUTOR: Raúl Cruz
'   FECHA: 29/12/2000
'**************************'***********************************************************

    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    
    On Error GoTo ErrManager
    
    Set varResultado = New ADODB.Recordset
    
    varComando = "SELECT vchSerialNumber, " & _
                 "       iOwnerId         " & _
                 "FROM   Customer_Product " & _
                 "WHERE  iProductId = " & Me.proProductId
            
    varResultado.Open varComando, Me.proConexion
    
    If varResultado.EOF Then
        MsgBox "El producto Digitado no existe. No se almacenara la información del Asunto.", vbCritical, App.Title
        MetValidarProducto = ""
    Else
        If IsNull(varResultado.Fields("iOwnerId")) Then
            MsgBox "El Producto No tiene Cliente Asociado. No se almacenara la información del Asunto.", vbInformation, App.Title
            MetValidarProducto = ""
        Else
            If varResultado.Fields("iOwnerId") <> CLng(Trim(Me.proCompanyId)) Then
                MsgBox "El producto no pertece a este Cliente. No se almacenara la información del Asunto.", vbInformation, App.Title
                MetValidarProducto = ""
            Else
                If IsNull(varResultado.Fields("vchSerialNumber")) Then
                    MsgBox "El producto no tiene codigo de enlace asociado. No se almacenara la información del Asunto.", vbInformation, App.Title
                    MetValidarProducto = ""
                Else
                    MetValidarProducto = varResultado.Fields("vchSerialNumber")
                End If
            End If
        End If
    End If
    Exit Function
ErrManager:
    SubGMuestraError
End Function


