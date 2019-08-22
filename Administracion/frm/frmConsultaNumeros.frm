VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmConsultaNumeros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas de números"
   ClientHeight    =   9615
   ClientLeft      =   3045
   ClientTop       =   990
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdCambiarEstado 
      Caption         =   "Ca&mbiar Estado"
      Height          =   315
      Left            =   7320
      TabIndex        =   43
      Top             =   720
      Width           =   1500
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   675
      Left            =   0
      TabIndex        =   40
      Top             =   9120
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   16431
      _ExtentY        =   1191
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtCantidadSeleccionados 
         Height          =   285
         Left            =   5220
         TabIndex        =   19
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSeleccionarTodosNumeros 
         Caption         =   "&Seleccionar Todos"
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Top             =   60
         Width           =   1935
      End
      Begin VB.CommandButton cmdDeseleccionarTodosNumeros 
         Caption         =   "&Deseleccionar Todos"
         Height          =   285
         Left            =   7320
         TabIndex        =   20
         Top             =   60
         Width           =   1935
      End
      Begin VB.Label lblCantidadRegistrosSeleccion 
         Caption         =   "Cantidad Registros Seleccionados"
         Height          =   195
         Left            =   2640
         TabIndex        =   41
         Top             =   90
         Width           =   2445
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   5625
      Left            =   0
      TabIndex        =   28
      Top             =   3510
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   16431
      _ExtentY        =   9922
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      Begin MSFlexGridLib.MSFlexGrid grdNumeros 
         Height          =   5355
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   9446
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   315
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   9285
         _Version        =   65536
         _ExtentX        =   16378
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Resultado de la Consulta"
         ForeColor       =   16777215
         BackColor       =   12620376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   6
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3225
      Left            =   0
      TabIndex        =   23
      Top             =   360
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   16431
      _ExtentY        =   5689
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      Begin VB.ComboBox cboCodigoTipoLinea 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.ComboBox cboTipoLinea 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2880
         Width           =   2010
      End
      Begin VB.CommandButton cmdClasificar 
         Caption         =   "&Clasificar seleccionados"
         Height          =   315
         Left            =   5760
         TabIndex        =   15
         Top             =   2880
         Width           =   1950
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "&Asignar"
         Height          =   315
         Left            =   7800
         TabIndex        =   16
         Top             =   2880
         Width           =   1500
      End
      Begin VB.CommandButton cmdBuscarNumeroInicial 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3390
         TabIndex        =   3
         ToolTipText     =   "Buscar  número inicial"
         Top             =   810
         Width           =   315
      End
      Begin VB.ComboBox cboCodigoCiudad 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   90
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox cboNombreCiudad 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   1995
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   3750
         TabIndex        =   32
         Top             =   -90
         Width           =   5565
         Begin VB.CheckBox chkClasificaciones 
            Caption         =   "Usar clasificaciones en conjunto"
            Height          =   405
            Left            =   3630
            TabIndex        =   9
            ToolTipText     =   "Si se marca esta opción los números deberán tener todas las clasificaciones seleccionadas."
            Top             =   480
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.CommandButton cmdDeseleccionarTodos 
            Caption         =   "&Deseleccionar Todos"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3510
            TabIndex        =   11
            Top             =   1380
            Width           =   1935
         End
         Begin VB.CommandButton cmdSeleccionarTodos 
            Caption         =   "&Seleccionar Todos"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3510
            TabIndex        =   10
            Top             =   1080
            Width           =   1935
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   255
            Left            =   30
            TabIndex        =   33
            Top             =   90
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Clasificación"
            ForeColor       =   16777215
            BackColor       =   12620376
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grdClasificacion 
            Height          =   1065
            Left            =   30
            TabIndex        =   8
            Top             =   360
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   1879
            _Version        =   393216
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            GridLines       =   0
            SelectionMode   =   1
         End
         Begin VB.Label lblMensaje 
            BackColor       =   &H00C09258&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Si no selecciona ninguna clasificación, se mostrarán todos los números que no se encuentren clasificados."
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   30
            TabIndex        =   39
            Top             =   1680
            Width           =   5490
         End
         Begin VB.Label lblColorRegistrosSeleccionados 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   1440
            Width           =   195
         End
         Begin VB.Label lblRegistrosSeleccionados 
            Caption         =   "Registros Seleccionados"
            Height          =   195
            Left            =   930
            TabIndex        =   37
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Label lblRegistroSinSeleccionar 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   195
            Left            =   480
            TabIndex        =   36
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdLimpiarControles 
         Caption         =   "&LimpiarControles"
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   2880
         Width           =   1500
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   0
         TabIndex        =   12
         Top             =   2880
         Width           =   1500
      End
      Begin VB.ComboBox cboCodigoEstado 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox cboNombreEstado 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   1995
      End
      Begin VB.TextBox txtNumeroInicial 
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Top             =   810
         Width           =   1635
      End
      Begin VB.Frame fraTipo 
         Height          =   1725
         Left            =   0
         TabIndex        =   25
         Top             =   1080
         Width           =   3735
         Begin VB.CheckBox ChkConsecutivo 
            Caption         =   "Numeración Consecutiva"
            Height          =   315
            Left            =   1440
            TabIndex        =   46
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.TextBox TxtContiene 
            Height          =   315
            Left            =   1650
            TabIndex        =   45
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtCantidad 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1680
            TabIndex        =   7
            Text            =   "1"
            Top             =   600
            Width           =   1965
         End
         Begin VB.TextBox txtNumeroFinal 
            Height          =   315
            Left            =   1650
            TabIndex        =   5
            Top             =   180
            Width           =   1965
         End
         Begin VB.OptionButton optCantidad 
            Caption         =   "       Cantidad:     Max 32000"
            Height          =   435
            Left            =   270
            TabIndex        =   6
            Top             =   480
            Width           =   1305
         End
         Begin VB.OptionButton optNumeroFinal 
            Caption         =   "Número Final:"
            Height          =   195
            Left            =   270
            TabIndex        =   4
            Top             =   210
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblContiene 
            Caption         =   "Contiene:"
            Height          =   255
            Left            =   720
            TabIndex        =   44
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblCantidad 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad:"
            Height          =   195
            Left            =   840
            TabIndex        =   27
            Top             =   480
            Width           =   675
         End
         Begin VB.Label lblNumeroInicial 
            AutoSize        =   -1  'True
            Caption         =   "Número Final:"
            Height          =   195
            Left            =   570
            TabIndex        =   26
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.Label lblLinea 
         AutoSize        =   -1  'True
         Caption         =   "Línea"
         Height          =   195
         Left            =   3120
         TabIndex        =   42
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label lblCiudad 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   1050
         TabIndex        =   34
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   1050
         TabIndex        =   30
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblNumeroFinal 
         AutoSize        =   -1  'True
         Caption         =   "Número Inicial:"
         Height          =   195
         Left            =   570
         TabIndex        =   24
         Top             =   900
         Width           =   1050
      End
   End
   Begin Threed.SSPanel pnlTitulo 
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   16431
      _ExtentY        =   450
      _StockProps     =   15
      Caption         =   "Filtros de la consulta"
      ForeColor       =   16777215
      BackColor       =   12620376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Se agrega en el evento QueryUnload del form
'       la invocacion al metodo MetVerificarAsignacionOnyx de colNumero
'       FECHA:       2009/09/02
'       VERSION:     1.0.000
'       REQUERIMIENTO: 5290
'*******************************************************************
'*******************************************************************
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Se agrega el boton Cambiar Estado y
'       cambio en el load de la forma, cuando el llamado es de administracion
'       para que muestre el nuevo command
'       REQUERIMIENTO:          5322
'       VERSION:       1.0.100
'       FECHA:       2009/10/10
'*******************************************************************
Option Explicit

Public proConexion As ADODB.Connection

Private varEstadoNumero As colEstadoNumero
Private varClasificacion As colClasificacion
Private varciudad As colCiudad

'Variables para manejo de rangos de celda
Private varFShift As Integer
Private varFPosicion As Integer
Private varFPosicionFinal As Integer

Public proLlamadoAdministracion As Boolean
Public proNumeros As colNumero
Public proRegion As String
Public proValoresCampoProducto As colValoresCampoProducto 'Tipos de línea
Public proTipoLineaEdicion As colTipoLineaEdicion 'Tipos de línea en edición
Public proNo As String 'Valor parametrizado para NO
Public proTipoLineaBasico As Boolean 'Indica si el tipo de línea seleccionado es básico o no
Public proCodigoTipoLineaBasica As String 'Codigo del tipo de línea, si es básico
Public proIndiceTipoLineaEdicion As Long  'índice del tipo de línea en proTipoLineaEdicion, si NO es básico
Public proCodCiudad As String 'Código de la ciudad de instalación
Public proIndiceInstalado As Integer
Public proIndiceSeleccionado As Integer
Public proLogin As String

Private Sub cboNombreCiudad_Click()
    On Error GoTo ErrManager
    
        Me.cboCodigoCiudad.ListIndex = Me.cboNombreCiudad.ListIndex
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cboNombreEstado_Click()
    On Error GoTo ErrManager
    
        Me.cboCodigoEstado.ListIndex = Me.cboNombreEstado.ListIndex
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdAsignar_Click()
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo ErrManager
    
    Dim Script As String
    Dim ScriptCity As String
    Dim ScriptNombre As String
    Dim ScriptNumeros As String
    Dim cityCode As String
    Dim areaCode As String
    Dim strEstado As String
    Dim strproClasificacionId As String
    Dim proClasificacionDescripcion As String
    Dim strClasificacionNet As String
    Dim varResultados As ADODB.Recordset
    Dim varResultadosCity As ADODB.Recordset
    Dim varResultadosEstado As ADODB.Recordset
    Dim varResultadosNumeros As ADODB.Recordset
    Dim varResultadoNumeroNet As ADODB.Recordset
    
    'Validar los parámetros seleccionados
    If Me.cboCodigoCiudad.Text = "" Then
        MsgBox "Debe seleccionar la ciudad a buscar.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Trim(Me.cboCodigoEstado.Text) = "" Then
        MsgBox "Debe seleccionar el estado a buscar.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Trim(Me.txtNumeroInicial.Text) <> "" Then
        If Trim(Me.txtNumeroFinal.Text) = "" And Trim(Me.txtCantidad.Text) = "" Then
            MsgBox "Debe seleccionar el número final o la cantidad de registros a encontrar partiendo del número inicial.", vbInformation, App.Title
            Exit Sub
        End If
    End If

    
    'Consulta para la validacion de que flujo seguir
    Set varResultados = New ADODB.Recordset
    Script = "SELECT vchMetododAtributo " & _
                 "FROM AtributosSoapWebService " & _
                 "WHERE vchMetodo = 'NetCracker'"
    varResultados.Open Script, Me.proConexion
    
    'consulta para traer el codigo de la cuidad
    Set varResultadosCity = New ADODB.Recordset
    ScriptCity = "SELECT " & _
                 "Ind.vchCodRegion, " & _
                 "Ind.vchIndicativo, " & _
                 "SUBSTRING (Ciu.vchCodigoCiudad , 1, 5) As vchCodigoCiudad " & _
                 "FROM " & _
                 "CT_IndicativoCiudadROC Ind " & _
                 "INNER JOIN ct_CiudadDANE Ciu " & _
                 "ON Ind.vchCityName = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(Ciu.vchCiudad, 'Á', 'A'), 'É','E'), 'Í', 'I'), 'Ó', 'O'), 'Ú','U') " & _
                 "WHERE Ind.vchCodRegion = " & Chr(39) & cboCodigoCiudad.Text & Chr(39) & " "
    varResultadosCity.Open ScriptCity, Me.proConexion
    
    'consulta para conparar la clasificacion de netcraker y onyx
    If (Me.grdClasificacion.Text = "NUMEROS DORADOS") Then
        strproClasificacionId = ""
        proClasificacionDescripcion = ""
        strClasificacionNet = ""
    Else
        Set varResultadosNumeros = New ADODB.Recordset
        ScriptNumeros = "SELECT cla.iClasificacionId, " & _
                    "cla.vchClasificacionNet, " & _
                    "ct.vchClasificacion " & _
                    "FROM ClasificacionNetCracker cla " & _
                    "INNER JOIN CT_clasificacion ct " & _
                    "ON cla.iClasificacionId = ct.iClasificacionId " & _
                    "WHERE cla.iClasificacionId = " & Me.grdClasificacion.Text
        varResultadosNumeros.Open ScriptNumeros, Me.proConexion
        strproClasificacionId = varResultadosNumeros("iClasificacionId")
        proClasificacionDescripcion = varResultadosNumeros("vchClasificacion")
        strClasificacionNet = varResultadosNumeros("vchClasificacionNet")
    End If
    
    
    'consulta para conparar el estado de netcraker y onyx
    Set varResultadosEstado = New ADODB.Recordset
    ScriptNombre = "SELECT vchEstadoNetCracker FROM EstadoNetCracker WHERE vchEstadoOnyx = " & Chr(39) & cboCodigoEstado.Text & Chr(39)
    varResultadosEstado.Open ScriptNombre, Me.proConexion
    
    'variables necesarias para el consumo del servicio web
    cityCode = varResultadosCity("vchCodigoCiudad")
    areaCode = varResultadosCity("vchIndicativo")
    strEstado = varResultadosEstado("vchEstadoNetCracker")
    
    While varResultados.EOF = False
        If (varResultados("vchMetododAtributo") = "true") Then
                    
            Dim resultWS As Object
            Dim tipo As String
            Dim classWS As claRequestWs
            Dim objetoPrueba As claRequestWs
            Dim varContador As Integer
            Dim consecutiveCheck As String
            
            Set classWS = New claRequestWs
            
            tipo = "getNumbers"
            Screen.MousePointer = 11
            Set classWS.proConexion = Me.proConexion
            Set objetoPrueba = New claRequestWs
            Set classWS.coleccionPrueba = New Collection
            
            If (ChkConsecutivo.Value = 1) Then
                consecutiveCheck = "true"
            Else
                consecutiveCheck = "false"
            End If
            
            objetoPrueba.crm_in_use = "TCRM"
            classWS.coleccionPrueba.Add objetoPrueba.crm_in_use, "crm_in_use"
            objetoPrueba.request_id = "Example_PMO-001"
            classWS.coleccionPrueba.Add objetoPrueba.request_id, "request_id"
            objetoPrueba.transaction_id = "Example_PMO-001"
            classWS.coleccionPrueba.Add objetoPrueba.transaction_id, "transaction_id"
            objetoPrueba.city_code = cityCode
            classWS.coleccionPrueba.Add objetoPrueba.city_code, "city_code"
            objetoPrueba.country_code = "57"
            classWS.coleccionPrueba.Add objetoPrueba.country_code, "country_code"
            objetoPrueba.area_code = areaCode
            classWS.coleccionPrueba.Add objetoPrueba.area_code, "area_code"
            objetoPrueba.consecutive_number = consecutiveCheck
            classWS.coleccionPrueba.Add objetoPrueba.consecutive_number, "consecutive_number"
            objetoPrueba.quantity_numbers = Me.txtCantidad.Text
            classWS.coleccionPrueba.Add objetoPrueba.quantity_numbers, "quantity_numbers"
            objetoPrueba.number_mask = TxtContiene.Text
            classWS.coleccionPrueba.Add objetoPrueba.number_mask, "number_mask"
            objetoPrueba.initial_number = Me.txtNumeroInicial.Text
            classWS.coleccionPrueba.Add objetoPrueba.initial_number, "initial_number"
            objetoPrueba.final_number = Me.txtNumeroFinal.Text
            classWS.coleccionPrueba.Add objetoPrueba.final_number, "final_number"
            objetoPrueba.category = strClasificacionNet
            classWS.coleccionPrueba.Add objetoPrueba.category, "category"
            objetoPrueba.status = strEstado
            classWS.coleccionPrueba.Add objetoPrueba.status, "status"
            
            classWS.coleccionPrueba.Add objetoPrueba
            Set resultWS = classWS.RequestPeticionWs(tipo)
            
            If (dataWS = "") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            Me.grdNumeros.rows = 1
            Me.grdNumeros.Redraw = False
            
            For varContador = 1 To resultWS.Count
                                
                Dim getNumero As String
                'Dim ScriptNumeroNet As String
                getNumero = resultWS.Item("item" & varContador).Item(1).Item("number")
                
                'Set varResultadoNumeroNet = New ADODB.Recordset
                
                'ScriptNumeroNet = "SELECT chUpdateBy, dtUpdateDate FROM CT_Numeros " & _
                                  "WHERE vchNumero = " & Chr(39) & Mid(getNumero, 4) & Chr(39) & " AND chRegionCode = " & Chr(39) & cboCodigoCiudad.Text & Chr(39)
                'varResultadoNumeroNet.Open ScriptNumeroNet, Me.proConexion
                
                
                Me.grdNumeros.AddItem cboCodigoCiudad.Text & vbTab & _
                                      cboNombreCiudad.Text & vbTab & _
                                      Mid(getNumero, 4) & vbTab & _
                                      cboCodigoEstado.Text & vbTab & _
                                      cboNombreEstado.Text & vbTab & _
                                      strproClasificacionId & vbTab & _
                                      proClasificacionDescripcion & vbTab & _
                                      "" & vbTab & _
                                      ""
                              
            Next varContador
            
            Me.grdNumeros.Row = 0
            Me.grdNumeros.Col = 0
            Me.grdNumeros.Redraw = True
                        
        Else
        
            Screen.MousePointer = 11
            Set Me.proNumeros = Nothing
            Set Me.proNumeros = New colNumero
            Set Me.proNumeros.proConexion = Me.proConexion
            Set Me.proNumeros.proClasificacion = varClasificacion
            
            Me.proNumeros.proCantidadNumeros = Me.txtCantidad.Text
            Me.proNumeros.proEstado = Me.cboCodigoEstado.Text
            Me.proNumeros.proNumeroInicial = Me.txtNumeroInicial.Text
            Me.proNumeros.proNumeroFinal = Me.txtNumeroFinal.Text
            Me.proNumeros.proRegionCode = Me.cboCodigoCiudad.Text
            Me.proNumeros.proUsarConjuntoClasificaciones = Me.chkClasificaciones.Value
                
            If Me.proNumeros.MetConsultarNumeros Then
                Call SubFPintarGridNumeros
            Else
                MsgBox "Error al consultar los números.", vbCritical, App.Title
                Screen.MousePointer = 0
                Exit Sub
            End If
            
        End If
        varResultados.MoveNext
    Wend
   
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdBuscarNumeroInicial_Click()
   On Error GoTo ErrorManager

    If cboCodigoCiudad.ListIndex > -1 And cboCodigoEstado.ListIndex > -1 Then
        Set frmRangosNumeros.proConexion = Me.proConexion
        frmRangosNumeros.proRegionCode = cboCodigoCiudad.List(cboCodigoCiudad.ListIndex)
        frmRangosNumeros.proEstadoNumero = cboCodigoEstado.List(cboCodigoEstado.ListIndex)
        frmRangosNumeros.Show vbModal
        txtNumeroInicial.Text = frmRangosNumeros.proInicio
    Else
        MsgBox "Debe seleccionar una ciudad y un estado"
    End If

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdCambiarEstado_Click()
    On Error GoTo ErrManager
    Dim i As Integer
    
    If Not Me.proNumeros Is Nothing Then
        If Me.proNumeros.proSeleccionados > 0 Then
            Set frmCambiarEstadoNumerosSeleccionados.proConexion = Me.proConexion
            frmCambiarEstadoNumerosSeleccionados.proLogin = Me.proLogin
            
            Set frmCambiarEstadoNumerosSeleccionados.proNumeros = Me.proNumeros
            frmCambiarEstadoNumerosSeleccionados.Show vbModal
            If frmCambiarEstadoNumerosSeleccionados.proGuardado Then
                cmdBuscar_Click
            End If
            'Unload frmCambiarEstadoNumerosSeleccionados
        Else
            MsgBox "Por favor seleccione los números que desea cambiar de estado"
        End If
    Else
        MsgBox "Primero debe realizar una consulta"
    End If
    Exit Sub
    
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdClasificar_Click()
    On Error GoTo ErrManager
    Dim i As Integer
    
    If Not Me.proNumeros Is Nothing Then
        If Me.proNumeros.proSeleccionados > 0 Then
            Set frmClasificarNumerosSeleccionados.proConexion = Me.proConexion
            
            Set frmClasificarNumerosSeleccionados.proNumeros = Me.proNumeros
            frmClasificarNumerosSeleccionados.Show vbModal, Me
            If frmClasificarNumerosSeleccionados.proGuardado Then
                SubFGenerarConsulta (frmClasificarNumerosSeleccionados.cbClasificacion.ListIndex)
            End If
            Unload frmClasificarNumerosSeleccionados
        Else
            MsgBox "Por favor seleccione los números que desea clasificar"
        End If
    Else
        MsgBox "Primero debe realizar una consulta"
    End If
    Exit Sub
    
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdDeseleccionarTodos_Click()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
        For varContador = 1 To varClasificacion.Count
            varClasificacion.Item(varContador).proSeleccionado = "N"
        Next varContador
    
        Call SubFLlenarGridClasificacion
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdDeseleccionarTodosNumeros_Click()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    If Me.proNumeros Is Nothing Then
        Exit Sub
    End If
    
    For varContador = 1 To Me.proNumeros.Count
        Me.proNumeros.Item(varContador).proSeleccionado = "N"
    Next varContador
    
    Me.proNumeros.proSeleccionados = 0
    Me.txtCantidadSeleccionados.Text = 0
    
    Call SubFPintarGridNumeros
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdLimpiarControles_Click()
    On Error GoTo ErrManager
    
        Me.txtNumeroInicial.Text = ""
        Me.txtNumeroFinal.Text = ""
        Me.txtCantidad.Text = ""
        
        Me.cboNombreEstado.ListIndex = -1
        
        Me.cboNombreCiudad.ListIndex = -1
        
        Me.chkClasificaciones.Value = 0
        
        Call cmdDeseleccionarTodos_Click
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdSeleccionarTodos_Click()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
        For varContador = 1 To varClasificacion.Count
            varClasificacion.Item(varContador).proSeleccionado = "S"
        Next varContador
    
        Call SubFLlenarGridClasificacion
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub



Private Sub cmdSeleccionarTodosNumeros_Click()
    Dim varContador As Integer
    On Error GoTo ErrManager
        
    If Me.proNumeros Is Nothing Then
        Exit Sub
    End If
    For varContador = 1 To Me.proNumeros.Count
        Me.proNumeros.Item(varContador).proSeleccionado = "S"
    Next varContador
    
    Me.proNumeros.proSeleccionados = Me.proNumeros.Count
    Me.txtCantidadSeleccionados.Text = Me.proNumeros.Count
    
    Call SubFPintarGridNumeros
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    Dim varIndice As Integer
    Dim i As Integer
    On Error GoTo ErrManager
        If Not Me.proNumeros Is Nothing Then
            For i = 1 To Me.proNumeros.Count
                Me.proNumeros.Remove i
            Next
        End If
        'Inicializar Controles
        Me.optNumeroFinal.Value = True
        Me.optCantidad.Value = False
        
        Me.txtCantidad.Enabled = False
        Me.txtNumeroFinal.Enabled = True
        Me.txtCantidad.BackColor = &HE0E0E0
        Me.txtCantidadSeleccionados.BackColor = Me.lblColorRegistrosSeleccionados.BackColor
        
        Call SubFInicializarGridClasificacion
        
        'Inicializar combo de Ciudades
        
        Set varciudad = New colCiudad
        Set varciudad.proConexion = Me.proConexion
        
        If varciudad.MetConsultar Then
            Call SubFLlenarComboCiudades
        Else
            MsgBox "Error al consultar las ciudades.", vbCritical, App.Title
            Exit Sub
        End If
        
        'Inicializar combo de estados
        Set varEstadoNumero = New colEstadoNumero
        Set varEstadoNumero.proConexion = Me.proConexion
        
        If varEstadoNumero.MetConsulta Then
            Call SubFLlenarComboEstado
        Else
            MsgBox "Error al consultar los estados de los números.", vbCritical, App.Title
            Exit Sub
        End If

        'Inicializar combo de tipos de línea
        If Not proTipoLineaEdicion Is Nothing Then
           SubFLlenarComboTipoLinea
        End If
        
        'Cargar por defecto la ciudad de instalación
        For varIndice = 0 To cboCodigoCiudad.ListCount - 1
            If cboCodigoCiudad.List(varIndice) = proCodCiudad Then
                cboCodigoCiudad.ListIndex = varIndice
                cboNombreCiudad.ListIndex = varIndice
                Exit For
            End If
        Next

        'Llenar el grid de clasificacion
        Set varClasificacion = New colClasificacion
        Set varClasificacion.proConexion = Me.proConexion
        
        If varClasificacion.FunGConsulta Then
            Call SubFLlenarGridClasificacion
        Else
            MsgBox "Error al consultar la información de clasificación.", vbCritical, App.Title
            Exit Sub
        End If
            
        Call SubFInicializarGridNumeros
        
        If Me.proLlamadoAdministracion Then
            Me.cmdDeseleccionarTodosNumeros.Visible = False
            Me.cmdSeleccionarTodosNumeros.Visible = False
            Me.lblCantidadRegistrosSeleccion.Caption = "Cantidad Registros"
            Me.cmdClasificar.Visible = True
             '/* 1.0.100  -  Inicio */
            Me.cmdCambiarEstado.Visible = True
            '/* 1.0.100  -  Fin */
        Else
            Me.cmdDeseleccionarTodosNumeros.Visible = True
            Me.cmdSeleccionarTodosNumeros.Visible = True
            Me.lblCantidadRegistrosSeleccion.Caption = "Cantidad Registros Seleccionados"
            Me.cmdClasificar.Visible = False
             '/* 1.0.100  -  Inicio */
            Me.cmdCambiarEstado.Visible = True
            '/* 1.0.100  -  Fin */
        End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim varContador As Integer, varIndice As Integer, varMaximo As Long, varError As String
    On Error GoTo ErrManager
    
    'Si no hay números seleccionados, salir
    If Me.proNumeros Is Nothing Then Exit Sub
    If Me.proNumeros.proSeleccionados = 0 Then Exit Sub
    proIndiceSeleccionado = cboTipoLinea.ListIndex + 1
    If cboTipoLinea.ListIndex < 0 Then
        cboTipoLinea.SetFocus
        MsgBox "Debe seleccionar un tipo de línea para asignar números", vbCritical, App.Title
        Cancel = 1
        Exit Sub
    End If

    'Asignar variables indicando el tipo de línea seleccionado
    proTipoLineaBasico = (cboTipoLinea.ItemData(cboTipoLinea.ListIndex) = 1)
    If proTipoLineaBasico Then
        proCodigoTipoLineaBasica = cboCodigoTipoLinea.List(cboTipoLinea.ListIndex)
    Else
        proIndiceTipoLineaEdicion = cboCodigoTipoLinea.List(cboTipoLinea.ListIndex)
    End If

    'Validar que el número de líneas asignado corresponda con la cantidad parametrizada
    varMaximo = 0
    If proTipoLineaBasico Then 'Tipo de línea básica
        For varContador = 1 To proTipoLineaEdicion.Count
            If proTipoLineaEdicion.Item(varContador).proUser15 = proNo Then 'No es backup
                If proTipoLineaEdicion.Item(varContador).proUser1 = proCodigoTipoLineaBasica Then 'Para todos los tipos de linea en edición con el tipo de línea seleccionado
                    varMaximo = varMaximo + 1
                    varMaximo = varMaximo - proTipoLineaEdicion.Item(varContador).proContadorNumeros
                End If
            End If
        Next
    Else 'Tipo de línea en edición
        varIndice = proValoresCampoProducto.BuscarIndiceProValorId(proTipoLineaEdicion.Item(proTipoLineaEdicion.IndexOf(proIndiceTipoLineaEdicion)).proUser1)
        If varIndice > -1 Then
            varMaximo = varMaximo + proValoresCampoProducto.Item(varIndice).proMaximo
            varMaximo = varMaximo - proTipoLineaEdicion.Item(proTipoLineaEdicion.IndexOf(proIndiceTipoLineaEdicion)).proContadorNumeros
        End If
    End If

    If proNumeros.proSeleccionados > varMaximo Then
        MsgBox "No se pueden asignar más números que los seleccionados para este tipo de línea (" & varMaximo & ")"
        Cancel = 1
        Call cmdDeseleccionarTodosNumeros_Click
        Exit Sub  '--- 1.0.000 (se agrego el exit sub)
    End If
    
    '/* 1.0.000  -  Inicio */
    'Si selecciono varios nros. en estado libre tengo que validar la no existencia en softswitch de estos nros.
    If proNumeros.proEstado = "L" Then
        If Not proNumeros.MetVerificarAsignacionOnyx(varError) Then
            MsgBox varError, vbInformation, App.Title
            Cancel = 1
            Call cmdDeseleccionarTodosNumeros_Click
            Exit Sub
        End If
    End If
    '/* 1.0.000  -  Fin */
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdClasificacion_DblClick()
    On Error GoTo ErrManager
    
    If Me.grdClasificacion.Row = -1 Then
        Exit Sub
    End If
        
    If varClasificacion.Item(Me.grdClasificacion.Row + 1).proSeleccionado = "N" Then
        varClasificacion.Item(Me.grdClasificacion.Row + 1).proSeleccionado = "S"
        Me.grdClasificacion.CellBackColor = Me.lblColorRegistrosSeleccionados.BackColor
    Else
        varClasificacion.Item(Me.grdClasificacion.Row + 1).proSeleccionado = "N"
        Me.grdClasificacion.CellBackColor = Me.lblRegistroSinSeleccionar.BackColor
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub grdNumeros_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrManager
    
    varFPosicion = Me.grdNumeros.Row
    varFShift = Shift
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub subFLimpiarSeleccion()
Dim varCuenta As Integer
Dim varCuentaColumna As Integer
On Error GoTo ErrorManager

    
    For varCuenta = 1 To Me.proNumeros.Count
        If Me.proNumeros(varCuenta).proSeleccionado = "S" Then
                Me.proNumeros(varCuenta).proSeleccionado = False
                Me.grdNumeros.Row = varCuenta
                For varCuentaColumna = 0 To Me.grdNumeros.Cols - 1
                        Me.grdNumeros.Col = varCuentaColumna
                        Me.grdNumeros.CellBackColor = Me.lblRegistroSinSeleccionar.BackColor
                Next varCuentaColumna
        End If
    Next varCuenta
    
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub grdNumeros_SelChange()
    Dim varPosicion1 As Integer
    Dim varPosicion2 As Integer
    Dim varCuenta As Integer
    Dim varCuentaColumna As Integer
    Dim varBandera As Integer

    On Error GoTo ErrManager
    
    'If Me.proLlamadoAdministracion Then
    '    Exit Sub
    'End If
    
    Me.grdNumeros.Redraw = False
    
    varFPosicion = Me.grdNumeros.RowSel
    varFPosicionFinal = Me.grdNumeros.Row
    
    If varFPosicion = 0 Then varFPosicion = 1
    If varFPosicionFinal = 0 Then varFPosicionFinal = 1
    
    If varFPosicion > varFPosicionFinal Then
        varPosicion1 = varFPosicionFinal
        varPosicion2 = varFPosicion
    Else
        varPosicion1 = varFPosicion
        varPosicion2 = varFPosicionFinal
    End If
    
    'Si la tecla es shift, selecciona únicamente el rango indicado.
    If varFShift = 1 Then
        'Debe borrar lo demás
        subFLimpiarSeleccion
        varBandera = 0
    'Si la tecla es ctrl agrega a la selección anterior
    ElseIf varFShift = 2 Then
        varBandera = 1
    'Si la tecla es shift + ctrl agrega el rango a lo seleccionado
    ElseIf varFShift = 3 Then
        varBandera = 2
    Else
        subFLimpiarSeleccion
        varPosicion2 = varPosicion1
    End If
    
    If varFShift = 2 Or varFShift = 1 Then
        If varPosicion1 <> varPosicion2 Then
            Me.proNumeros.proSeleccionados = 0
            For varCuenta = varPosicion1 To varPosicion2
                Me.proNumeros(varCuenta).proSeleccionado = "S"
            
                Me.grdNumeros.Row = varCuenta
                
                For varCuentaColumna = 0 To Me.grdNumeros.Cols - 1
                        Me.grdNumeros.Col = varCuentaColumna
                        Me.grdNumeros.CellBackColor = Me.lblColorRegistrosSeleccionados.BackColor
                Next varCuentaColumna
            Next varCuenta
        Else
                For varCuenta = varPosicion1 To varPosicion2
                    If Me.proNumeros(varCuenta).proSeleccionado = "S" Then
                        Me.proNumeros(varCuenta).proSeleccionado = "N"
                    Else
                        Me.proNumeros(varCuenta).proSeleccionado = "S"
                    End If
                    
                    Me.grdNumeros.Row = varCuenta
                    
                    If Me.proNumeros(varCuenta).proSeleccionado = "S" Then
                        For varCuentaColumna = 0 To Me.grdNumeros.Cols - 1
                                Me.grdNumeros.Col = varCuentaColumna
                                Me.grdNumeros.CellBackColor = Me.lblColorRegistrosSeleccionados.BackColor
                        Next varCuentaColumna
                    Else
                        For varCuentaColumna = 0 To Me.grdNumeros.Cols - 1
                                Me.grdNumeros.Col = varCuentaColumna
                                Me.grdNumeros.CellBackColor = Me.lblRegistroSinSeleccionar.BackColor
                        Next varCuentaColumna
                    End If
                Next varCuenta
        End If
    ElseIf varFShift = 3 Then 'Shift y control
        For varCuenta = varPosicion1 To varPosicion2
            Me.proNumeros(varCuenta).proSeleccionado = "S"
            
            Me.grdNumeros.Row = varCuenta
            
            For varCuentaColumna = 0 To Me.grdNumeros.Cols - 1
                    Me.grdNumeros.Col = varCuentaColumna
                    Me.grdNumeros.CellBackColor = Me.lblColorRegistrosSeleccionados.BackColor
            Next varCuentaColumna
        Next varCuenta
    End If
    
    Me.proNumeros.proSeleccionados = 0
    For varCuenta = 1 To Me.proNumeros.Count
        If Me.proNumeros.Item(varCuenta).proSeleccionado = "S" Then
            Me.proNumeros.proSeleccionados = Me.proNumeros.proSeleccionados + 1
        End If
    Next varCuenta
    
    Me.txtCantidadSeleccionados.Text = Me.proNumeros.proSeleccionados
    Me.grdNumeros.Redraw = True
    Me.grdNumeros.Row = 0
    
    Exit Sub
ErrManager:
    SubGMuestraError
    Me.grdNumeros.Redraw = True
End Sub

Private Sub optCantidad_Click()
    On Error GoTo ErrManager
    
        Me.txtCantidad.Enabled = True
        Me.txtNumeroFinal.Enabled = False
        Me.txtNumeroFinal.Text = ""
        Me.txtNumeroFinal.BackColor = &HE0E0E0
        Me.txtCantidad.BackColor = &HFFFFFF
        Me.txtCantidad.SetFocus
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub optNumeroFinal_Click()
    On Error GoTo ErrManager
    
        Me.txtNumeroFinal.Enabled = True
        Me.txtCantidad.Enabled = False
        Me.txtCantidad.Text = ""
        Me.txtNumeroFinal.BackColor = &HFFFFFF
        Me.txtCantidad.BackColor = &HE0E0E0
        Me.txtNumeroFinal.SetFocus
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtCantidad_GotFocus()
    On Error GoTo ErrManager
    
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
        KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtCantidad_Validate(Cancel As Boolean)
    On Error GoTo ErrManager
    
    If Trim(Me.txtCantidad.Text) <> "" Then
        If CDbl(Trim(Me.txtCantidad.Text)) > 32000 Or CDbl(Trim(Me.txtCantidad.Text)) <= 0 Then
            MsgBox "El valor debe ser entre 1 y 32000.", vbInformation, App.Title
            Cancel = True
        End If
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtNumeroFinal_GotFocus()
    On Error GoTo ErrManager
    
        Me.txtNumeroFinal.SelStart = 0
        Me.txtNumeroFinal.SelLength = Len(Me.txtNumeroFinal.Text)
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtNumeroFinal_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
        KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtNumeroInicial_GotFocus()
    On Error GoTo ErrManager
    
        Me.txtNumeroInicial.SelStart = 0
        Me.txtNumeroFinal.SelLength = Len(Me.txtNumeroInicial.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtNumeroInicial_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridClasificacion()
    On Error GoTo ErrManager
    
        With Me.grdClasificacion
            .rows = 0
            .Cols = 3
            .ColWidth(0) = 0
            .ColWidth(1) = 2915
            .ColWidth(2) = 0
        End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboEstado()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cboCodigoEstado.Clear
    Me.cboNombreEstado.Clear
    
    Me.cboCodigoEstado.AddItem "0"
    Me.cboNombreEstado.AddItem "<<< TODOS >>>"
    
    For varContador = 1 To varEstadoNumero.Count
        Me.cboCodigoEstado.AddItem varEstadoNumero.Item(varContador).proEstadoNumero
        Me.cboNombreEstado.AddItem varEstadoNumero.Item(varContador).proDescripcionEstado
    Next
    
    Me.cboCodigoEstado.ListIndex = -1
    Me.cboNombreEstado.ListIndex = -1
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboTipoLinea()
    Dim varContador As Integer, varIndice As Integer
    Dim varTipoUsado As Boolean
    On Error GoTo ErrManager

    Me.cboCodigoTipoLinea.Clear
    Me.cboTipoLinea.Clear
    proIndiceInstalado = 0
    'Agregar Básicas
    For varIndice = 1 To proValoresCampoProducto.Count
        If proValoresCampoProducto.Item(varIndice).proMaximo = 1 Then
            'Revisar si existen tipos de línea en edición para el tipo básico
            varTipoUsado = False
            For varContador = 1 To proTipoLineaEdicion.Count
                If proTipoLineaEdicion.Item(varContador).proUser15 = proNo Then 'No es backup
                    If proTipoLineaEdicion.Item(varContador).proUser1 = proValoresCampoProducto.Item(varIndice).proValorId Then
                        varTipoUsado = True
                        Exit For
                    End If
                End If
            Next
            If varTipoUsado Then 'Agregar tipo de línea básico
                Me.cboCodigoTipoLinea.AddItem proValoresCampoProducto.Item(varIndice).proValorId 'Código del tipo de línea
                Me.cboTipoLinea.AddItem proValoresCampoProducto.Item(varIndice).proValorDesc
                Me.cboTipoLinea.ItemData(cboTipoLinea.NewIndex) = 1 'Indica que el tipo de línea seleccionado es básico
            End If
        End If
    Next

    'Agregar tipos de línea en edición
    For varContador = 1 To proTipoLineaEdicion.Count
        If proTipoLineaEdicion.Item(varContador).proUser15 = proNo Then 'No es backup
            varIndice = proValoresCampoProducto.BuscarIndiceProValorId(proTipoLineaEdicion.Item(varContador).proUser1)
            If varIndice > -1 Then
                If proValoresCampoProducto.Item(varIndice).proMaximo <> 1 Then
                    Me.cboCodigoTipoLinea.AddItem proTipoLineaEdicion.Item(varContador).proNovedadDetalleDatosProductoId 'Índice en proTipoLineaEdicion (=proNovedadDetalleDatosProducto)
                    Me.cboTipoLinea.AddItem proValoresCampoProducto.Item(varIndice).proValorDesc & " (" & proTipoLineaEdicion.Item(varContador).proNovedadDetalleDatosProductoId & ")"
                    Me.cboTipoLinea.ItemData(cboTipoLinea.NewIndex) = 0 'Indica que el tipo de línea seleccionado NO es básico
                    If proIndiceInstalado = 0 And proTipoLineaEdicion.Item(varContador).proNovedad = False Then
                        proIndiceInstalado = cboTipoLinea.ListCount
                    End If
                End If
            End If
        End If
    Next
    If cboCodigoTipoLinea.ListCount <> 1 Then
        Me.cboTipoLinea.ListIndex = -1
        Me.cboCodigoTipoLinea.ListIndex = -1
    Else
        Me.cboTipoLinea.ListIndex = 0
        Me.cboCodigoTipoLinea.ListIndex = 0
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboCiudades()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cboCodigoCiudad.Clear
    Me.cboNombreCiudad.Clear
    
    Me.cboCodigoCiudad.AddItem "0"
    Me.cboNombreCiudad.AddItem "<<< TODAS >>>"
    For varContador = 1 To varciudad.Count
        Me.cboCodigoCiudad.AddItem varciudad.Item(varContador).proCodigoCiudad
        Me.cboNombreCiudad.AddItem varciudad.Item(varContador).proNombreCiudad
    Next
    
    Me.cboCodigoCiudad.ListIndex = -1
    Me.cboNombreCiudad.ListIndex = -1
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarGridClasificacion()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdClasificacion.rows = 0
        
    For varContador = 1 To varClasificacion.Count
        Me.grdClasificacion.AddItem varClasificacion.Item(varContador).proClasificacionId & vbTab & _
                                    varClasificacion.Item(varContador).proClasificacion & vbTab & _
                                    varClasificacion.Item(varContador).proRecordStatus
        
        If varClasificacion.Item(varContador).proRecordStatus = 0 Then
            Me.grdClasificacion.RowHeight(Me.grdClasificacion.rows - 1) = 0
        End If
        
        If varClasificacion.Item(varContador).proSeleccionado = "S" Then
            Me.grdClasificacion.Col = 1
            Me.grdClasificacion.Row = Me.grdClasificacion.rows - 1
            Me.grdClasificacion.CellBackColor = Me.lblColorRegistrosSeleccionados.BackColor
        End If
    Next varContador
    If Me.grdClasificacion.rows <> 0 Then
        Me.grdClasificacion.Row = 0
        Me.grdClasificacion.Col = 1
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGridNumeros()
    On Error GoTo ErrManager:
    
    With Me.grdNumeros
        .Cols = 9
        .rows = 1
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 0
        .TextMatrix(0, 0) = "Codigo Ciudad"
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 1425
        .TextMatrix(0, 1) = "Ciudad"
        
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1215
        .TextMatrix(0, 2) = "Numero"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 0
        .TextMatrix(0, 3) = "Codigo Estado"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 960
        .TextMatrix(0, 4) = "Estado"
        
        .Col = 5
        .CellAlignment = 4
        .ColWidth(5) = 0
        .TextMatrix(0, 5) = "Codigo Clasificacion"
        
        .Col = 6
        .CellAlignment = 4
        .ColWidth(6) = 1800
        .TextMatrix(0, 6) = "Clasificacion"
        
        .Col = 7
        .CellAlignment = 4
        .ColWidth(7) = 1455
        .TextMatrix(0, 7) = "Usuario"
        
        .Col = 8
        .CellAlignment = 4
        .ColWidth(8) = 2040
        .TextMatrix(0, 8) = "Fecha"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridNumeros()
    Dim varContador As Integer
    Dim varContadorAux As Integer
    On Error GoTo ErrManager
    
    Me.grdNumeros.rows = 1
    Me.grdNumeros.Redraw = False
    For varContador = 1 To Me.proNumeros.Count
        Me.grdNumeros.AddItem Me.proNumeros.Item(varContador).proRegionCode & vbTab & _
                              Me.proNumeros.Item(varContador).proRegionCodeDescripcion & vbTab & _
                              Me.proNumeros.Item(varContador).proNumero & vbTab & _
                              Me.proNumeros.Item(varContador).proEstadoNumero & vbTab & _
                              Me.proNumeros.Item(varContador).proEstadoNumeroDescripcion & vbTab & _
                              Me.proNumeros.Item(varContador).proClasificacionId & vbTab & _
                              Me.proNumeros.Item(varContador).proClasificacionDescripcion & vbTab & _
                              Me.proNumeros.Item(varContador).proUpdateBy & vbTab & _
                              Me.proNumeros.Item(varContador).proUpdateDate
                              
        If Me.proNumeros.Item(varContador).proSeleccionado = "S" Then
            Me.grdNumeros.Row = Me.grdNumeros.rows - 1
            For varContadorAux = 0 To Me.grdNumeros.Cols - 1
                Me.grdNumeros.Col = varContadorAux
                Me.grdNumeros.CellBackColor = Me.lblColorRegistrosSeleccionados.BackColor
            Next varContadorAux
        End If
                              
    Next varContador
    
    Me.grdNumeros.Row = 0
    Me.grdNumeros.Col = 0
    Me.grdNumeros.Redraw = True
    If Me.proLlamadoAdministracion Then
        Me.txtCantidadSeleccionados.Text = Me.proNumeros.Count
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFGenerarConsulta(indice As Integer)
    Me.grdClasificacion.Row = indice
    grdClasificacion_DblClick
    If frmClasificarNumerosSeleccionados.proModo = 3 And frmClasificarNumerosSeleccionados.cbCodigoClasificacion.Text = "3" Then
        SubConsultaEstado ("L")
    ElseIf frmClasificarNumerosSeleccionados.proModo = 3 And frmClasificarNumerosSeleccionados.cbCodigoClasificacion.Text = "4" Then
        SubConsultaEstado ("L")
    ElseIf frmClasificarNumerosSeleccionados.cbCodigoClasificacion.Text = "3" Then
        SubConsultaEstado ("F")
    ElseIf frmClasificarNumerosSeleccionados.cbCodigoClasificacion.Text = "4" Then
        SubConsultaEstado ("V")
    End If
    Me.txtNumeroInicial.Text = Me.proNumeros.Item(1).proNumero
    Me.optNumeroFinal.Value = True
    Me.txtNumeroFinal.Text = Me.proNumeros.Item(Me.proNumeros.Count).proNumero
    cmdBuscar_Click
End Sub

Private Sub SubConsultaEstado(estado As String)
    Dim i As Integer
    For i = 1 To varEstadoNumero.Count
        If varEstadoNumero.Item(i).proEstadoNumero = estado Then
            Me.cboCodigoEstado.ListIndex = i
            Me.cboNombreEstado.ListIndex = i
            Exit For
        End If
    Next
End Sub



