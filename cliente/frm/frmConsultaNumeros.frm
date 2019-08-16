VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmConsultaNumeros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas de n�meros"
   ClientHeight    =   9540
   ClientLeft      =   3045
   ClientTop       =   990
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   9330
   Begin Threed.SSPanel SSPanel2 
      Height          =   675
      Left            =   0
      TabIndex        =   32
      Top             =   9150
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
         TabIndex        =   36
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSeleccionarTodosNumeros 
         Caption         =   "&Seleccionar Todos"
         Height          =   285
         Left            =   60
         TabIndex        =   34
         Top             =   60
         Width           =   1935
      End
      Begin VB.CommandButton cmdDeseleccionarTodosNumeros 
         Caption         =   "&Deseleccionar Todos"
         Height          =   285
         Left            =   7320
         TabIndex        =   33
         Top             =   60
         Width           =   1935
      End
      Begin VB.Label lblCantidadRegistrosSeleccion 
         Caption         =   "Cantidad Registros Seleccionados"
         Height          =   195
         Left            =   2640
         TabIndex        =   35
         Top             =   90
         Width           =   2445
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   6345
      Left            =   0
      TabIndex        =   20
      Top             =   2790
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   16431
      _ExtentY        =   11192
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
         Height          =   6075
         Left            =   0
         TabIndex        =   13
         Top             =   270
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   10716
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   9285
         _Version        =   65536
         _ExtentX        =   16378
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Resultado de la Consulta"
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   2505
      Left            =   0
      TabIndex        =   15
      Top             =   270
      Width           =   9315
      _Version        =   65536
      _ExtentX        =   16431
      _ExtentY        =   4419
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
         TabIndex        =   37
         ToolTipText     =   "Buscar  n�mero inicial"
         Top             =   810
         Width           =   315
      End
      Begin VB.ComboBox cboCodigoCiudad 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   27
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
         TabIndex        =   24
         Top             =   -90
         Width           =   5565
         Begin VB.CheckBox chkClasificaciones 
            Caption         =   "Usar clasificaciones en conjunto"
            Height          =   405
            Left            =   3630
            TabIndex        =   8
            ToolTipText     =   "Si se marca esta opci�n los n�meros deber�n tener todas las clasificaciones seleccionadas."
            Top             =   480
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.CommandButton cmdDeseleccionarTodos 
            Caption         =   "&Deseleccionar Todos"
            Height          =   285
            Left            =   3510
            TabIndex        =   10
            Top             =   1380
            Width           =   1935
         End
         Begin VB.CommandButton cmdSeleccionarTodos 
            Caption         =   "&Seleccionar Todos"
            Height          =   285
            Left            =   3510
            TabIndex        =   9
            Top             =   1080
            Width           =   1935
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   255
            Left            =   30
            TabIndex        =   25
            Top             =   90
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Clasificaci�n"
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
            TabIndex        =   7
            Top             =   360
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   1879
            _Version        =   393216
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            GridLines       =   0
         End
         Begin VB.Label lblMensaje 
            BackColor       =   &H00C09258&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Si no selecciona ninguna clasificaci�n, se mostrar�n todos los n�meros que no se encuentren clasificados."
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   30
            TabIndex        =   31
            Top             =   1680
            Width           =   5490
         End
         Begin VB.Label lblColorRegistrosSeleccionados 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   195
            Left            =   480
            TabIndex        =   30
            Top             =   1440
            Width           =   195
         End
         Begin VB.Label lblRegistrosSeleccionados 
            Caption         =   "Registros Seleccionados"
            Height          =   195
            Left            =   930
            TabIndex        =   29
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Label lblRegistroSinSeleccionar 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   195
            Left            =   480
            TabIndex        =   28
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdLimpiarControles 
         Caption         =   "&LimpiarControles"
         Height          =   315
         Left            =   7560
         TabIndex        =   12
         Top             =   2130
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   2130
         Width           =   1695
      End
      Begin VB.ComboBox cboCodigoEstado 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   450
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
         Height          =   1005
         Left            =   60
         TabIndex        =   17
         Top             =   1080
         Width           =   3705
         Begin VB.TextBox txtCantidad 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1650
            TabIndex        =   6
            Top             =   570
            Width           =   1965
         End
         Begin VB.TextBox txtNumeroFinal 
            Height          =   315
            Left            =   1650
            TabIndex        =   4
            Top             =   180
            Width           =   1965
         End
         Begin VB.OptionButton optCantidad 
            Caption         =   "       Cantidad:     Max 32000"
            Height          =   435
            Left            =   270
            TabIndex        =   5
            Top             =   480
            Width           =   1305
         End
         Begin VB.OptionButton optNumeroFinal 
            Caption         =   "N�mero Final:"
            Height          =   195
            Left            =   270
            TabIndex        =   3
            Top             =   210
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblCantidad 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad:"
            Height          =   195
            Left            =   840
            TabIndex        =   19
            Top             =   480
            Width           =   675
         End
         Begin VB.Label lblNumeroInicial 
            AutoSize        =   -1  'True
            Caption         =   "N�mero Final:"
            Height          =   195
            Left            =   570
            TabIndex        =   18
            Top             =   210
            Width           =   975
         End
      End
      Begin VB.Label lblCiudad 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   1050
         TabIndex        =   26
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   1050
         TabIndex        =   22
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblNumeroFinal 
         AutoSize        =   -1  'True
         Caption         =   "N�mero Inicial:"
         Height          =   195
         Left            =   570
         TabIndex        =   16
         Top             =   900
         Width           =   1050
      End
   End
   Begin Threed.SSPanel pnlTitulo 
      Height          =   255
      Left            =   0
      TabIndex        =   14
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
Option Explicit

Public proConexion As ADODB.Connection

Private varEstadoNumero As colEstadoNumero
Private varClasificacion As colClasificacion
Private varCiudad As colCiudad

'Variables para manejo de rangos de celda
Private varFShift As Integer
Private varFPosicion As Integer
Private varFPosicionFinal As Integer


Public proLlamadoAdministracion As Boolean

Public proNumeros As colNumero

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

Private Sub cmdBuscar_Click()
    On Error GoTo ErrManager
    
    'Validar los par�metros seleccionados
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
            MsgBox "Debe seleccionar el n�mero final o la cantidad de registros a encontrar partiendo del n�mero inicial.", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    
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
        MsgBox "Error al consultar los n�meros.", vbCritical, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
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
    On Error GoTo ErrManager
    
        'Inicializar Controles
        Me.optNumeroFinal.Value = True
        Me.optCantidad.Value = False
        
        Me.txtCantidad.Enabled = False
        Me.txtNumeroFinal.Enabled = True
        Me.txtCantidad.BackColor = &HE0E0E0
        Me.txtCantidadSeleccionados.BackColor = Me.lblColorRegistrosSeleccionados.BackColor
        
        Call SubFInicializarGridClasificacion
        
        'Inicializar combo de Ciudades
        
        Set varCiudad = New colCiudad
        Set varCiudad.proConexion = Me.proConexion
        
        If varCiudad.MetConsultar Then
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
            MsgBox "Error al consultar los estados de los n�meros.", vbCritical, App.Title
            Exit Sub
        End If
        
        'Llenar el grid de clasificacion
        Set varClasificacion = New colClasificacion
        Set varClasificacion.proConexion = Me.proConexion
        
        If varClasificacion.FunGConsulta Then
            Call SubFLlenarGridClasificacion
        Else
            MsgBox "Error al consultar la informaci�n de clasificaci�n.", vbCritical, App.Title
            Exit Sub
        End If
            
        Call SubFInicializarGridNumeros
        
        If Me.proLlamadoAdministracion Then
            Me.cmdDeseleccionarTodosNumeros.Visible = False
            Me.cmdSeleccionarTodosNumeros.Visible = False
            Me.lblCantidadRegistrosSeleccion.Caption = "Cantidad Registros"
        Else
            Me.cmdDeseleccionarTodosNumeros.Visible = True
            Me.cmdSeleccionarTodosNumeros.Visible = True
            Me.lblCantidadRegistrosSeleccion.Caption = "Cantidad Registros Seleccionados"
        End If
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
    
    If Me.proLlamadoAdministracion Then
        Exit Sub
    End If
    
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
    
    'Si la tecla es shift, selecciona �nicamente el rango indicado.
    If varFShift = 1 Then
        'Debe borrar lo dem�s
        subFLimpiarSeleccion
        varBandera = 0
    'Si la tecla es ctrl agrega a la selecci�n anterior
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

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
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
            .Rows = 0
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

Private Sub SubFLlenarComboCiudades()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cboCodigoCiudad.Clear
    Me.cboNombreCiudad.Clear
    
    Me.cboCodigoCiudad.AddItem "0"
    Me.cboNombreCiudad.AddItem "<<< TODAS >>>"
    For varContador = 1 To varCiudad.Count
        Me.cboCodigoCiudad.AddItem varCiudad.Item(varContador).proCodigoCiudad
        Me.cboNombreCiudad.AddItem varCiudad.Item(varContador).proNombreCiudad
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
    
    Me.grdClasificacion.Rows = 0
        
    For varContador = 1 To varClasificacion.Count
        Me.grdClasificacion.AddItem varClasificacion.Item(varContador).proClasificacionId & vbTab & _
                                    varClasificacion.Item(varContador).proClasificacion & vbTab & _
                                    varClasificacion.Item(varContador).proRecordStatus
        
        If varClasificacion.Item(varContador).proRecordStatus = 0 Then
            Me.grdClasificacion.RowHeight(Me.grdClasificacion.Rows - 1) = 0
        End If
        
        If varClasificacion.Item(varContador).proSeleccionado = "S" Then
            Me.grdClasificacion.Col = 1
            Me.grdClasificacion.Row = Me.grdClasificacion.Rows - 1
            Me.grdClasificacion.CellBackColor = Me.lblColorRegistrosSeleccionados.BackColor
        End If
    Next varContador
    If Me.grdClasificacion.Rows <> 0 Then
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
        .Rows = 1
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
    
    Me.grdNumeros.Rows = 1
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
            Me.grdNumeros.Row = Me.grdNumeros.Rows - 1
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