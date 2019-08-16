VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAvanceNumeros 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1170
   ClientLeft      =   4680
   ClientTop       =   5595
   ClientWidth     =   6105
   Icon            =   "frmAvanceNumeros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1170
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlPorcentaje 
      Height          =   1155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _Version        =   65536
      _ExtentX        =   10716
      _ExtentY        =   2037
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
      BevelInner      =   1
      Begin Threed.SSPanel pnlEstado 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   750
         Width           =   4155
         _Version        =   65536
         _ExtentX        =   7329
         _ExtentY        =   556
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlEstadoProceso 
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   420
         Width           =   4155
         _Version        =   65536
         _ExtentX        =   7329
         _ExtentY        =   503
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
         FloodColor      =   16761024
      End
      Begin MSForms.CommandButton cmdRefrescar 
         Height          =   375
         Left            =   4350
         TabIndex        =   5
         Top             =   90
         Width           =   1635
         Caption         =   "Refrescar"
         Size            =   "2884;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   690
         Width           =   1665
         Caption         =   "Salir"
         Size            =   "2937;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblTituloEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado del Proceso:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   150
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmAvanceNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const EstadoInicializando = "Inicializando..."
Const EstadoInsertandoTemporal = "Insertando en temporal..."
Const EstadoInsertandoLog = "Insertando en LOG..."
Const EstadoClasificando = "Clasificando..."
Const EstadoActivandoNumeros = "Activando Numeros..."
Const EstadoInsertandoNumeros = "Insertando Numeros..."
Const EstadoProcesoFinalizado = "Proceso Finalizado"

Public proConexion As ADODB.Connection

Private varParametroTelefonia As claParametrosTelefonia

Private Sub cmdRefrescar_Click()
    Dim varCantidad As Double
    Dim varContador As Double
    Dim varPorcentaje As Double
    On Error GoTo ErrManager
    
    'Buscar el estado en el que se encuentra el proceso
    Set varParametroTelefonia = Nothing
    Set varParametroTelefonia = New claParametrosTelefonia
    Set varParametroTelefonia.proConexion = Me.proConexion
    
    varParametroTelefonia.proParametro = "Estado Insercion Numeros"
    
    If varParametroTelefonia.MetConsultarParametro Then
        Me.pnlEstadoProceso.FloodShowPct = True
        Me.pnlEstadoProceso.FloodType = 1
        Select Case UCase(varParametroTelefonia.proValor)
            Case UCase(EstadoInicializando)
                Me.pnlEstadoProceso.FloodPercent = 1
                Me.pnlEstado.Caption = EstadoInicializando
            Case UCase(EstadoInsertandoTemporal)
                
                'Buscar la cantidad
                varParametroTelefonia.proParametro = "Cantidad"
                
                If varParametroTelefonia.MetConsultarParametro Then
                    varCantidad = Val(varParametroTelefonia.proValor)
                Else
                    MsgBox "Error al consultar la cantidad de registros a insertar.", vbCritical, App.Title
                    Exit Sub
                End If
                
                'Buscar la cantidad de registros procesados
                varParametroTelefonia.proParametro = "Contador"
                
                If varParametroTelefonia.MetConsultarParametro Then
                    varContador = Val(varParametroTelefonia.proValor)
                Else
                    MsgBox "Error al consultar la cantidad de registros a insertar.", vbCritical, App.Title
                    Exit Sub
                End If
                
                varPorcentaje = (varContador * 100) / varCantidad
                varPorcentaje = (varPorcentaje * 30) / 100
                varPorcentaje = varPorcentaje + 10
                
                Me.pnlEstadoProceso.FloodPercent = varPorcentaje
                Me.pnlEstado.Caption = EstadoInsertandoTemporal
                
            Case UCase(EstadoInsertandoLog)
                Me.pnlEstadoProceso.FloodPercent = 40
                Me.pnlEstado.Caption = EstadoInsertandoLog
            Case UCase(EstadoActivandoNumeros)
                Me.pnlEstadoProceso.FloodPercent = 45
                Me.pnlEstado.Caption = EstadoActivandoNumeros
            Case UCase(EstadoInsertandoNumeros)
                Me.pnlEstadoProceso.FloodPercent = 50
                Me.pnlEstado.Caption = EstadoInsertandoNumeros
            Case UCase(EstadoClasificando)
                Me.pnlEstadoProceso.FloodPercent = 60
                Me.pnlEstado.Caption = EstadoClasificando
            Case UCase(EstadoProcesoFinalizado)
                Me.pnlEstadoProceso.FloodPercent = 100
                Me.pnlEstado.Caption = EstadoProcesoFinalizado
        End Select
    Else
        MsgBox "Error al consultar el estado de ejecución.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo ErrManager
    
    Unload Me
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    Dim varCantidad As Double
    Dim varContador As Double
    Dim varPorcentaje As Double
    On Error GoTo ErrManager
    
    'Buscar el estado en el que se encuentra el proceso
    Set varParametroTelefonia = Nothing
    Set varParametroTelefonia = New claParametrosTelefonia
    Set varParametroTelefonia.proConexion = Me.proConexion
    
    varParametroTelefonia.proParametro = "Estado Insercion Numeros"
    
    If varParametroTelefonia.MetConsultarParametro Then
        Me.pnlEstadoProceso.FloodShowPct = True
        Me.pnlEstadoProceso.FloodType = 1
        Select Case UCase(varParametroTelefonia.proValor)
            Case UCase(EstadoInicializando)
                Me.pnlEstadoProceso.FloodPercent = 1
                Me.pnlEstado.Caption = EstadoInicializando
            Case UCase(EstadoInsertandoTemporal)
                
                'Buscar la cantidad
                varParametroTelefonia.proParametro = "Cantidad"
                
                If varParametroTelefonia.MetConsultarParametro Then
                    varCantidad = Val(varParametroTelefonia.proValor)
                Else
                    MsgBox "Error al consultar la cantidad de registros a insertar.", vbCritical, App.Title
                    Exit Sub
                End If
                
                'Buscar la cantidad de registros procesados
                varParametroTelefonia.proParametro = "Contador"
                
                If varParametroTelefonia.MetConsultarParametro Then
                    varContador = Val(varParametroTelefonia.proValor)
                Else
                    MsgBox "Error al consultar la cantidad de registros a insertar.", vbCritical, App.Title
                    Exit Sub
                End If
                
                varPorcentaje = (varContador * 100) / varCantidad
                varPorcentaje = (varPorcentaje * 30) / 100
                varPorcentaje = varPorcentaje + 10
                
                Me.pnlEstadoProceso.FloodPercent = varPorcentaje
                Me.pnlEstado.Caption = EstadoInsertandoTemporal
                
            Case UCase(EstadoInsertandoLog)
                Me.pnlEstadoProceso.FloodPercent = 40
                Me.pnlEstado.Caption = EstadoInsertandoLog
            Case UCase(EstadoActivandoNumeros)
                Me.pnlEstadoProceso.FloodPercent = 45
                Me.pnlEstado.Caption = EstadoActivandoNumeros
            Case UCase(EstadoInsertandoNumeros)
                Me.pnlEstadoProceso.FloodPercent = 50
                Me.pnlEstado.Caption = EstadoInsertandoNumeros
            Case UCase(EstadoClasificando)
                Me.pnlEstadoProceso.FloodPercent = 60
                Me.pnlEstado.Caption = EstadoClasificando
            Case UCase(EstadoProcesoFinalizado)
                Me.pnlEstadoProceso.FloodPercent = 100
                Me.pnlEstado.Caption = EstadoProcesoFinalizado
        End Select
    Else
        MsgBox "Error al consultar el estado de ejecución.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrManager
    
     'Buscar el estado en el que se encuentra el proceso
    Set varParametroTelefonia = Nothing
    Set varParametroTelefonia = New claParametrosTelefonia
    Set varParametroTelefonia.proConexion = Me.proConexion
    
    varParametroTelefonia.proParametro = "Estado Insercion Numeros"
    
    If varParametroTelefonia.MetConsultarParametro Then
        If varParametroTelefonia.proValor = EstadoProcesoFinalizado Then
            frmAdminNumeros.cmdGenerarNumeros.Enabled = True
            frmAdminNumeros.cmdReclasificarNumeros.Enabled = True
            frmAdminNumeros.cmdPorcentajeAvance.Enabled = False
        End If
    Else
        MsgBox "Error al consultar el estado de ejecución.", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
