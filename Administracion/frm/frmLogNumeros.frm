VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmLogNumeros 
   BorderStyle     =   0  'None
   Caption         =   "Log Numeros"
   ClientHeight    =   9195
   ClientLeft      =   4005
   ClientTop       =   1320
   ClientWidth     =   7605
   Icon            =   "frmLogNumeros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   6060
      TabIndex        =   15
      Top             =   8550
      Width           =   1485
   End
   Begin Threed.SSPanel pnlBuscar 
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   7830
      Width           =   7605
      _Version        =   65536
      _ExtentX        =   13414
      _ExtentY        =   1931
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.CheckBox chkFiltrarxRango 
         Caption         =   "Filtrar por rango"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   2145
      End
      Begin VB.CheckBox chkVerUltimaEjecucion 
         Caption         =   "Ver solamente la última ejecución"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   420
         Value           =   1  'Checked
         Width           =   2715
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   6060
         TabIndex        =   9
         Top             =   330
         Width           =   1485
      End
      Begin VB.TextBox txtNumeroFinal 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4380
         MaxLength       =   12
         TabIndex        =   8
         Top             =   690
         Width           =   1575
      End
      Begin VB.TextBox txtNumeroInicial 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4380
         MaxLength       =   12
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   7485
         _Version        =   65536
         _ExtentX        =   13203
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Filtro por rango de números para la última ejecución"
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
      Begin MSForms.Label Label1 
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   720
         Width           =   1050
         VariousPropertyBits=   276824091
         Caption         =   "Número Final:"
         Size            =   "1852;344"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblNumeroInicial 
         Height          =   195
         Left            =   3300
         TabIndex        =   5
         Top             =   390
         Width           =   1065
         VariousPropertyBits=   276824091
         Caption         =   "Número Inicial:"
         Size            =   "1879;344"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   7605
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   7605
      _Version        =   65536
      _ExtentX        =   13414
      _ExtentY        =   13414
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
      Begin Threed.SSPanel pnlAvance 
         Height          =   855
         Left            =   720
         TabIndex        =   12
         Top             =   3270
         Visible         =   0   'False
         Width           =   6375
         _Version        =   65536
         _ExtentX        =   11245
         _ExtentY        =   1508
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
         BorderWidth     =   1
         BevelInner      =   1
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   5250
            TabIndex        =   16
            Top             =   420
            Width           =   1035
         End
         Begin Threed.SSPanel pnlPorcentaje 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   420
            Width           =   5025
            _Version        =   65536
            _ExtentX        =   8864
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
            FloodColor      =   16761024
         End
         Begin VB.Label lblMensaje 
            Caption         =   "Cargando Información en el grid..."
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   150
            Width           =   2565
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdLogNumeros 
         Height          =   7575
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   13361
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin Threed.SSPanel pnlTitulo 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      _Version        =   65536
      _ExtentX        =   13414
      _ExtentY        =   450
      _StockProps     =   15
      Caption         =   "Log de la última ejecución"
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
Attribute VB_Name = "frmLogNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection

Private varLogNumeros As colLogNumero
Private varParametrosTelefonia As claParametrosTelefonia

Private varFecha As String
Private varHora As String
Private varCancelar As Boolean

Private Sub chkFiltrarxRango_Click()
    On Error GoTo ErrManager
    
    If Me.chkFiltrarxRango.Value = 1 Then
        Me.txtNumeroFinal.BackColor = &H80000005
        Me.txtNumeroInicial.BackColor = &H80000005
        Me.txtNumeroFinal.Enabled = True
        Me.txtNumeroInicial.Enabled = True
    Else
        Me.txtNumeroFinal.BackColor = &H8000000F
        Me.txtNumeroInicial.BackColor = &H8000000F
        Me.txtNumeroFinal.Enabled = False
        Me.txtNumeroInicial.Enabled = False
        Me.txtNumeroInicial.Text = ""
        Me.txtNumeroFinal.Text = ""
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo ErrManager
    
    If Me.chkFiltrarxRango.Value = 1 Then
        If Trim(Me.txtNumeroInicial.Text) = "" Then
            MsgBox "Debe digitar el número inicial del rango.", vbInformation, App.Title
            Exit Sub
        End If
        
        If Trim(Me.txtNumeroFinal.Text) = "" Then
            MsgBox "Debe digitar el número final del rango.", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = 11
    
    Me.pnlBuscar.Enabled = False
    
    
    
    Set varLogNumeros = Nothing
    Set varLogNumeros = New colLogNumero
    Set varLogNumeros.proConexion = Me.proConexion
    
    varLogNumeros.proFecha = varFecha & " " & varHora
    varLogNumeros.proFiltrarxRango = Me.chkFiltrarxRango.Value
    varLogNumeros.proNumeroInicial = Trim(Me.txtNumeroInicial.Text)
    varLogNumeros.proNumeroFinal = Trim(Me.txtNumeroFinal.Text)
    varLogNumeros.proVerUltimaEjecucion = Me.chkVerUltimaEjecucion.Value
    
    If varLogNumeros.MetConsultarxFecha Then
        varCancelar = False
        Call SubFPintarGrid
    Else
        MsgBox "Error al buscal la información en el LOG.", vbCritical, App.Title
        Screen.MousePointer = 0
        Me.pnlBuscar.Enabled = True
        Exit Sub
    End If
    
    Me.pnlBuscar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Me.pnlBuscar.Enabled = True
    Screen.MousePointer = 0
    SubGMuestraError
End Sub



Private Sub cmdCancelar_Click()
    On Error GoTo ErrManager
    
    varCancelar = True
    
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
    Dim varParametroFecha As String
    Dim varParametroHora As String
    On Error GoTo ErrManager
    
    Call SubFInicializarGrid
    
    'Buscar la fecha de la última ejecución
    varParametroFecha = "Fecha Insercion Numeros"
    varParametroHora = "Hora Insercion Numeros"

    Set varParametrosTelefonia = New claParametrosTelefonia
    Set varParametrosTelefonia.proConexion = Me.proConexion
    
    varParametrosTelefonia.proParametro = varParametroFecha
    
    If varParametrosTelefonia.MetConsultarParametro Then
        varFecha = varParametrosTelefonia.proValor
    Else
        MsgBox "Error al recuperar la fecha de la última ejecución.", vbCritical, App.Title
        Exit Sub
    End If
    
    varParametrosTelefonia.proParametro = varParametroHora
    
    If varParametrosTelefonia.MetConsultarParametro Then
        varHora = varParametrosTelefonia.proValor
    Else
        MsgBox "Error al recuperar la hora de la última ejecución.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGrid()
    On Error GoTo ErrManager
    
    With Me.grdLogNumeros
    
        .Rows = 1
        .Cols = 5
        .Row = 0
        
        .Col = 0
        .ColWidth(0) = 1000
        .CellAlignment = 4
        .TextMatrix(0, 0) = "Ciudad"
        
        .Col = 1
        .ColWidth(1) = 1000
        .CellAlignment = 4
        .TextMatrix(0, 1) = "Numero"
        
        .Col = 2
        .ColWidth(2) = 2000
        .CellAlignment = 4
        .TextMatrix(0, 2) = "Mensaje"
        
        .Col = 3
        .ColWidth(3) = 1000
        .CellAlignment = 4
        .TextMatrix(0, 3) = "Usuario"
        
        .Col = 4
        .ColWidth(4) = 1500
        .CellAlignment = 4
        .TextMatrix(0, 4) = "Fecha"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGrid()
    Dim varContador As Long
    Dim varPorcentaje As Double
    On Error GoTo ErrManager
    
    Me.pnlAvance.Visible = True
    Me.grdLogNumeros.Rows = 1
    Me.grdLogNumeros.Enabled = False
    Me.grdLogNumeros.Redraw = False
    Me.pnlPorcentaje.FloodShowPct = True
    Me.pnlPorcentaje.FloodType = 1
    
    For varContador = 1 To varLogNumeros.Count
        Me.grdLogNumeros.AddItem varLogNumeros.Item(varContador).proRegionCode & vbTab & _
                                 varLogNumeros.Item(varContador).proNumero & vbTab & _
                                 varLogNumeros.Item(varContador).proMensaje & vbTab & _
                                 varLogNumeros.Item(varContador).proUsuario & vbTab & _
                                 varLogNumeros.Item(varContador).proFecha
        
        If (varContador Mod 10) = 0 Then
            varPorcentaje = varContador * 100
            varPorcentaje = varPorcentaje / varLogNumeros.Count
            Me.pnlPorcentaje.FloodPercent = varPorcentaje
        End If
        
        If (varContador Mod 500) = 0 Then
            Me.grdLogNumeros.Redraw = True
            DoEvents
            Me.grdLogNumeros.Redraw = False
        End If
        
        If varCancelar = True Then
            Exit For
        End If
    Next varContador
    
    Me.pnlAvance.Visible = False
    Me.grdLogNumeros.Enabled = True
    Me.grdLogNumeros.Redraw = True
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

Private Sub txtNumeroInicial_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
