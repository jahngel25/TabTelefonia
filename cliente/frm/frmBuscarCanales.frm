VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmBuscarCanales 
   Caption         =   "Búsqueda de Canales Activos"
   ClientHeight    =   6585
   ClientLeft      =   540
   ClientTop       =   3555
   ClientWidth     =   14340
   Icon            =   "frmBuscarCanales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   14340
   Begin VB.Frame fraBusqueda 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   14325
      Begin VB.ComboBox cboCategoriaNombre 
         Height          =   315
         ItemData        =   "frmBuscarCanales.frx":0CCA
         Left            =   5880
         List            =   "frmBuscarCanales.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   1965
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "&Regresar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11160
         TabIndex        =   6
         ToolTipText     =   "Refleja las cuotas del Tiempo de Financiación"
         Top             =   1620
         Width           =   1665
      End
      Begin VB.ComboBox cboCategoriaCodigo 
         Height          =   315
         ItemData        =   "frmBuscarCanales.frx":0CCE
         Left            =   5880
         List            =   "frmBuscarCanales.frx":0CD0
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1140
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtIdIncidente 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   2
         Tag             =   "Es indispensable indicar el nombre del descuento"
         Top             =   1140
         Width           =   1965
      End
      Begin VB.Frame fraTitulo 
         BackColor       =   &H00C09258&
         Caption         =   "  Datos de Búsqueda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   14355
      End
      Begin VB.CommandButton cmdBuscarLogCanales 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9360
         TabIndex        =   5
         Top             =   1620
         Width           =   1665
      End
      Begin MSMask.MaskEdBox txtEnlace 
         Height          =   345
         Left            =   1920
         TabIndex        =   4
         Top             =   1560
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "???AA##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox dtFechaDesde 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   0
         Top             =   330
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox dtFechaHasta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblIgual 
         BackStyle       =   0  'Transparent
         Caption         =   "dd/mm/yyyy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D6980A&
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   19
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblIgual 
         BackStyle       =   0  'Transparent
         Caption         =   "dd/mm/yyyy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D6980A&
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   18
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblTipoAsunto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Asunto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3960
         TabIndex        =   17
         Top             =   1140
         Width           =   1785
      End
      Begin VB.Label lblFechaHasta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Asunto Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label lbFechaDesde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Asunto Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   330
         Width           =   1785
      End
      Begin VB.Label lblIdIncidente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id del Incidente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   1140
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enlace"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   1560
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      TabIndex        =   7
      Top             =   1980
      Width           =   14355
      Begin VB.Frame fraTitulo 
         BackColor       =   &H00C09258&
         Caption         =   " Resultado de la Búsqueda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   14385
      End
      Begin MSFlexGridLib.MSFlexGrid grdLogCanales 
         Height          =   3915
         Left            =   90
         TabIndex        =   8
         Top             =   300
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   6906
         _Version        =   393216
         Rows            =   5
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
End
Attribute VB_Name = "frmBuscarCanales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'       DESCRIPCION: Formulario para la busqueda en la tabla de log
'       de canales para un cliente en particular
'       Autor: TOPGROUP S.A.
'       Fecha: 27/07/2009
'       Version:              1.0.000
'       Requerimiento:        3488
'*******************************************************************
Option Explicit

Public varColLogCanales As colLogCanales
Public proCompanyId As String


'Propiedad de conexion
Public proConexion As ADODB.Connection

Private Sub SubFLlenarComboCategorias()
   
    Dim varContador As Integer
    Dim proCategorias As colCategorias
    On Error GoTo ErrManager
    
    Me.cboCategoriaCodigo.Clear
    Me.cboCategoriaNombre.Clear
    
    Me.cboCategoriaNombre.AddItem "<<< TODOS >>>", 0
    Me.cboCategoriaCodigo.AddItem "0", 0
    
    Set proCategorias = New colCategorias
    Set proCategorias.proConexion = Me.proConexion
    
    If (proCategorias.MetConsultar) Then
        For varContador = 1 To proCategorias.Count
            'Saco las OT, porque no se requieren para la busqueda del log
            If (proCategorias.Item(varContador).proCategoriaID <> "1") Then
                Me.cboCategoriaNombre.AddItem proCategorias.Item(varContador).proDescripcion
                Me.cboCategoriaCodigo.AddItem proCategorias.Item(varContador).proCategoriaID
            End If
        Next varContador
    End If
    
    Me.cboCategoriaCodigo.ListIndex = 0
    Me.cboCategoriaNombre.ListIndex = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFInicializarGridLog()
    On Error GoTo ErrManager:
    
    With Me.grdLogCanales
        .Cols = 9
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .ColWidth(0) = 1000
        .TextMatrix(0, 0) = "ID Incidente"
        
        
        .Col = 1
        .CellAlignment = 4
        .ColWidth(1) = 1000
        .TextMatrix(0, 1) = "Usuario"
        
        .Col = 2
        .CellAlignment = 4
        .ColWidth(2) = 1500
        .TextMatrix(0, 2) = "Tipo de Asunto"
        
        .Col = 3
        .CellAlignment = 4
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = "Tipo de Linea"
        
        .Col = 4
        .CellAlignment = 4
        .ColWidth(4) = 1800
        .TextMatrix(0, 4) = "Código de Enlace"
        
        .Col = 5
        .CellAlignment = 4
        .ColWidth(5) = 1800
        .TextMatrix(0, 5) = "Canales Activos"
        
        .Col = 6
        .CellAlignment = 4
        .ColWidth(6) = 1800
        .TextMatrix(0, 6) = "Diferencia C/Activos"
        
        .Col = 7
        .CellAlignment = 4
        .ColWidth(7) = 1800
        .TextMatrix(0, 7) = "Canales Calculados"
        
        .Col = 8
        .CellAlignment = 4
        .ColWidth(8) = 1800
        .TextMatrix(0, 8) = "Fecha"
        
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintaLogCanales()
    Dim varContador As Integer
    Dim varLogCanal As claLogCanales
    On Error GoTo ErrManager
    
    Me.grdLogCanales.Rows = 1
    Me.grdLogCanales.Redraw = False
    For varContador = 1 To Me.varColLogCanales.Count
        
        Set varLogCanal = Nothing
        Set varLogCanal = New claLogCanales
        
        varLogCanal.proIncidentId = varColLogCanales.Item(varContador).proIncidentId
        varLogCanal.proUserId = varColLogCanales.Item(varContador).proUserId
        varLogCanal.proIncidentCategory = varColLogCanales.Item(varContador).proIncidentCategory
        varLogCanal.proTipoLinea = varColLogCanales.Item(varContador).proTipoLinea
        varLogCanal.proSerialNumber = varColLogCanales.Item(varContador).proSerialNumber
        varLogCanal.proCanalesEnUso = varColLogCanales.Item(varContador).proCanalesEnUso
        varLogCanal.proDiferenciaCanales = varColLogCanales.Item(varContador).proDiferenciaCanales
        varLogCanal.proCanalesCalculados = varColLogCanales.Item(varContador).proCanalesCalculados
        varLogCanal.proFechaNovedad = varColLogCanales.Item(varContador).proFechaNovedad
        
        Me.grdLogCanales.AddItem varLogCanal.proIncidentId & vbTab & _
                              varLogCanal.proUserId & vbTab & _
                              varLogCanal.proIncidentCategory & vbTab & _
                              varLogCanal.proTipoLinea & vbTab & _
                              varLogCanal.proSerialNumber & vbTab & _
                              varLogCanal.proCanalesEnUso & vbTab & _
                              varLogCanal.proDiferenciaCanales & vbTab & _
                              varLogCanal.proCanalesCalculados & vbTab & _
                              varLogCanal.proFechaNovedad

                              
    Next varContador
    
    Set varLogCanal = Nothing
    Me.grdLogCanales.Row = 0
    Me.grdLogCanales.Col = 0
    Me.grdLogCanales.Redraw = True
    Exit Sub
    
ErrManager:
    SubGMuestraError
End Sub




Private Sub cboCategoriaNombre_Click()
On Error GoTo ErrManager
    Me.cboCategoriaCodigo.ListIndex = Me.cboCategoriaNombre.ListIndex
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdBuscarLogCanales_Click()
    On Error GoTo ErrManager
    Dim conMascaraFecha As String
    conMascaraFecha = "  /  /    "
    
    Set Me.varColLogCanales = Nothing
    Set Me.varColLogCanales = New colLogCanales
    Set Me.varColLogCanales.proConexion = Me.proConexion
    
    
    'valido que si ingreso fecha desde haya ingresado fecha hasta
    If Me.dtFechaDesde.FormattedText <> conMascaraFecha And Me.dtFechaHasta.FormattedText = conMascaraFecha Then
        MsgBox "Debe especificar Fecha Hasta si especifico Fecha Desde.", vbInformation, App.Title
        Exit Sub
    End If
    
    'valido que si ingreso fecha hasta haya ingresado fecha desde
    If Me.dtFechaDesde.FormattedText = conMascaraFecha And Me.dtFechaHasta.FormattedText <> conMascaraFecha Then
        MsgBox "Debe especificar Fecha Desde si especifico Fecha Hasta.", vbInformation, App.Title
        Exit Sub
    End If
    
    'valido que si ingreso fechaDesde sea una fecha valida
    If Me.dtFechaDesde.Text <> conMascaraFecha And IsDate(Me.dtFechaDesde.Text) = False Then
        MsgBox "Fecha Desde inválida.", vbInformation, App.Title
        Exit Sub
    End If
    
    'valido que si ingreso fechaHasta sea una fecha valida
    If Me.dtFechaHasta.FormattedText <> conMascaraFecha And IsDate(Me.dtFechaHasta.FormattedText) = False Then
        MsgBox "Fecha Hasta inválida.", vbInformation, App.Title
        Exit Sub
    End If
    
    'validar que fecha desde sea menor a fecha hasta
    If Me.dtFechaDesde.FormattedText <> conMascaraFecha And Me.dtFechaHasta.FormattedText <> conMascaraFecha Then
        If CDate(Me.dtFechaDesde.FormattedText) > CDate(Me.dtFechaHasta.FormattedText) Then
            MsgBox "Fecha Desde no puede ser mayor a Fecha Hasta.", vbInformation, App.Title
            Exit Sub
        End If
        'le aplico la conversion a las fechas si son distintas de vacio
        Me.varColLogCanales.proFechaDesde = FunGFechaAMD(Me.dtFechaDesde.FormattedText)
        Me.varColLogCanales.proFechaHasta = FunGFechaAMD(Me.dtFechaHasta.FormattedText)
    End If
    
    
    Screen.MousePointer = 11
    
    
    
    Me.varColLogCanales.proIncidentId = Me.txtIdIncidente.Text
    Me.varColLogCanales.proSerialNumber = Trim(Me.txtEnlace.Text)
    Me.varColLogCanales.proTipoAsunto = IIf(Me.cboCategoriaCodigo.Text <> "0", Me.cboCategoriaCodigo.Text, "")
    Me.varColLogCanales.proCompanyId = Me.proCompanyId
    
    If Me.varColLogCanales.MetConsultarLogCanales Then
        If Me.varColLogCanales.Count = 0 Then
           MsgBox "No se han encontrado resultados para el criterio ingresado", vbInformation, App.Title
           Call SubFPintaLogCanales
           Screen.MousePointer = 0
           Exit Sub
        End If
        
       Call SubFPintaLogCanales
    Else
        MsgBox "Error al buscar la información en el Log de canales.", vbCritical, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdRegresar_Click()
    On Error GoTo ErrManager
    
    Unload Me
    
    Exit Sub
ErrManager:
    SubGMuestraError

End Sub





Private Sub Form_Load()
On Error GoTo ErrorManager
        
    'Inicializar combo de asuntos
    Call SubFLlenarComboCategorias
    Call SubFInicializarGridLog
    Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub






Private Sub txtEnlace_GotFocus()
On Error GoTo ErrorManager

        Me.txtEnlace.SelStart = 0
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtIdIncidente_KeyPress(KeyAscii As Integer)
On Error GoTo ErrManager
    
     KeyAscii = FunGLeeNumerico(KeyAscii)
    
Exit Sub
ErrManager:
    SubGMuestraError
End Sub
