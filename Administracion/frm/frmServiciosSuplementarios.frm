VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmServiciosSuplementarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Servicios suplementarios"
   ClientHeight    =   7125
   ClientLeft      =   3750
   ClientTop       =   4035
   ClientWidth     =   6750
   Icon            =   "frmServiciosSuplementarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   6750
   Begin Threed.SSPanel SSPanel6 
      Height          =   1395
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   2461
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
      Begin VB.Image Image1 
         Height          =   480
         Left            =   90
         Picture         =   "frmServiciosSuplementarios.frx":0CCA
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.Frame fraFondoFiltro 
      Height          =   555
      Left            =   540
      TabIndex        =   16
      Top             =   0
      Width           =   6195
      Begin VB.ComboBox cboCodigoProducto 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboNombreProducto 
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   180
         Width           =   4905
      End
      Begin VB.Label lblProducto 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   17
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C09258&
      Caption         =   "Edición de Valores"
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
      Left            =   540
      TabIndex        =   5
      Top             =   5250
      Width           =   6225
   End
   Begin VB.Frame FraDatosGenerales 
      BackColor       =   &H00C09258&
      Caption         =   "Datos Generales"
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
      Left            =   510
      TabIndex        =   0
      Top             =   570
      Width           =   6255
   End
   Begin VB.Frame fraFondoValor 
      Height          =   3885
      Left            =   540
      TabIndex        =   1
      Top             =   840
      Width           =   6195
      Begin MSFlexGridLib.MSFlexGrid grdServiciosSup 
         Height          =   3675
         Left            =   30
         TabIndex        =   2
         Top             =   180
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   6482
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame fraFondoBotones 
      Height          =   495
      Left            =   540
      TabIndex        =   3
      Top             =   4740
      Width           =   6195
      Begin VB.CommandButton cmdValores 
         Caption         =   "&Valores Servicio"
         Enabled         =   0   'False
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
         Left            =   3090
         TabIndex        =   23
         ToolTipText     =   "Nuevo Tramo"
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
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
         Left            =   4530
         TabIndex        =   6
         ToolTipText     =   "Nuevo Tramo"
         Top             =   150
         Width           =   1245
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
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
         Left            =   30
         TabIndex        =   4
         ToolTipText     =   "Nuevo Tramo"
         Top             =   150
         Width           =   1185
      End
   End
   Begin VB.Frame fraFondoEdicion 
      Height          =   1665
      Left            =   540
      TabIndex        =   7
      Top             =   5400
      Width           =   6195
      Begin VB.ComboBox cboCodTipo 
         Height          =   315
         ItemData        =   "frmServiciosSuplementarios.frx":1994
         Left            =   5220
         List            =   "frmServiciosSuplementarios.frx":19A1
         TabIndex        =   24
         Top             =   870
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox cboTipoServicio 
         Height          =   315
         ItemData        =   "frmServiciosSuplementarios.frx":19AE
         Left            =   1410
         List            =   "frmServiciosSuplementarios.frx":19BB
         TabIndex        =   22
         Top             =   870
         Width           =   4335
      End
      Begin VB.Frame fraBotones 
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   1170
         Width           =   5805
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
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
            Left            =   4560
            TabIndex        =   15
            ToolTipText     =   "Nuevo Tramo"
            Top             =   150
            Width           =   1185
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "&Guardar"
            Enabled         =   0   'False
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
            Left            =   60
            TabIndex        =   14
            ToolTipText     =   "Nuevo Tramo"
            Top             =   150
            Width           =   1185
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
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
            Left            =   1260
            TabIndex        =   13
            ToolTipText     =   "Nuevo Tramo"
            Top             =   150
            Width           =   1185
         End
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   11
         Top             =   480
         Width           =   4305
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   10
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   21
         Top             =   870
         Width           =   345
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   9
         Top             =   510
         Width           =   900
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   8
         Top             =   180
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmServiciosSuplementarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public proConexion As ADODB.Connection
Public proProducto As colProductMaster
Public proServiciosSup As colServiciosSup
Public proclaServiciosSup As claServiciosSup

'Parámetros
Public proProductNumber As String
Public proCampo As String

Public proCampoPadre As String



Private Sub cboCodigoProducto_Click()
    On Error GoTo ErrorManager
    
    proProductNumber = Me.cboCodigoProducto.List(cboCodigoProducto.ListIndex)
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cboNombreProducto_Click()
    Dim i As Byte
        On Error GoTo ErrorManager
    Me.cboCodigoProducto.ListIndex = cboNombreProducto.ListIndex
    If cboNombreProducto.ListIndex = -1 Then
        MsgBox "Debe seleccionar un Producto de la lista"
    Else
       Me.cmdNuevo.Enabled = True
        proProductNumber = cboCodigoProducto.List(cboCodigoProducto.ListIndex)
        Call SubFInicializarGrid
        If Me.proServiciosSup Is Nothing Then
            Set Me.proServiciosSup = New colServiciosSup
        End If
        Set Me.proServiciosSup.proConexion = Me.proConexion
        proServiciosSup.prochProductNumber = proProductNumber
        If Me.proServiciosSup.FunGConsulta Then
            Call SubFPintarGrid
        Else
            MsgBox "Error al consultar los Servicios Suplementarios existentes.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboProductos()
    
    On Error GoTo ErrorManager
    
    Set Me.proProducto = New colProductMaster
    Set Me.proProducto.proConexion = Me.proConexion
    
    If Me.proProducto.MetConsultar Then
        Call SubFPintarComboProductos
    Else
        MsgBox "Error al consultar los productos.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarComboProductos()
    
    Dim varContador As Integer
    On Error GoTo ErrorManager
    
    Me.cboCodigoProducto.Clear
    Me.cboNombreProducto.Clear
    
    For varContador = 1 To Me.proProducto.Count
        Me.cboNombreProducto.AddItem Me.proProducto.Item(varContador).proDescription
        Me.cboCodigoProducto.AddItem Me.proProducto.Item(varContador).proProductNumber
    Next varContador
    
    Me.cboCodigoProducto.ListIndex = -1
    Me.cboNombreProducto.ListIndex = -1
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub



Private Sub cboTipoServicio_Click()
    Me.cboCodTipo.ListIndex = Me.cboTipoServicio.ListIndex
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrManager
    
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtDescripcion.Enabled = False
    
    
    
    Me.cmdNuevo.Enabled = True
    Me.grdServiciosSup.Enabled = True
    
    Me.cmdGuardar.Enabled = False
    Me.cmdCancelar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
    
    
    If Me.grdServiciosSup.Row > 0 Then
        Me.cmdModificar.Enabled = True
        Me.cmdEliminar.Enabled = True
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdEliminar_Click()
    Dim varValordatos As claServiciosSup
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    If MsgBox("Desea eliminar el valor [" & Me.proServiciosSup.Item(Me.grdServiciosSup.Row).provchNombreServicio & "]?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Set varValordatos = New claServiciosSup
    Set varValordatos.proConexion = Me.proConexion
    varValordatos.proiServicioSuplementarioId = proServiciosSup.Item(Me.grdServiciosSup.Row).proiServicioSuplementarioId
    varValordatos.prochProductNumber = Me.proProductNumber
    
    If varValordatos.FunGEliminar Then
        Set Me.proServiciosSup = Nothing
        Set Me.proServiciosSup = New colServiciosSup
        Set Me.proServiciosSup.proConexion = Me.proConexion
        proServiciosSup.prochProductNumber = proProductNumber
        If Me.proServiciosSup.FunGConsulta Then
            Call SubFPintarGrid
            MsgBox "El registro se eliminó exitosamente.", vbInformation, App.Title
            Call cmdCancelar_Click
        Else
            MsgBox "Error al consultar los ServiciosSup existentes.", vbCritical, App.Title
        End If
    Else
        MsgBox "Error al eliminar el registro.", vbCritical, App.Title
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
    Dim varValordatos As claServiciosSup
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    If Trim(Me.txtDescripcion.Text) = "" Then
        MsgBox "Debe digitar la descripción del valor.", vbInformation, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    If Trim(Me.cboTipoServicio.Text) = "" Then
        MsgBox "Debe seleccionar un tipo.", vbInformation, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    
    Set varValordatos = New claServiciosSup
    Set varValordatos.proConexion = Me.proConexion
    
    varValordatos.proiServicioSuplementarioId = Trim(Me.txtCodigo.Text)
    varValordatos.provchNombreServicio = Trim(Me.txtDescripcion.Text)
    varValordatos.prochProductNumber = Me.proProductNumber
    varValordatos.prochTipoServicio = Me.cboCodTipo.List(cboCodTipo.ListIndex)
    
    
    If varValordatos.FunGGuardar Then
        Set Me.proServiciosSup = Nothing
        Set Me.proServiciosSup = New colServiciosSup
        Set Me.proServiciosSup.proConexion = Me.proConexion
        proServiciosSup.prochProductNumber = proProductNumber
        If Me.proServiciosSup.FunGConsulta Then
            Call SubFPintarGrid
            MsgBox "El registro se actualizó exitosamente.", vbInformation, App.Title
            Call cmdCancelar_Click
        Else
            MsgBox "Error al consultar los ServiciosSup existentes.", vbCritical, App.Title
        End If
    Else
        MsgBox "Error al actualizar la informacion.", vbCritical, App.Title
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub


Private Sub cmdModificar_Click()
    On Error GoTo ErrManager
    
    Me.txtDescripcion.Enabled = True
    
    Me.cmdNuevo.Enabled = False
    Me.grdServiciosSup.Enabled = False
    
    Me.cmdModificar.Enabled = False
    
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdEliminar.Enabled = False
    
    
    Me.txtCodigo.Text = Me.proServiciosSup.Item(Me.grdServiciosSup.Row).proiServicioSuplementarioId
    Me.txtDescripcion.Text = Me.proServiciosSup.Item(Me.grdServiciosSup.Row).provchNombreServicio
    
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Sub cmdNuevo_Click()
    On Error GoTo ErrManager
    
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtDescripcion.Enabled = True
    
    Me.cmdNuevo.Enabled = False
    Me.grdServiciosSup.Enabled = False
    
    Me.cmdModificar.Enabled = False
    
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdEliminar.Enabled = False
    
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdValores_Click()

On Error GoTo ErrManager
        
    If UCase(Me.proServiciosSup.Item(Me.grdServiciosSup.Row).prochTipoServicio) = "T" Or UCase(Me.proServiciosSup.Item(Me.grdServiciosSup.Row).prochTipoServicio) = "C" Then
        MsgBox "Para este tipo de servicio no se deben llenar valores, solo aplica para tipos combos.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    End If
    
    Set frmValoresServSuplementarios.proConexion = Me.proConexion
    
    frmValoresServSuplementarios.proServicioSuplementarioId = Me.proServiciosSup.Item(Me.grdServiciosSup.Row).proiServicioSuplementarioId
    frmValoresServSuplementarios.proProductNumber = proProductNumber
    frmValoresServSuplementarios.Show vbModal
    
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    
    Call SubFLlenarComboProductos
    
       
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGrid()
    On Error GoTo ErrManager
    
    With Me.grdServiciosSup
        .Cols = 4
        .Rows = 1
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 2000
        .CellAlignment = 4
        .TextMatrix(0, 0) = "Código"
        
        .Col = 1
        .ColWidth(1) = 3360
        .CellAlignment = 4
        .TextMatrix(0, 1) = "Descripción"
        
        .Col = 2
        .ColWidth(2) = 1000
        .CellAlignment = 4
        .TextMatrix(0, 2) = "Tipo"
        
        .Col = 3
        .ColWidth(3) = 0
        .CellAlignment = 3
        .TextMatrix(0, 3) = "Activo"
        
        .SelectionMode = flexSelectionByRow
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGrid()
    Dim varContador As Integer
    Dim varColumna As Integer
    Dim varTipoServicio As String
    
    On Error GoTo ErrManager
    
    Me.grdServiciosSup.Rows = 1
    For varContador = 1 To Me.proServiciosSup.Count
        If Me.proServiciosSup.Item(varContador).prochTipoServicio = "C" Then
            varTipoServicio = "Checkbox"
        ElseIf Me.proServiciosSup.Item(varContador).prochTipoServicio = "L" Then
                    varTipoServicio = "Combo"
        ElseIf Me.proServiciosSup.Item(varContador).prochTipoServicio = "T" Then
                    varTipoServicio = "Texto"
        End If
        
        Me.grdServiciosSup.AddItem Me.proServiciosSup.Item(varContador).proiServicioSuplementarioId & vbTab & _
                              Me.proServiciosSup.Item(varContador).provchNombreServicio & vbTab & _
                              varTipoServicio
    Next varContador
    
    Me.cmdModificar.Enabled = False
    Me.cmdEliminar.Enabled = False
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.cboCodTipo.ListIndex = -1
    Me.cboTipoServicio.ListIndex = -1
    Me.grdServiciosSup.Row = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdServiciosSup_Click()
    On Error GoTo ErrManager
    If Me.grdServiciosSup.Row > 0 And (proServiciosSup.Count > 0) Then
        Me.cmdModificar.Enabled = True
        Me.cmdEliminar.Enabled = True
        Me.cmdValores.Enabled = True
        'Me.proValorId = Me.proServiciosSup.Item(Me.grdServiciosSup.Row).proiServicioSuplementarioId
        Me.txtCodigo.Text = Me.proServiciosSup.Item(Me.grdServiciosSup.Row).proiServicioSuplementarioId
        Me.txtDescripcion.Text = Me.proServiciosSup.Item(Me.grdServiciosSup.Row).provchNombreServicio
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub txtDescripcion_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtDescripcion.SelStart = 0
    Me.txtDescripcion.SelLength = Len(Me.txtDescripcion.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeAlfaNumerico(KeyAscii, 0)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
