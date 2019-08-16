VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmValor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de valores"
   ClientHeight    =   6690
   ClientLeft      =   4380
   ClientTop       =   4035
   ClientWidth     =   6720
   Icon            =   "frmValor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6720
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
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Width           =   6645
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
   End
   Begin VB.Frame fraFondoValor 
      Height          =   4485
      Left            =   0
      TabIndex        =   1
      Top             =   150
      Width           =   6675
      Begin MSFlexGridLib.MSFlexGrid grdValores 
         Height          =   4335
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame fraFondoBotones 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4530
      Width           =   6675
      Begin VB.CommandButton cmdActivar 
         Caption         =   "&Activar"
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
         Left            =   4230
         TabIndex        =   15
         ToolTipText     =   "Nuevo Tramo"
         Top             =   180
         Width           =   1185
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
         Left            =   5430
         TabIndex        =   8
         ToolTipText     =   "Nuevo Tramo"
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo / Buscar"
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
         TabIndex        =   6
         ToolTipText     =   "Nuevo Tramo"
         Top             =   150
         Width           =   2085
      End
      Begin VB.Frame fraInactivos 
         Height          =   435
         Left            =   2160
         TabIndex        =   4
         Top             =   60
         Width           =   2025
         Begin VB.CheckBox chkValoresActivos 
            Caption         =   "Ver valores inactivos"
            Height          =   195
            Left            =   150
            TabIndex        =   5
            Top             =   180
            Width           =   1845
         End
      End
   End
   Begin VB.Frame fraFondoEdicion 
      Height          =   1545
      Left            =   0
      TabIndex        =   9
      Top             =   5190
      Width           =   6645
      Begin VB.Frame fraBotones 
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Top             =   1050
         Width           =   6585
         Begin VB.CommandButton cmdInsertar 
            Caption         =   "&Insertar"
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
            Left            =   5340
            TabIndex        =   20
            ToolTipText     =   "Nuevo Tramo"
            Top             =   150
            Width           =   1185
         End
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
            Left            =   4140
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
            ToolTipText     =   "Nuevo Tramo"
            Top             =   150
            Width           =   1185
         End
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1410
         TabIndex        =   13
         Top             =   480
         Width           =   4305
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1410
         TabIndex        =   12
         Top             =   150
         Width           =   1245
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo:"
         Enabled         =   0   'False
         Height          =   345
         Left            =   270
         TabIndex        =   14
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   510
         Width           =   885
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   390
         TabIndex        =   10
         Top             =   180
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection
Public proValores As colValor

'Parámetros
Public proProductNumber As String
Public proCampo As String
Public proPermitirInsertar As Boolean
Public proCampoPadre As String

Public proValorId As String
Public proValorIdPadre As String
Public ProOrden As Integer
Dim Buscando As Boolean


Private Sub chkValoresActivos_Click()
    On Error GoTo ErrManager
        
        If Me.chkValoresActivos.Value = 0 Then
            Me.cmdActivar.Caption = "&Desactivar"
        End If
        
        Call SubFPintarGrid
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdActivar_Click()
    Dim varValordatos As claValor
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    Set varValordatos = New claValor
    Set varValordatos.proConexion = Me.proConexion
    
    varValordatos.proValorId = Me.proValores.Item(Me.grdValores.Row).proValorId
    varValordatos.proValorDesc = Me.proValores.Item(Me.grdValores.Row).proValorDesc
    
    If Me.cmdActivar.Caption = "&Activar" Then
        varValordatos.proRecordStatus = 1
    Else
        varValordatos.proRecordStatus = 0
    End If
    
    If varValordatos.MetModificar Then
        Set Me.proValores = Nothing
        Set Me.proValores = New colValor
        Set Me.proValores.proConexion = Me.proConexion
   
        If Me.proValores.MetConsultar Then
            Call SubFPintarGrid
            MsgBox "El registro se actualizó exitosamente.", vbInformation, App.Title
            Call cmdCancelar_Click
        Else
            MsgBox "Error al consultar los valores existentes.", vbCritical, App.Title
        End If
    Else
        MsgBox "Error al actualizar la informacion.", vbCritical, App.Title
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrManager
    cmdInsertar.Enabled = (Me.grdValores.RowSel > 0)
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtDescripcion.Enabled = False
    Me.chkActivo.Enabled = False
    Me.chkActivo.Value = 0
    
    Me.cmdNuevo.Enabled = True
    'Me.grdValores.Enabled = True
    
    Me.cmdGuardar.Enabled = False
    Me.cmdCancelar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
    Me.chkValoresActivos.Enabled = True
    
    If Me.grdValores.Row > 0 Then
        Me.cmdActivar.Enabled = True
        Me.cmdModificar.Enabled = True
        Me.cmdEliminar.Enabled = True
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdEliminar_Click()
    Dim varValordatos As claValor
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    If MsgBox("Desea eliminar el valor [" & Me.proValores.Item(Me.grdValores.Row).proValorDesc & "]?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Set varValordatos = New claValor
    Set varValordatos.proConexion = Me.proConexion
    
    varValordatos.proValorId = Me.proValores.Item(Me.grdValores.Row).proValorId
    
    If varValordatos.MetEliminar Then
        Set Me.proValores = Nothing
        Set Me.proValores = New colValor
        Set Me.proValores.proConexion = Me.proConexion
   
        If Me.proValores.MetConsultar Then
            Call SubFPintarGrid
            MsgBox "El registro se eliminó exitosamente.", vbInformation, App.Title
            Call cmdCancelar_Click
        Else
            MsgBox "Error al consultar los valores existentes.", vbCritical, App.Title
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
    Dim varValordatos As claValor
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    cmdInsertar.Enabled = (Me.grdValores.RowSel > 0)
    If Trim(Me.txtDescripcion.Text) = "" Then
        MsgBox "Debe digitar la descripción del valor.", vbInformation, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Set varValordatos = New claValor
    Set varValordatos.proConexion = Me.proConexion
    
    varValordatos.proValorId = Trim(Me.txtCodigo.Text)
    varValordatos.proValorDesc = Trim(Me.txtDescripcion.Text)
    varValordatos.proRecordStatus = Val(Me.chkActivo.Value)
    
    If varValordatos.MetModificar Then
        Set Me.proValores = Nothing
        Set Me.proValores = New colValor
        Set Me.proValores.proConexion = Me.proConexion
   
        If Me.proValores.MetConsultar Then
            Call SubFPintarGrid
            MsgBox "El registro se actualizó exitosamente.", vbInformation, App.Title
            Call cmdCancelar_Click
        Else
            MsgBox "Error al consultar los valores existentes.", vbCritical, App.Title
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

Private Sub cmdInsertar_Click()
    
    On Error GoTo ErrManager
    Dim varclaValoresCampoProducto As New claValoresCampoProducto
    Set varclaValoresCampoProducto.proConexion = Me.proConexion
    varclaValoresCampoProducto.proProductNumber = proProductNumber
    varclaValoresCampoProducto.proCampo = Trim(proCampo)
    varclaValoresCampoProducto.proValorId = Me.proValorId
    varclaValoresCampoProducto.proValorIdPadre = Me.proValorIdPadre
    If Not varclaValoresCampoProducto.MetValidarExistencia Then
        varclaValoresCampoProducto.MetInsertar
    End If
    Unload Me
Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdModificar_Click()
    On Error GoTo ErrManager
    
    Me.txtDescripcion.Enabled = True
    Me.chkActivo.Enabled = True
    Me.cmdNuevo.Enabled = False
    'Me.grdValores.Enabled = False
    Me.chkValoresActivos.Enabled = False
    Me.cmdModificar.Enabled = False
    
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdEliminar.Enabled = False
    Me.cmdActivar.Enabled = False
    
    Me.txtCodigo.Text = Me.proValores.Item(Me.grdValores.Row).proValorId
    Me.txtDescripcion.Text = Me.proValores.Item(Me.grdValores.Row).proValorDesc
    
    If Me.proValores.Item(Me.grdValores.Row).proRecordStatus = "1" Then
        Me.chkActivo.Value = 1
        Me.cmdActivar.Caption = "&Desactivar"
    Else
        Me.chkActivo.Value = 0
        Me.cmdActivar.Caption = "&Activar"
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Sub cmdNuevo_Click()
    On Error GoTo ErrManager
    Buscando = True
    Me.txtCodigo.Text = ""
    Me.txtDescripcion.Text = ""
    Me.txtDescripcion.Enabled = True
    Me.chkActivo.Enabled = True
    Me.chkActivo.Value = 1
    
    Me.cmdNuevo.Enabled = False
    'Me.grdValores.Enabled = False
    Me.chkValoresActivos.Enabled = False
    Me.cmdModificar.Enabled = False
    
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdEliminar.Enabled = False
    Me.cmdActivar.Enabled = False
    cmdInsertar.Enabled = False
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    Buscando = False
    Call SubFInicializarGrid
    
    If Me.proValores Is Nothing Then
        Set Me.proValores = New colValor
    End If
    
    Set Me.proValores.proConexion = Me.proConexion
    
    If Me.proValores.MetConsultar Then
        Call SubFPintarGrid
    Else
        MsgBox "Error al consultar los valores existentes.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGrid()
    On Error GoTo ErrManager
    grdValores.Clear
    With Me.grdValores
        .Cols = 3
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
        .ColWidth(2) = 0
        .CellAlignment = 4
        .TextMatrix(0, 2) = "Activo"
        .SelectionMode = flexSelectionByRow
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGrid()
    Dim varContador As Integer
    Dim varColumna As Integer
    On Error GoTo ErrManager
    
    Me.grdValores.Rows = 1
    For varContador = 1 To Me.proValores.Count
        Me.grdValores.AddItem Me.proValores.Item(varContador).proValorId & vbTab & _
                              Me.proValores.Item(varContador).proValorDesc & vbTab & _
                              Me.proValores.Item(varContador).proRecordStatus
                              
        If Me.chkValoresActivos.Value = 1 Then
            If Me.proValores.Item(varContador).proRecordStatus = "0" Then
                Me.grdValores.Row = varContador
                For varColumna = 0 To Me.grdValores.Cols - 1
                    Me.grdValores.Col = varColumna
                    Me.grdValores.CellBackColor = &HC0E0FF
                Next varColumna
                Me.grdValores.RowHeight(varContador) = 240
            End If
        Else
            If Me.proValores.Item(varContador).proRecordStatus = "0" Then
                Me.grdValores.RowHeight(varContador) = 0
            End If
        End If
    Next varContador
    If Not Buscando Then
        Me.cmdActivar.Enabled = False
        Me.cmdModificar.Enabled = False
        Me.cmdEliminar.Enabled = False
        Me.txtCodigo.Text = ""
        Me.txtDescripcion.Text = ""
        Me.chkActivo.Value = 0
    End If
    Me.grdValores.Row = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdValores_Click()
    On Error GoTo ErrManager
    If Me.grdValores.Row > 0 Then
        Buscando = False
        Call cmdCancelar_Click
        Call cmdModificar_Click
        cmdInsertar.Enabled = proPermitirInsertar
        Me.cmdActivar.Enabled = True
        Me.cmdModificar.Enabled = True
        Me.cmdEliminar.Enabled = True
        Me.proValorId = Me.proValores.Item(Me.grdValores.Row).proValorId
        Me.txtCodigo.Text = Me.proValores.Item(Me.grdValores.Row).proValorId
        Me.txtDescripcion.Text = Me.proValores.Item(Me.grdValores.Row).proValorDesc
        If Me.proValores.Item(Me.grdValores.Row).proRecordStatus = "1" Then
            Me.chkActivo.Value = 1
            Me.cmdActivar.Caption = "&Desactivar"
        Else
            Me.chkActivo.Value = 0
            Me.cmdActivar.Caption = "&Activar"
        End If
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub txtDescripcion_Change()

    On Error GoTo ErrManager
    
    If Trim(txtDescripcion.Text) = "" Or Not Buscando Then Exit Sub
    Call SubFInicializarGrid
    If Me.proValores Is Nothing Then
        Set Me.proValores = New colValor
    End If
    
    Set Me.proValores.proConexion = Me.proConexion
    
    If Me.proValores.MetConsultarSemejantes(txtDescripcion.Text) Then
        Call SubFPintarGrid
    Else
        MsgBox "Error al consultar los valores existentes.", vbCritical, App.Title
        Exit Sub
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
