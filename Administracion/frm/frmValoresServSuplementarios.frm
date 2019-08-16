VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmValoresServSuplementarios 
   Caption         =   "Valores de los Servicios Suplementarios"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFondoEdicion 
      Height          =   1335
      Left            =   540
      TabIndex        =   12
      Top             =   5490
      Width           =   5805
      Begin VB.CheckBox chkDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Default"
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox txtvalor 
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
         TabIndex        =   17
         Top             =   150
         Width           =   4125
      End
      Begin VB.Frame fraBotones 
         Height          =   1455
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   5805
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            Left            =   4560
            TabIndex        =   14
            ToolTipText     =   "Nuevo Tramo"
            Top             =   150
            Width           =   1185
         End
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         TabIndex        =   18
         Top             =   180
         Width           =   435
      End
   End
   Begin VB.Frame fraFondoBotones 
      Height          =   495
      Left            =   540
      TabIndex        =   9
      Top             =   4740
      Width           =   5775
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
         TabIndex        =   11
         ToolTipText     =   "Nuevo Tramo"
         Top             =   150
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
         Left            =   4530
         TabIndex        =   10
         ToolTipText     =   "Nuevo Tramo"
         Top             =   150
         Width           =   1185
      End
   End
   Begin VB.Frame fraFondoValor 
      Height          =   3885
      Left            =   540
      TabIndex        =   7
      Top             =   840
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid grdValorServicios 
         Height          =   3675
         Left            =   30
         TabIndex        =   8
         Top             =   180
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   6482
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
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
      TabIndex        =   6
      Top             =   570
      Width           =   5805
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
      Width           =   5805
   End
   Begin VB.Frame fraFondoFiltro 
      Height          =   555
      Left            =   540
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox cboNombre 
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   4905
      End
      Begin VB.ComboBox cboCodigo 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblServicioSuplementario 
         AutoSize        =   -1  'True
         Caption         =   "Servicio:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   630
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   1395
      Left            =   0
      TabIndex        =   0
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
         Picture         =   "frmValoresServSuplementarios.frx":0000
         Top             =   330
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmValoresServSuplementarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public proConexion As ADODB.Connection
Public proColServiciosSuplementarios As colServiciosSup
Public proColValorServicio As colValorServicio
Public proclaValorServicio  As claValorServicio



'Parámetros
Public proServicioSuplementarioId As Integer
Public proProductNumber As String
Public sValorAnt As String
Public sDefaultAnt As String


Private Sub SubFLlenarComboServiciosSuplementarios()
    
    On Error GoTo ErrorManager
    
    Set Me.proColServiciosSuplementarios = New colServiciosSup
    Set Me.proColServiciosSuplementarios.proConexion = Me.proConexion
    Me.proColServiciosSuplementarios.prochProductNumber = Me.proProductNumber
    If Me.proProductNumber = "" Or IsNull(Me.proProductNumber) Then
        If Me.proColServiciosSuplementarios.FunGConsultaTodos Then
            Call SubFPintarComboServiciosSup
        Else
            MsgBox "Error al consultar los servicios suplementarios.", vbCritical, App.Title
            Exit Sub
        End If
    Else
        If Me.proColServiciosSuplementarios.FunGConsulta Then
            Call SubFPintarComboServiciosSup
        Else
            MsgBox "Error al consultar los servicios suplementarios.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarComboServiciosSup()
    
    Dim varContador As Integer
    Dim iIndice As Integer
    On Error GoTo ErrorManager
    
    Me.cboCodigo.Clear
    Me.cboNombre.Clear
    iIndice = 0
    For varContador = 1 To Me.proColServiciosSuplementarios.Count
        If Not IsNull(proServicioSuplementarioId) Then
            If proServicioSuplementarioId = Me.proColServiciosSuplementarios.Item(varContador).proiServicioSuplementarioId Then
                iIndice = varContador
            End If
        End If
        
        Me.cboNombre.AddItem Me.proColServiciosSuplementarios.Item(varContador).provchNombreServicio
        Me.cboCodigo.AddItem Me.proColServiciosSuplementarios.Item(varContador).proiServicioSuplementarioId
    Next varContador
    Me.cboCodigo.ListIndex = iIndice - 1
    Me.cboNombre.ListIndex = iIndice - 1
   
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub


Private Sub SubFInicializarGrid()
    On Error GoTo ErrManager
    
    With Me.grdValorServicios
        .Cols = 4
        .Rows = 1
        
        .Row = 0
        .Col = 0
        .ColWidth(0) = 1500
        .CellAlignment = 4
        .TextMatrix(0, 0) = "Código Servicio"
        
        .Col = 1
        .ColWidth(1) = 3360
        .CellAlignment = 4
        .TextMatrix(0, 1) = "Valor"
        
        .Col = 2
        .ColWidth(2) = 1000
        .CellAlignment = 4
        .TextMatrix(0, 2) = "Default"
        
        
        .Col = 3
        .ColWidth(3) = 0
        .CellAlignment = 4
        .TextMatrix(0, 3) = "Activo"
        
        .SelectionMode = flexSelectionByRow
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cboNombre_Click()
Dim i As Byte
        On Error GoTo ErrorManager
    Me.cboCodigo.ListIndex = cboNombre.ListIndex
    If cboNombre.ListIndex = -1 Then
        MsgBox "Debe seleccionar un servicio suplementario de la lista"
    Else
       Me.cmdNuevo.Enabled = True
        proServicioSuplementarioId = cboCodigo.List(cboCodigo.ListIndex)
        Call SubFInicializarGrid
        If Me.proColValorServicio Is Nothing Then
            Set Me.proColValorServicio = New colValorServicio
        End If
        Set Me.proColValorServicio.proConexion = Me.proConexion
        proColValorServicio.proServicioSuplementarioId = proServicioSuplementarioId
        If Me.proColValorServicio.FunGConsulta Then
            Call SubFPintarGrid
        Else
            MsgBox "Error al consultar los valores del servicio suplementario seleccionado.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
Private Sub SubFPintarGrid()
    Dim varContador As Integer
    Dim varColumna As Integer
    Dim strDefault As String
    On Error GoTo ErrManager
    
    Me.grdValorServicios.Rows = 1
    For varContador = 1 To Me.proColValorServicio.Count
        If Me.proColValorServicio.Item(varContador).proDefault = True Then
            strDefault = "SI"
        Else
            strDefault = "NO"
        End If
            
        Me.grdValorServicios.AddItem Me.proColValorServicio.Item(varContador).proServicioSuplementarioId & vbTab & _
                              Me.proColValorServicio.Item(varContador).proValor & vbTab & _
                              strDefault & vbTab & _
                              Me.proColValorServicio.Item(varContador).protiRecordStatus
    Next varContador
    
    Me.cmdModificar.Enabled = False
    Me.cmdEliminar.Enabled = False
    Me.txtvalor.Text = ""
    Me.chkDefault.Value = False
    Me.grdValorServicios.Row = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ErrManager
    
    Me.txtvalor.Text = ""
    Me.chkDefault.Value = False
   
    
    Me.cmdNuevo.Enabled = True
    Me.grdValorServicios.Enabled = True
    
    Me.cmdGuardar.Enabled = False
    Me.cmdCancelar.Enabled = False
    Me.cmdEliminar.Enabled = False
    
    
    
    If Me.grdValorServicios.Row > 0 Then
        Me.cmdModificar.Enabled = True
        Me.cmdEliminar.Enabled = True
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdEliminar_Click()
Dim varValorServicio As claValorServicio
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    If MsgBox("Desea eliminar el valor [" & Me.proColValorServicio.Item(Me.grdValorServicios.Row).proValor & "]?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Set varValorServicio = New claValorServicio
    Set varValorServicio.proConexion = Me.proConexion
    varValorServicio.proServicioSuplementarioId = proColValorServicio.Item(Me.grdValorServicios.Row).proServicioSuplementarioId
    varValorServicio.proValor = proColValorServicio.Item(Me.grdValorServicios.Row).proValor
    varValorServicio.proDefault = proColValorServicio.Item(Me.grdValorServicios.Row).proDefault
    
    If varValorServicio.FunGEliminar Then
        Set Me.proColValorServicio = Nothing
        Set Me.proColValorServicio = New colValorServicio
        Set Me.proColValorServicio.proConexion = Me.proConexion
        proColValorServicio.proServicioSuplementarioId = cboCodigo.List(cboCodigo.ListIndex)
        If Me.proColValorServicio.FunGConsulta Then
            Call SubFPintarGrid
            MsgBox "El registro se eliminó exitosamente.", vbInformation, App.Title
            Call cmdCancelar_Click
        Else
            MsgBox "Error al consultar los valores del servicio suplementario existentes.", vbCritical, App.Title
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
    Dim varValorServicio As claValorServicio
    Dim sDefault As String
    Dim varContador As Integer
    Dim sExisteDefault As Integer
    Dim sValorDefault As String
    Dim iCambiarDefault As Integer
    Dim iEsNuevo As Integer
    Dim iRefrescaGrid As Integer
    
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    If Trim(Me.txtvalor.Text) = "" Then
        MsgBox "Debe digitar el valor.", vbInformation, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    If Me.chkDefault.Value Then
        sDefault = "1"
    Else
        sDefault = "0"
    End If
    
    
    Set varValorServicio = New claValorServicio
    Set varValorServicio.proConexion = Me.proConexion
    
    
    
    'Debe validar que en la coleccion de valores para este servicio suplementario solo exista un default
    sExisteDefault = 0
    iCambiarDefault = 0
    iEsNuevo = 0
    iRefrescaGrid = 0
    For varContador = 1 To Me.proColValorServicio.Count
        If Me.proColValorServicio.Item(varContador).proDefault = True Then
            sExisteDefault = 1
            sValorDefault = Me.proColValorServicio.Item(varContador).proValor
        End If
    Next varContador
    
    If Me.sValorAnt <> "" And Me.sDefaultAnt <> "" Then
        'va a moficar uno que ya existe
        If Trim(Me.sValorAnt) <> Trim(Me.txtvalor.Text) Then
         'Modifico la descripción
         'Debe eliminar este y guardar el nuevo
            varValorServicio.proServicioSuplementarioId = proColValorServicio.Item(Me.grdValorServicios.Row).proServicioSuplementarioId
            varValorServicio.proValor = Me.sValorAnt
            varValorServicio.proDefault = Me.sDefaultAnt
            varValorServicio.FunGEliminar
            iEsNuevo = 1
            Me.sValorAnt = ""
            Me.sDefaultAnt = ""
        ElseIf Me.sDefaultAnt <> Me.chkDefault.Value Then
            'Modifico el default
            If Me.chkDefault.Value = 1 And sExisteDefault = 1 Then
                'Ya existe un Default
                iCambiarDefault = 1
            End If
            Me.sDefaultAnt = ""
            Me.sValorAnt = ""
        End If
    Else ' es nuevo
        iEsNuevo = 1
        'Modifico el default
        If Me.chkDefault.Value = 1 And sExisteDefault = 1 Then
            'Ya existe un Default
            iCambiarDefault = 1
        End If
    End If
    'Si es nuevo
    If sExisteDefault = 1 And iCambiarDefault = 1 Then
        If MsgBox("ya existe un valor default. Desea cambiarlo por este nuevo?", vbYesNo + vbInformation, App.Title) = vbYes Then
            'Cambia el dafault del que ya lo tiene
            varValorServicio.proServicioSuplementarioId = Me.proServicioSuplementarioId
            varValorServicio.proValor = sValorDefault
            varValorServicio.proValorAnt = sValorDefault
            varValorServicio.proDefault = "0"
            If varValorServicio.FunGGuardar = False Then
                MsgBox "Error al actualizar la informacion.", vbCritical, App.Title
                Screen.MousePointer = 0
                Exit Sub
            Else
                'Actualiza el default del modificado
                varValorServicio.proServicioSuplementarioId = Me.proServicioSuplementarioId
                varValorServicio.proValor = Me.txtvalor.Text
                varValorServicio.proValorAnt = Me.txtvalor.Text
                varValorServicio.proDefault = "1"
                If varValorServicio.FunGGuardar = False Then
                    MsgBox "Error al actualizar la informacion.", vbCritical, App.Title
                    Screen.MousePointer = 0
                    Exit Sub
                Else
                    iRefrescaGrid = 1
                End If
            End If
        Else
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If iEsNuevo = 1 Then
        'Guarda el nuevo valor
        varValorServicio.proServicioSuplementarioId = Me.proServicioSuplementarioId
        varValorServicio.proValor = Me.txtvalor.Text
        varValorServicio.proValorAnt = ""
        varValorServicio.proDefault = sDefault
        If varValorServicio.FunGGuardar Then
            iRefrescaGrid = 1
        Else
            MsgBox "Error al actualizar la informacion.", vbCritical, App.Title
        End If
    End If
    'Actualiza Grid
    If iRefrescaGrid = 1 Then
        Set Me.proColValorServicio = Nothing
        Set Me.proColValorServicio = New colValorServicio
        Set Me.proColValorServicio.proConexion = Me.proConexion
        proColValorServicio.proServicioSuplementarioId = Me.proServicioSuplementarioId
        If Me.proColValorServicio.FunGConsulta Then
            Call SubFPintarGrid
            MsgBox "El actualizó la información exitosamente.", vbInformation, App.Title
            Call cmdCancelar_Click
        Else
            MsgBox "Error al consultar los valores del servicio suplementario existente.", vbCritical, App.Title
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdModificar_Click()
  On Error GoTo ErrManager
    
    Me.txtvalor.Enabled = True
    
    Me.cmdNuevo.Enabled = False
    Me.grdValorServicios.Enabled = False
    
    Me.cmdModificar.Enabled = False
    
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdEliminar.Enabled = False
    
    Me.txtvalor.Text = Me.proColValorServicio.Item(Me.grdValorServicios.Row).proValor
    sValorAnt = Me.proColValorServicio.Item(Me.grdValorServicios.Row).proValor
    Me.chkDefault.Value = IIf(Me.proColValorServicio.Item(Me.grdValorServicios.Row).proDefault = True, 1, 0)
    sDefaultAnt = Me.chkDefault.Value
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdNuevo_Click()
    'valido si ese tipo de servicio es combo para permitir llenar los valores
    Dim varClaServicioSupl As claServiciosSup
    Set varClaServicioSupl = New claServiciosSup
    
    Set varClaServicioSupl.proConexion = Me.proConexion
    varClaServicioSupl.proiServicioSuplementarioId = Me.cboCodigo.List(cboCodigo.ListIndex)
    varClaServicioSupl.prochTipoServicio = ""
    If varClaServicioSupl.FunGConsulta Then
        If UCase(varClaServicioSupl.prochTipoServicio) = "T" Or UCase(varClaServicioSupl.prochTipoServicio) = "C" Then
            MsgBox "Para este tipo de servicio no se deben llenar valores, solo aplica para tipos combos.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
    Else
        MsgBox "Error al consultar tipo de servicio suplementario.", vbCritical, App.Title
        Exit Sub
    End If
    Me.txtvalor.Enabled = True
    Me.chkDefault.Enabled = True
    Me.cmdGuardar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdEliminar.Enabled = True
    
End Sub

Private Sub Form_Load()
  On Error GoTo ErrManager
    

    Call SubFLlenarComboServiciosSuplementarios
    
       
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub grdValorServicios_Click()
 On Error GoTo ErrManager
    If Me.grdValorServicios.Row > 0 And (proColValorServicio.Count > 0) Then
        Me.cmdModificar.Enabled = True
        Me.cmdEliminar.Enabled = True
        Me.txtvalor.Text = Me.proColValorServicio.Item(Me.grdValorServicios.Row).proValor
        Me.chkDefault.Value = IIf(Me.proColValorServicio.Item(Me.grdValorServicios.Row).proDefault = True, 1, 0)
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
