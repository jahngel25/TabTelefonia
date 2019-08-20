VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAsignarValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar valores al Parámetro"
   ClientHeight    =   7215
   ClientLeft      =   3990
   ClientTop       =   2865
   ClientWidth     =   7140
   Icon            =   "frmAsignarValores.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7140
   Begin VB.CommandButton cmdCrearValores 
      Caption         =   "&Crear Nuevos Valores..."
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
      Left            =   4890
      TabIndex        =   34
      Top             =   6870
      Width           =   2205
   End
   Begin VB.CommandButton cmdAgregarValor 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3390
      Picture         =   "frmAsignarValores.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Agregar Descuento"
      Top             =   4320
      Width           =   330
   End
   Begin VB.CommandButton cmdQuitarValor 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3390
      Picture         =   "frmAsignarValores.frx":1054
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Quitar Descuento"
      Top             =   5640
      Width           =   330
   End
   Begin MSFlexGridLib.MSFlexGrid grdSinAsignar 
      Height          =   3195
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CheckBox chkInterfase 
      Height          =   195
      Left            =   930
      TabIndex        =   18
      Top             =   480
      Width           =   195
   End
   Begin VB.TextBox txtPosicion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   900
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1560
      Width           =   405
   End
   Begin VB.CheckBox chkObligatorioOT 
      Caption         =   "OT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1230
      TabIndex        =   33
      Top             =   2100
      Width           =   1005
   End
   Begin VB.CheckBox chkObligatorioAtencion 
      Caption         =   "Atención"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1230
      TabIndex        =   32
      Top             =   1860
      Width           =   1005
   End
   Begin VB.CheckBox chkObligatorioVenta 
      Caption         =   "Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1230
      TabIndex        =   13
      Top             =   1620
      Width           =   1005
   End
   Begin VB.TextBox txtTamano 
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
      Left            =   1230
      TabIndex        =   12
      Top             =   1290
      Width           =   405
   End
   Begin VB.TextBox txtMascara 
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
      Left            =   1230
      TabIndex        =   11
      Top             =   990
      Width           =   1845
   End
   Begin VB.TextBox txtTipo 
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
      Left            =   1230
      TabIndex        =   10
      Top             =   690
      Width           =   1845
   End
   Begin VB.TextBox txtEtiqueta 
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
      Left            =   1230
      TabIndex        =   9
      Top             =   390
      Width           =   1845
   End
   Begin VB.TextBox txtCampo 
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
      Left            =   1230
      TabIndex        =   8
      Top             =   90
      Width           =   1845
   End
   Begin MSFlexGridLib.MSFlexGrid grdAsignados 
      Height          =   3195
      Left            =   60
      TabIndex        =   29
      Top             =   60
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmAsignarValores.frx":13DE
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblInterfase 
      AutoSize        =   -1  'True
      Caption         =   "Enviar:"
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
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblPosicionInterfase 
      AutoSize        =   -1  'True
      Caption         =   "Posición:"
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
      Left            =   210
      TabIndex        =   21
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label lblDescripcionBilling 
      Caption         =   "Si se encuentra habilitado indica que este campo debe se enviado en la interfase con Billing."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1470
      TabIndex        =   20
      Top             =   420
      Width           =   2385
   End
   Begin VB.Label lblDescripcionPosicion 
      Caption         =   "Indica la posición en la cual se debe enviar el campo dentro de la interfase."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1470
      TabIndex        =   19
      Top             =   1470
      Width           =   2295
   End
   Begin VB.Label lblTamano 
      Caption         =   "Tamaño:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   1290
      Width           =   645
   End
   Begin VB.Label lblTipo 
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
      Left            =   240
      TabIndex        =   6
      Top             =   690
      Width           =   345
   End
   Begin VB.Label lblEtiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Etiqueta:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   390
      Width           =   615
   End
   Begin VB.Label lblCampo 
      AutoSize        =   -1  'True
      Caption         =   "Campo:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   90
      Width           =   540
   End
   Begin VB.Label lblMascara 
      AutoSize        =   -1  'True
      Caption         =   "Máscara:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   990
      Width           =   675
   End
   Begin VB.Label lblObligatorio 
      AutoSize        =   -1  'True
      Caption         =   "Editable en:"
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
      Left            =   150
      TabIndex        =   2
      Top             =   1590
      Width           =   825
   End
End
Attribute VB_Name = "frmAsignarValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proParametroProducto As claParametroProducto
Public proConexion As ADODB.Connection
Public proValoresAsignados As colValor
Public proValoresSinAsignar As colValor

Private Sub cmdAgregarValor_Click()
    Dim varValoresCampoProducto As claValoresCampoProducto
    On Error GoTo ErrManager
    
    If Me.grdSinAsignar.Row = 0 Then
        MsgBox "Debe seleccionar el valor que desea agregar.", vbInformation, App.Title
        Exit Sub
    End If
    
    Set varValoresCampoProducto = New claValoresCampoProducto
    Set varValoresCampoProducto.proConexion = Me.proConexion
    
    varValoresCampoProducto.proProductNumber = Me.proParametroProducto.proProductNumber
    varValoresCampoProducto.proCampo = Me.proParametroProducto.proCampo
    varValoresCampoProducto.proValorId = Me.proValoresSinAsignar.Item(Me.grdSinAsignar.Row).proValorId
    
    If varValoresCampoProducto.MetInsertar Then
        Set Me.proValoresSinAsignar = New colValor
        Set Me.proValoresSinAsignar.proConexion = Me.proConexion
        
        Me.proValoresSinAsignar.proProductMaster = varValoresCampoProducto.proProductNumber
        Me.proValoresSinAsignar.proCampo = varValoresCampoProducto.proCampo
        
        If Me.proValoresSinAsignar.MetConsultarSinAsignar Then
            Call SubFPintarGridSinAsignar
        Else
            MsgBox "Error al consultar los valores.", vbCritical, App.Title
            Exit Sub
        End If
            
        Set Me.proValoresAsignados = New colValor
        Set Me.proValoresAsignados.proConexion = Me.proConexion
        
        Me.proValoresAsignados.proProductMaster = varValoresCampoProducto.proProductNumber
        Me.proValoresAsignados.proCampo = varValoresCampoProducto.proCampo
        
        If Me.proValoresAsignados.MetConsultarAsignados Then
            Call SubFPintarGridAsignados
        Else
            MsgBox "Error al consultar los valores Asignados.", vbCritical, App.Title
            Exit Sub
        End If
    Else
        MsgBox "Error al Insertar el valor.", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdCrearValores_Click()
    On Error GoTo ErrManager
    
    Set frmValor.proConexion = Me.proConexion
    frmValor.Show vbModal
    
    If Me.proValoresSinAsignar.MetConsultarSinAsignar Then
        Call SubFPintarGridSinAsignar
    Else
        MsgBox "Error al consultar los valores.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdQuitarValor_Click()
    Dim varValoresCampoProducto As claValoresCampoProducto
    
    On Error GoTo ErrManager
    If Me.grdAsignados.Row = 0 Then
        MsgBox "Debe seleccionar el valor que desea Quitar.", vbInformation, App.Title
        Exit Sub
    End If
    
    Set varValoresCampoProducto = New claValoresCampoProducto
    Set varValoresCampoProducto.proConexion = Me.proConexion
    
    varValoresCampoProducto.proProductNumber = Me.proParametroProducto.proProductNumber
    varValoresCampoProducto.proCampo = Me.proParametroProducto.proCampo
    varValoresCampoProducto.proValorId = Me.proValoresAsignados.Item(Me.grdAsignados.Row).proValorId
    
    If varValoresCampoProducto.MetEliminar Then
        Set Me.proValoresSinAsignar = New colValor
        Set Me.proValoresSinAsignar.proConexion = Me.proConexion
        
        Me.proValoresSinAsignar.proProductMaster = varValoresCampoProducto.proProductNumber
        Me.proValoresSinAsignar.proCampo = varValoresCampoProducto.proCampo
        
        If Me.proValoresSinAsignar.MetConsultarSinAsignar Then
            Call SubFPintarGridSinAsignar
        Else
            MsgBox "Error al consultar los valores.", vbCritical, App.Title
            Exit Sub
        End If
            
        Set Me.proValoresAsignados = New colValor
        Set Me.proValoresAsignados.proConexion = Me.proConexion
        
        Me.proValoresAsignados.proProductMaster = varValoresCampoProducto.proProductNumber
        Me.proValoresAsignados.proCampo = varValoresCampoProducto.proCampo
        
        If Me.proValoresAsignados.MetConsultarAsignados Then
            Call SubFPintarGridAsignados
        Else
            MsgBox "Error al consultar los valores Asignados.", vbCritical, App.Title
            Exit Sub
        End If
    Else
        MsgBox "Error al quitar el valor.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    Me.txtCampo.Text = Me.proParametroProducto.proCampo
    Me.txtEtiqueta.Text = Me.proParametroProducto.proEtiqueta
    
    Select Case Me.proParametroProducto.proTipo
        Case "T"
            Me.txtTipo.Text = "Texto"
        Case "L"
            Me.txtTipo.Text = "Lista"
        Case "F"
            Me.txtTipo.Text = "Fecha"
    End Select
    
    Select Case Me.proParametroProducto.proMascara
        Case ""
            Me.txtMascara.Text = ""
        Case "N"
            Me.txtMascara.Text = "Numérico"
        Case "A"
            Me.txtMascara.Text = "AlfaNumérico"
    End Select
    
    Me.txtTamano.Text = Me.proParametroProducto.proTamaño
    
    If Me.proParametroProducto.proObligatorioVenta = True Then
        Me.chkObligatorioVenta.Value = 1
    Else
        Me.chkObligatorioVenta.Value = 0
    End If
    
    If Me.proParametroProducto.proObligatorioAtencion = True Then
        Me.chkObligatorioAtencion.Value = 1
    Else
        Me.chkObligatorioAtencion.Value = 0
    End If
    
    If Me.proParametroProducto.proObligatorioOT = True Then
        Me.chkObligatorioOT.Value = 1
    Else
        Me.chkObligatorioOT.Value = 0
    End If
    
    If Me.proParametroProducto.proIDInterfase = True Then
        Me.chkInterfase.Value = 1
    Else
        Me.chkInterfase.Value = 0
    End If
    
    Me.txtPosicion.Text = Me.proParametroProducto.proPosicionInterfase
    
    Call SubFInicializarGrids
    
    Set Me.proValoresSinAsignar = New colValor
    Set Me.proValoresSinAsignar.proConexion = Me.proConexion
    
    Me.proValoresSinAsignar.proProductMaster = Me.proParametroProducto.proProductNumber
    Me.proValoresSinAsignar.proCampo = Me.proParametroProducto.proCampo
        
    If Me.proValoresSinAsignar.MetConsultarSinAsignar Then
        Call SubFPintarGridSinAsignar
    Else
        MsgBox "Error al consultar los valores.", vbCritical, App.Title
        Exit Sub
    End If
    
    Set Me.proValoresAsignados = New colValor
    Set Me.proValoresAsignados.proConexion = Me.proConexion
    
    Me.proValoresAsignados.proProductMaster = Me.proParametroProducto.proProductNumber
    Me.proValoresAsignados.proCampo = Me.proParametroProducto.proCampo
    
    If Me.proValoresAsignados.MetConsultarAsignados Then
        Call SubFPintarGridAsignados
    Else
        MsgBox "Error al consultar los valores Asignados.", vbCritical, App.Title
        Exit Sub
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFInicializarGrids()
    On Error GoTo ErrManager
    
    With Me.grdSinAsignar
        .Cols = 3
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .TextMatrix(0, 0) = "Codigo"
        .ColWidth(0) = 1000
        
        .Col = 1
        .CellAlignment = 4
        .TextMatrix(0, 1) = "Valor"
        .ColWidth(1) = 1700
        
        .Col = 2
        .CellAlignment = 4
        .TextMatrix(0, 2) = "tiRecordStatus"
        .ColWidth(2) = 0
        
        .Row = 0
        .Col = 0
    End With
    
    With Me.grdAsignados
        .Cols = 3
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .TextMatrix(0, 0) = "Codigo"
        .ColWidth(0) = 1000
        
        .Col = 1
        .CellAlignment = 4
        .TextMatrix(0, 1) = "Valor"
        .ColWidth(1) = 1700
        
        .Col = 2
        .CellAlignment = 4
        .TextMatrix(0, 2) = "tiRecordStatus"
        .ColWidth(2) = 0
        
        .Row = 0
        .Col = 0
    End With
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridSinAsignar()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdSinAsignar.Rows = 1
    For varContador = 1 To Me.proValoresSinAsignar.Count
        Me.grdSinAsignar.AddItem Me.proValoresSinAsignar.Item(varContador).proValorId & vbTab & _
                                 Me.proValoresSinAsignar.Item(varContador).proValorDesc & vbTab & _
                                 Me.proValoresSinAsignar.Item(varContador).proRecordStatus
    Next varContador
    Me.grdSinAsignar.Row = 0
    Me.grdSinAsignar.Col = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFPintarGridAsignados()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdAsignados.Rows = 1
    For varContador = 1 To Me.proValoresAsignados.Count
        Me.grdAsignados.AddItem Me.proValoresAsignados.Item(varContador).proValorId & vbTab & _
                                 Me.proValoresAsignados.Item(varContador).proValorDesc & vbTab & _
                                 Me.proValoresAsignados.Item(varContador).proRecordStatus
    Next varContador
    Me.grdAsignados.Col = 0
    Me.grdAsignados.Row = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub



