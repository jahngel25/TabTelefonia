VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmEstratoCiudad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estratos Por Ciudad"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   Icon            =   "frmEstratoCiudad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkVerActivos 
      Caption         =   "&Ver Activos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   3
      ToolTipText     =   "Muestra estratos activos o inactivos"
      Top             =   4575
      Value           =   1  'Checked
      Width           =   1290
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   150
      TabIndex        =   4
      ToolTipText     =   "Nueva configuración del estrato"
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdDesactivar 
      Caption         =   "&Desactivar"
      Height          =   390
      Left            =   1575
      TabIndex        =   5
      ToolTipText     =   "Desactivar la configuración del estrato"
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdActivar 
      Caption         =   "&Activar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Activar la configuración del estrato"
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Height          =   390
      Left            =   3075
      TabIndex        =   1
      ToolTipText     =   "Clic para generar la consulta"
      Top             =   225
      Width           =   1140
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4065
      Left            =   150
      TabIndex        =   7
      Top             =   825
      Width           =   4065
      _Version        =   65536
      _ExtentX        =   7170
      _ExtentY        =   7170
      _StockProps     =   14
      Caption         =   "Estratos Por Ciudad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSFlexGridLib.MSFlexGrid grdEstratosPorCiudad 
         Height          =   3360
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Estratos configurados"
         Top             =   375
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   5927
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin MSForms.ComboBox cmbConsultariTelefoniaCiudadId 
      Height          =   390
      Left            =   1650
      TabIndex        =   9
      Top             =   2250
      Visible         =   0   'False
      Width           =   1740
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3069;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCiudad 
      Height          =   240
      Left            =   225
      TabIndex        =   8
      Top             =   300
      Width           =   690
      VariousPropertyBits=   276824083
      Caption         =   "Ciudad"
      Size            =   "1217;423"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbConsultarCiudad 
      Height          =   390
      Left            =   975
      TabIndex        =   0
      ToolTipText     =   "Seleccione la ciudad a consultar"
      Top             =   225
      Width           =   2040
      VariousPropertyBits=   748701723
      DisplayStyle    =   7
      Size            =   "3598;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmEstratoCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection
Public proEstratoCiudad As colEstratoCiudad
Public proEstratoCiudad1 As claEstratoCiudad

Private Sub cmdActivar_Click()
On Error GoTo ErrorManager
    If Me.grdEstratosPorCiudad.Row = 0 Or Me.grdEstratosPorCiudad.RowHeight(Me.grdEstratosPorCiudad.RowSel) = 0 Then
        MsgBox "Debe seleccionar un estrato a reactivar", vbInformation, App.Title
        Exit Sub
    End If
    If MsgBox("¿Está seguro de activar el estrato?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    Set Me.proEstratoCiudad1 = Me.proEstratoCiudad.Item(Me.grdEstratosPorCiudad.Row)
    proEstratoCiudad1.FunGInsertar
    Consultar
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdConsultar_Click()
    Consultar
End Sub

Private Sub cmdDesactivar_Click()
On Error GoTo ErrorManager
    If Me.grdEstratosPorCiudad.Row = 0 Or Me.grdEstratosPorCiudad.RowHeight(Me.grdEstratosPorCiudad.RowSel) = 0 Then
        MsgBox "Debe seleccionar un estrato a desactivar", vbInformation, App.Title
        Exit Sub
    End If
    If Me.proEstratoCiudad.FunGEliminar(Me.grdEstratosPorCiudad.Row) = False Then
       ' MsgBox "No fue posible desactivar el estrato", vbInformation + vbOKOnly, App.Title
    End If
    Call FunGPintaEstratoCiudad
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdNuevo_Click()
    Set frmNuevoEstratoCiudad.proConexion = Me.proConexion
    frmNuevoEstratoCiudad.Show vbModal
    If frmNuevoEstratoCiudad.proEstratoCiudad.proEstratoCiudadId <> 0 Then
        cmbConsultarCiudad.ListIndex = 0
    End If
    Consultar
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorManager
    Dim objNewMember As colCiudadOnyx
    Set objNewMember = New colCiudadOnyx
    Set objNewMember.proConexion = Me.proConexion
    objNewMember.FunGConsulta
    Call FunGLlenarCombosCiudad(cmbConsultariTelefoniaCiudadId, cmbConsultarCiudad, objNewMember, "Todas las ciudades")
    grdEstratosPorCiudad.Rows = 2
    grdEstratosPorCiudad.Cols = 2
    grdEstratosPorCiudad.FixedRows = 1
    grdEstratosPorCiudad.TextMatrix(0, 0) = "Ciudad"
    grdEstratosPorCiudad.TextMatrix(0, 1) = "Estrato"
    grdEstratosPorCiudad.Rows = 1
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Function FunGPintaEstratoCiudad() As Boolean
Dim varContador As Integer
On Error GoTo ErrorManager
    grdEstratosPorCiudad.Rows = 2
    grdEstratosPorCiudad.Cols = 2
    grdEstratosPorCiudad.TextMatrix(0, 0) = "Ciudad"
    grdEstratosPorCiudad.TextMatrix(0, 1) = "Estrato"
    grdEstratosPorCiudad.FixedRows = 1
    grdEstratosPorCiudad.ColWidth(0) = 1500
    grdEstratosPorCiudad.ColWidth(1) = 1500
    grdEstratosPorCiudad.Rows = 1
    For varContador = 1 To Me.proEstratoCiudad.Count
        Me.grdEstratosPorCiudad.AddItem Me.proEstratoCiudad.Item(varContador).proNombreCiudad & vbTab & _
        Me.proEstratoCiudad.Item(varContador).proNombreEstrato
    Next varContador
    grdEstratosPorCiudad.Row = 0
    grdEstratosPorCiudad.Col = 0
    FunGPintaEstratoCiudad = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Private Sub chkVerActivos_Click()
On Error GoTo ErrorManager
    If chkVerActivos.Value = False Then
        cmdActivar.Enabled = True
        cmdDesactivar.Enabled = False
    Else
        cmdActivar.Enabled = False
        cmdDesactivar.Enabled = True
    End If
    Consultar
    FunGPintaEstratoCiudad
    Exit Sub
    
ErrorManager:
        SubGMuestraError
End Sub

Function Consultar()
    Dim varValor As Integer
    Set proEstratoCiudad = New colEstratoCiudad
    Set proEstratoCiudad.proConexion = Me.proConexion
    cmbConsultariTelefoniaCiudadId.ListIndex = cmbConsultarCiudad.ListIndex
    If chkVerActivos.Value = 0 Then
        varValor = 0
    Else
        varValor = 1
    End If
    If cmbConsultariTelefoniaCiudadId.Value = "" Then
        cmbConsultariTelefoniaCiudadId.Value = 0
    End If
    proEstratoCiudad.FunGConsulta cmbConsultariTelefoniaCiudadId.Value, varValor
    FunGPintaEstratoCiudad
End Function
