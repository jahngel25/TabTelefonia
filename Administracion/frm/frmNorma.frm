VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNorma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicación de Normas"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmNorma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6945
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdmodificar 
      Caption         =   "&Modificar"
      Height          =   390
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "Modificar los estratos de la configuración"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aplicación de Normas"
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   6615
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
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1290
      End
      Begin MSFlexGridLib.MSFlexGrid grdNorma 
         Height          =   3480
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Normas configuradas"
         Top             =   240
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   6138
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Height          =   390
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Clic para generar la consulta"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdActivar 
      Caption         =   "&Activar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Activar la norma configurada"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDesactivar 
      Caption         =   "&Desactivar"
      Height          =   390
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Desactivar la configuración de la norma"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Nueva configuración de norma"
      Top             =   4920
      Width           =   1215
   End
   Begin MSForms.ComboBox cbociudadId 
      Height          =   390
      Left            =   3480
      TabIndex        =   10
      Top             =   1200
      Width           =   2220
      VariousPropertyBits=   748701723
      DisplayStyle    =   7
      Size            =   "3916;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboCiudadNombre 
      Height          =   390
      Left            =   1095
      TabIndex        =   0
      ToolTipText     =   "Seleccione la ciudad a consultar"
      Top             =   120
      Width           =   3060
      VariousPropertyBits=   748701723
      DisplayStyle    =   7
      Size            =   "5397;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblCiudad 
      Height          =   240
      Left            =   195
      TabIndex        =   7
      Top             =   195
      Width           =   690
      VariousPropertyBits=   276824083
      Caption         =   "Ciudad"
      Size            =   "1217;423"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmNorma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public proConexion As ADODB.Connection
Public proNormas As ColNorma
Public proNorma As claNorma

Private Sub chkVerActivos_Click()
    On Error GoTo ErrorManager
    If chkVerActivos.Value = False Then
        cmdActivar.Enabled = True
        Me.cmdmodificar.Enabled = False
        cmdDesactivar.Enabled = False
    Else
        cmdActivar.Enabled = False
        Me.cmdmodificar.Enabled = True
        cmdDesactivar.Enabled = True
    End If
    Consultar
    Exit Sub
ErrorManager:
        SubGMuestraError
End Sub

Private Sub cmdActivar_Click()
On Error GoTo ErrorManager
    If Me.grdNorma.Row = 0 Or Me.grdNorma.RowHeight(Me.grdNorma.RowSel) = 0 Then
        MsgBox "Debe seleccionar una norma para activar", vbInformation, App.Title
        Exit Sub
    End If
    If MsgBox("¿Está seguro de activar la norma?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    Set Me.proNorma = Me.proNormas.Item(Me.grdNorma.Row)
    Me.proNorma.FunGInsertar
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
    If Me.grdNorma.Row = 0 Or Me.grdNorma.RowHeight(Me.grdNorma.RowSel) = 0 Then
        MsgBox "Debe seleccionar una norma a desactivar", vbInformation, App.Title
        Exit Sub
    End If
    If Me.proNormas.FunGEliminar(Me.grdNorma.Row) = False Then
        MsgBox "No fue posible desactivar la norma", vbInformation + vbOKOnly, App.Title
    End If
    Call FunGPintaNorma
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdModificar_Click()
  On Error GoTo ErrorManager
    If Me.grdNorma.Row = 0 Or Me.grdNorma.RowHeight(Me.grdNorma.RowSel) = 0 Then
        MsgBox "Debe seleccionar una norma del listado", vbInformation, App.Title
        Exit Sub
    End If
    Set frmNuevaNorma.proConexion = Me.proConexion
    frmNuevaNorma.proAccion = "M"
    Set frmNuevaNorma.proNorma = proNormas.Item(Me.grdNorma.Row)
    frmNuevaNorma.Show vbModal
    Consultar
    Exit Sub
ErrorManager:
    SubGMuestraError

End Sub

Private Sub cmdNuevo_Click()
    Set frmNuevaNorma.proConexion = Me.proConexion
    frmNuevaNorma.proAccion = "N"
    frmNuevaNorma.Show vbModal
    If Not frmNuevaNorma.proNorma Is Nothing Then
        If frmNuevaNorma.proNorma.proNormaId <> 0 Then
            Me.cboCiudadNombre.ListIndex = 0
            Consultar
        End If
    End If
End Sub

Private Sub Form_Activate()
    Me.cboCiudadNombre.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorManager
    Dim objNewMember As colCiudadOnyx
    Set objNewMember = New colCiudadOnyx
    Set objNewMember.proConexion = Me.proConexion
    objNewMember.FunGConsulta
    FunGLlenarCombosCiudad Me.cbociudadId, Me.cboCiudadNombre, objNewMember, "Todas las ciudades"
    With grdNorma
        .Rows = 2
        .Cols = 5
        .FixedRows = 1
        .ColWidth(0) = 1300
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ColWidth(3) = 3500
        .ColWidth(4) = 2000
        .Rows = 1
        .TextMatrix(0, 0) = "Ciudad"
        .TextMatrix(0, 1) = "Tipo línea"
        .TextMatrix(0, 2) = "Uso del servicio"
        .TextMatrix(0, 3) = "Norma"
        .TextMatrix(0, 4) = "Estratos"
    End With
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Function FunGPintaNorma() As Boolean
Dim varContador As Integer
On Error GoTo ErrorManager
    With grdNorma
        .Rows = 2
        .Cols = 5
        .FixedRows = 1
        .ColWidth(0) = 1300
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ColWidth(3) = 3500
        .ColWidth(4) = 2000
        .ColAlignment(3) = flexAlignLeftCenter
        .Rows = 1
        .TextMatrix(0, 0) = "Ciudad"
        .TextMatrix(0, 1) = "Tipo línea"
        .TextMatrix(0, 2) = "Uso del servicio"
        .TextMatrix(0, 3) = "Norma"
        .TextMatrix(0, 4) = "Estratos"
    End With
    For varContador = 1 To Me.proNormas.Count
      With proNormas.Item(varContador)
            Me.grdNorma.AddItem .proNombreCiudad & vbTab & _
                                .proTipoLinea & vbTab & _
                                .proUsoServicio & vbTab & _
                                .proCodigoNorma & " " & .proNombreNorma & vbTab & _
                                .proEstratos
        
        End With
    Next varContador
    grdNorma.Row = 0
    grdNorma.Col = 0
    FunGPintaNorma = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Private Sub Consultar()
    Dim varValor As Integer
    Set proNormas = New ColNorma
    Set proNormas.proConexion = Me.proConexion
    Me.cbociudadId.ListIndex = Me.cboCiudadNombre.ListIndex
    If chkVerActivos.Value = 0 Then
        varValor = 0
    Else
        varValor = 1
    End If
    proNormas.FunGConsulta Me.cbociudadId.Value, varValor
    FunGPintaNorma
End Sub
