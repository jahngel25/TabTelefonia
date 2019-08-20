VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmRangosNumeros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rangos de numeros por tipo"
   ClientHeight    =   4695
   ClientLeft      =   4665
   ClientTop       =   3375
   ClientWidth     =   6285
   Icon            =   "frmRangosNumeros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel6 
      Height          =   585
      Left            =   0
      TabIndex        =   8
      Top             =   4110
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1032
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
         Picture         =   "frmRangosNumeros.frx":0CCA
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.Frame fraFondoEdicion 
      Height          =   945
      Left            =   0
      TabIndex        =   5
      Top             =   -60
      Width           =   6225
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "&Consultar"
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
         Left            =   2700
         TabIndex        =   2
         ToolTipText     =   "Nuevo Tramo"
         Top             =   540
         Width           =   1065
      End
      Begin VB.TextBox txtInferior 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "0"
         Top             =   150
         Width           =   1245
      End
      Begin VB.TextBox txtSuperior 
         Alignment       =   1  'Right Justify
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
         Left            =   1410
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "999999999"
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Rango inferior"
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
         TabIndex        =   7
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Rango superior"
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
         Left            =   180
         TabIndex        =   6
         Top             =   510
         Width           =   1110
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdRangos 
      Height          =   3315
      Left            =   0
      TabIndex        =   3
      Top             =   900
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5847
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "&Seleccionar"
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
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "Nuevo Tramo"
      Top             =   4320
      Width           =   1185
   End
End
Attribute VB_Name = "frmRangosNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmRangosNumerosFormulario
' Fecha  : 12/10/2004 09:19
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : permitir seleccionar un rango de registros para reservar
'---------------------------------------------------------------------------------------
Option Explicit

Public proRegionCode As String
Public proEstadoNumero As String
Public proInicio As String
' Conexion
Public proConexion As ADODB.Connection

'Clases
Public procolRangosNumeros As colRangosNumeros

Private Sub cmdConsultar_Click()
   On Error GoTo ErrorManager
   Screen.MousePointer = vbHourglass
    If txtInferior.Text <> "" And Me.txtSuperior.Text <> "" Then
        grdRangos.Clear
        PintarGrid
        Call LlenarGrid
    Else
        MsgBox "Seleccione un rango de numeros para buscar"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdSelecionar_Click()
   On Error GoTo ErrorManager

    proInicio = procolRangosNumeros.Item(Me.grdRangos.Row).proInicio
    Me.Hide

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub


Private Sub Form_Activate()
   On Error GoTo ErrorManager

    Call PintarGrid

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Sub PintarGrid()
   On Error GoTo ErrorManager

    grdRangos.Rows = 1
    grdRangos.Cols = 3
    grdRangos.TextMatrix(0, 0) = "Inicio"
    grdRangos.TextMatrix(0, 1) = "Fin"
    grdRangos.TextMatrix(0, 2) = "Cuantos"
    grdRangos.ColWidth(0) = 1600
    grdRangos.ColWidth(1) = 1600
    grdRangos.ColWidth(2) = 1600

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Sub LlenarGrid()

    Dim varContador As Long
    
   On Error GoTo ErrorManager
    grdRangos.Visible = False
    Set procolRangosNumeros = New colRangosNumeros
    Set procolRangosNumeros.proConexion = Me.proConexion
    
    procolRangosNumeros.proNumeroMayor = Me.txtSuperior
    procolRangosNumeros.proNumeroMenor = Me.txtInferior
    procolRangosNumeros.proRegionCode = Me.proRegionCode
    procolRangosNumeros.proEstadoNumero = Me.proEstadoNumero
    If Not procolRangosNumeros.FunGConsulta Then
        MsgBox "No se pudieron consultar los Rangos"
        Exit Sub
    End If
    For varContador = 1 To Me.procolRangosNumeros.Count
        Me.grdRangos.AddItem Me.procolRangosNumeros.Item(varContador).proInicio & vbTab & _
        Me.procolRangosNumeros.Item(varContador).proFin & vbTab & _
        Me.procolRangosNumeros.Item(varContador).proCuantos
    Next varContador
    grdRangos.Visible = True
      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub


Private Sub grdRangos_Click()
   On Error GoTo ErrorManager

    If grdRangos.Row > -1 Then
        Me.cmdSelecionar.Enabled = True
    Else
        Me.cmdSelecionar.Enabled = False
    End If

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
 On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
 On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtInferior_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtInferior.SelStart = 0
    Me.txtInferior.SelLength = Len(Me.txtInferior.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtSuperior_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtSuperior.SelStart = 0
    Me.txtSuperior.SelLength = Len(Me.txtSuperior.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
