VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmEdicionReglas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de reglas para analisis de digitos"
   ClientHeight    =   3465
   ClientLeft      =   4455
   ClientTop       =   5850
   ClientWidth     =   7725
   Icon            =   "frmEdicionReglas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7725
   Begin Threed.SSPanel SSPanel6 
      Height          =   1035
      Left            =   30
      TabIndex        =   14
      Top             =   120
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1826
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
         Picture         =   "frmEdicionReglas.frx":0CCA
         Top             =   60
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
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
      Left            =   6480
      TabIndex        =   11
      ToolTipText     =   "EliminarTramo"
      Top             =   3180
      Width           =   1245
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   705
      Left            =   420
      TabIndex        =   15
      Top             =   0
      Width           =   7275
      _Version        =   65536
      _ExtentX        =   12832
      _ExtentY        =   1244
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00A7811D&
         Caption         =   $"frmEdicionReglas.frx":1994
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
         Height          =   465
         Left            =   210
         TabIndex        =   16
         Top             =   120
         Width           =   6915
      End
   End
   Begin VB.Frame FraReglas 
      Height          =   2475
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Width           =   7455
      Begin VB.ComboBox CboPosicionCodigo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Debe seleccionar un tipo de sección"
         Top             =   1740
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.ComboBox CboPosicionNombre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Debe seleccionar un tipo de sección"
         Top             =   1740
         Width           =   2880
      End
      Begin VB.CheckBox chkConsecutivo 
         Caption         =   "Consecutivo digitos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   10
         Top             =   2130
         Width           =   1995
      End
      Begin VB.TextBox TxtRepeticiones 
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
         Left            =   1800
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   7
         Tag             =   "Es indispensable ingresar el nombre de la moneda"
         Top             =   1320
         Width           =   645
      End
      Begin VB.TextBox TxtCantidad 
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
         Left            =   1800
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "Es indispensable ingresar el nombre de la moneda"
         Top             =   930
         Width           =   645
      End
      Begin VB.TextBox TxtDescripcion 
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
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   1
         Tag             =   "Es indispensable ingresar el nombre de la moneda"
         Top             =   540
         Width           =   5445
      End
      Begin VB.Label Label4 
         Caption         =   "Posición Digitos"
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
         Left            =   480
         TabIndex        =   9
         Top             =   1740
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Repeticiones"
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
         Left            =   720
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad de digitos"
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
         Left            =   240
         TabIndex        =   6
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
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
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label txtId 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   3
         Top             =   150
         Width           =   675
      End
      Begin VB.Label lblId 
         Caption         =   "ID"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   180
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmEdicionReglas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' OBJETIVO: permitir la creacion y modificacion de las Reglas
'****************************************************************
' parametros de entrada: Regla a ser modificada de frmReglas
' parametros de salida: actualiza la colecccion de Reglas
' AUTOR: Hernan Botache
' FECHA: 03/09/2004
'****************************************************************

Option Explicit

'Propiedad que almacena el objeto clasificacion
Public proRegla As claRegla
Sub subFCopiarDatosAGUI()
On Error GoTo ErrorManager
    'Copia la descripción de la clasificaciob
    Me.txtId = Me.proRegla.proReglaId
    If Me.proRegla.proReglaId <> "" Then
        Me.txtDescripcion.Text = Trim(Me.proRegla.proDescripcion)
        Me.TxtCantidad.Text = Trim(Me.proRegla.proCantidadDigitos)
        Me.TxtRepeticiones.Text = Trim(Me.proRegla.proRepeticiones)
        If Me.proRegla.proPosicionDigitos <> "" Then
            CboPosicionCodigo.Text = Trim(Me.proRegla.proPosicionDigitos)
        End If
        If UCase(Me.proRegla.proConsecutivoDigitos) = "S" Then
            Me.chkConsecutivo.Value = 1
        Else
            Me.chkConsecutivo.Value = 0
        End If
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Sub subFCopiarDatosAClase()
On Error GoTo ErrorManager
    'Copia la descripción de la clasificacion
    Me.proRegla.proDescripcion = Me.txtDescripcion
    Me.proRegla.proReglaId = Me.txtId
    Me.proRegla.proCantidadDigitos = Me.TxtCantidad
    Me.proRegla.proPosicionDigitos = Me.CboPosicionCodigo
    If Me.chkConsecutivo.Value = 1 Then
        Me.proRegla.proConsecutivoDigitos = "S"
    Else
        Me.proRegla.proConsecutivoDigitos = "N"
    End If
    Me.proRegla.proRepeticiones = Me.TxtRepeticiones
    
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cboGuardar_Click()
On Error GoTo ErrorManager
    If FunGVerificaDatos(Me.txtDescripcion) = False Then Exit Sub
    If FunGVerificaDatos(Me.TxtCantidad) = False Then Exit Sub
    If FunGVerificaDatos(Me.TxtRepeticiones) = False Then Exit Sub
    If FunGVerificaDatos(Me.CboPosicionNombre) = False Then Exit Sub
    'Copia los datos de controles a las propiedades de la clase
    subFCopiarDatosAClase
    'Almacena en la base
    If Me.proRegla.FunGGuardar = False Then
        MsgBox "No fue posible guardar los cambios", vbInformation + vbOKOnly, App.Title
    End If
    
    Unload Me
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cboEliminar_Click()

End Sub

Private Sub cmdGuardar_Click()
On Error GoTo ErrorManager
    If FunGVerificaDatos(Me.txtDescripcion) = False Then Exit Sub
    
    'Copia los datos de controles a las propiedades de la clase
    subFCopiarDatosAClase
    'Almacena en la base
    If Me.proRegla.FunGGuardar = False Then
        MsgBox "No fue posible guardar los cambios", vbInformation + vbOKOnly, App.Title
    Else
        MsgBox "Operación realizada con exito.", vbInformation, App.Title
   
    End If
    
    Unload Me
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager
    'Copia los datos de la clase a la ventana
    LlenarComboPosicion
    subFCopiarDatosAGUI
    
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Sub LlenarComboPosicion()
    Me.CboPosicionNombre.AddItem "Comienzo"
    Me.CboPosicionNombre.AddItem "Fin"
    Me.CboPosicionNombre.AddItem "Todos"
    Me.CboPosicionCodigo.AddItem "C"
    Me.CboPosicionCodigo.AddItem "F"
    Me.CboPosicionCodigo.AddItem "T"
    
End Sub
Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
 On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorManager

    KeyAscii = FunGLeeAlfaNumerico(KeyAscii)
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Private Sub TxtRepeticiones_KeyPress(KeyAscii As Integer)
 On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub CboPosicionNombre_Click()
On Error GoTo ErrorManager

    Me.CboPosicionCodigo.ListIndex = Me.CboPosicionNombre.ListIndex
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Private Sub CboPosicionCodigo_Click()
On Error GoTo ErrorManager

    Me.CboPosicionNombre.ListIndex = Me.CboPosicionCodigo.ListIndex
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub



