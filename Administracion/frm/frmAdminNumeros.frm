VERSION 5.00
Begin VB.Form frmAdminNumeros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Números"
   ClientHeight    =   3495
   ClientLeft      =   5715
   ClientTop       =   5265
   ClientWidth     =   4455
   Icon            =   "frmAdminNumeros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4455
   Begin VB.CommandButton cmdGenerarNumeros 
      Caption         =   "&Generar Números"
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
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdVerLog 
      Caption         =   "Ver &Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdPorcentajeAvance 
      Caption         =   "&Ver porcentaje de avance"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
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
      Left            =   3540
      TabIndex        =   7
      Top             =   750
      Width           =   855
   End
   Begin VB.ComboBox cboCodigoCiudad 
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
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboNombreCiudad 
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
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   2085
   End
   Begin VB.TextBox txtNumeroFinal 
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
      Left            =   1380
      MaxLength       =   12
      TabIndex        =   2
      Top             =   720
      Width           =   2085
   End
   Begin VB.TextBox txtNumeroInicial 
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
      Left            =   1380
      MaxLength       =   12
      TabIndex        =   1
      Top             =   390
      Width           =   2085
   End
   Begin VB.CommandButton cmdReclasificarNumeros 
      Caption         =   "&Reclasificar Números"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmAdminNumeros.frx":0CCA
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblCiudad 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad:"
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
      Left            =   630
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblNumeroFinal 
      AutoSize        =   -1  'True
      Caption         =   "Número Final:"
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
      Top             =   750
      Width           =   975
   End
   Begin VB.Label lblNumeroInicial 
      AutoSize        =   -1  'True
      Caption         =   "Número Incial:"
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
      TabIndex        =   5
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C09258&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAdminNumeros.frx":1994
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
      Height          =   675
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   4335
   End
End
Attribute VB_Name = "frmAdminNumeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const EstadoProcesoFinalizado = "Proceso Finalizado"

Public proConexion As ADODB.Connection
Public proUsuario As String

Private varciudades As colCiudad
Private varNumero As colNumero
Private varParametroTelefonia As claParametrosTelefonia


Private Sub cboNombreCiudad_Click()
    On Error GoTo ErrManager
    
    Me.cboCodigoCiudad.ListIndex = Me.cboNombreCiudad.ListIndex
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdGenerarNumeros_Click()
    On Error GoTo ErrManager
    
    If Trim(Me.cboCodigoCiudad.Text) = "" Then
        MsgBox "Debe llenar el campo CIUDAD.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Trim(Me.txtNumeroInicial.Text) = "" Then
        MsgBox "Debe llenar el campo NUMERO INICIAL.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Trim(Me.txtNumeroFinal.Text) = "" Then
        MsgBox "Debe llenar el campo NUMERO FINAL.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Val(Me.txtNumeroInicial.Text) > Val(Me.txtNumeroFinal.Text) Then
        MsgBox "El NUMERO INICIAL debe ser menor que el NUMERO FINAL.", vbInformation, App.Title
        Exit Sub
    End If
    
    Set varNumero = New colNumero
    Set varNumero.proConexion = Me.proConexion
    
    varNumero.proNumeroInicial = Me.txtNumeroInicial.Text
    varNumero.proNumeroFinal = Me.txtNumeroFinal.Text
    varNumero.proRegionCode = Me.cboCodigoCiudad.Text
    varNumero.proUsuario = Me.proUsuario
    
    If varNumero.MetGenerarNumeros Then
        MsgBox "El proceso se envío exitosamente. Consulte el avance del proceso."
        Me.cmdGenerarNumeros.Enabled = False
        Me.cmdReclasificarNumeros.Enabled = False
        Me.cmdPorcentajeAvance.Enabled = True
    Else
        MsgBox "Error al envíar el proceso.", vbCritical, App.Title
    End If
    
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub cmdLimpiar_Click()
    On Error GoTo ErrManager
    
    Me.cboNombreCiudad.ListIndex = -1
    Me.txtNumeroInicial.Text = ""
    Me.txtNumeroFinal.Text = ""
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdPorcentajeAvance_Click()
    On Error GoTo ErrManager

    Set frmAvanceNumeros.proConexion = Me.proConexion
    
    frmAvanceNumeros.Show (vbModal)
   
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdReclasificarNumeros_Click()
    Dim varTodos As String
    On Error GoTo ErrManager
    
    varTodos = "N"
    
    'Validar el ingreso de la información
    If Trim(Me.cboCodigoCiudad.Text) <> "" Or _
       Trim(Me.txtNumeroInicial.Text) <> "" Or _
       Trim(Me.txtNumeroFinal.Text) <> "" Then
       
        If Trim(Me.cboCodigoCiudad.Text) = "" Then
            MsgBox "Debe seleccionar la ciudad o limpiar todos los controles.", vbInformation, App.Title
            Exit Sub
        End If
        
        If Trim(Me.txtNumeroInicial.Text) = "" Then
            MsgBox "Debe digitar el valor inicial o limpiar todos los controles.", vbInformation, App.Title
            Exit Sub
        End If
        
        If Trim(Me.txtNumeroFinal.Text) = "" Then
            MsgBox "Debe digitar el valor final o limpiar todos los controles.", vbInformation, App.Title
            Exit Sub
        End If
    Else
        varTodos = "S"
    End If
    
    'Validar si existen registros para reclasificar
    Set varNumero = Nothing
    Set varNumero = New colNumero
    Set varNumero.proConexion = Me.proConexion
    
    varNumero.proCantidadNumeros = 0
    varNumero.proReclasificarTodos = varTodos
    varNumero.proRegionCode = Trim(Me.cboCodigoCiudad.Text)
    varNumero.proNumeroInicial = Trim(Me.txtNumeroInicial.Text)
    varNumero.proNumeroFinal = Trim(Me.txtNumeroFinal.Text)
    varNumero.proUsuario = Me.proUsuario
    If varNumero.MetValidarReclasificacion Then
        If varNumero.proCantidadNumeros = 0 Then
            MsgBox "No existen números a reclasificar.", vbInformation, App.Title
            Exit Sub
        End If
    Else
        MsgBox "Error al validar si exiten números para reclasificar.", vbCritical, App.Title
        Exit Sub
    End If
    
    'Reclasificar números
    If MsgBox("Existen [" & varNumero.proCantidadNumeros & "] número(s) para reclasificar. Desea realizar esta operación en este momento?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        If varNumero.MetReclasificarNumeros Then
            MsgBox "El proceso se envío exitosamente. Consulte el avance del proceso."
            Me.cmdGenerarNumeros.Enabled = False
            Me.cmdReclasificarNumeros.Enabled = False
            Me.cmdPorcentajeAvance.Enabled = True
        Else
            MsgBox "Error al reclasificar los numeros", vbCritical, App.Title
            Exit Sub
        End If
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdVerLog_Click()
    On Error GoTo ErrManager
    
    Set frmLogNumeros.proConexion = Me.proConexion
    frmLogNumeros.Show (vbModal)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
        
    'Inicializar el combo de ciudades
    Set varciudades = New colCiudad
    Set varciudades.proConexion = Me.proConexion
    
    If varciudades.MetConsultar Then
        Call SubFLlenarComboCiudades
    Else
        MsgBox "Error al consultar las ciudades.", vbCritical, App.Title
    End If
     
    'Verificar el estado de ejecución del proceso
    Set varParametroTelefonia = New claParametrosTelefonia
    Set varParametroTelefonia.proConexion = Me.proConexion
    
    varParametroTelefonia.proParametro = "Estado Insercion Numeros"
    
    If varParametroTelefonia.MetConsultarParametro Then
        If varParametroTelefonia.proValor <> EstadoProcesoFinalizado Then
            Me.cmdGenerarNumeros.Enabled = False
            Me.cmdReclasificarNumeros.Enabled = False
            Me.cmdPorcentajeAvance.Enabled = True
        End If
    Else
        MsgBox "Error al consultar el estado de ejecución.", vbCritical, App.Title
        Exit Sub
    End If
    
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

Private Sub SubFLlenarComboCiudades()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cboCodigoCiudad.Clear
    Me.cboNombreCiudad.Clear
    
    For varContador = 1 To varciudades.Count
        Me.cboCodigoCiudad.AddItem varciudades.Item(varContador).proCodigoCiudad
        Me.cboNombreCiudad.AddItem varciudades.Item(varContador).proNombreCiudad
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
