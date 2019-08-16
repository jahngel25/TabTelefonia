VERSION 5.00
Begin VB.Form frmEdicionUsersClasificacion 
   Caption         =   "Edición Usuario por Clasificación"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGuardar 
      Caption         =   "&Guardar"
      Default         =   -1  'True
      Height          =   345
      Left            =   6000
      TabIndex        =   7
      Top             =   1290
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7410
      Begin VB.ComboBox cmbClasificacionCodigo 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cmbUsuarioCodigo 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   765
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cmbUsuario 
         Height          =   315
         ItemData        =   "frmEdicionUsersClasificacion.frx":0000
         Left            =   1110
         List            =   "frmEdicionUsersClasificacion.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Debe seleccionar un usuario"
         Top             =   765
         Width           =   6210
      End
      Begin VB.ComboBox cmbClasificacion 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Debe seleccionar una clasificación"
         Top             =   360
         Width           =   6210
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario"
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   810
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Clasificación"
         Height          =   240
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmEdicionUsersClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection

Private varClasificacion As colClasificacion
Private varUsuarios As colUsuario
Public varUsersClasificacion As claUsersClasificacion


Private Sub btnGuardar_Click()
 On Error GoTo ErrorManager
    
    'Copia los datos de controles a las propiedades de la clase y almacena los datos
    Call FunGuardaDatos
    
    
Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmbClasificacion_Click()
  On Error GoTo ErrManager
    
        Me.cmbClasificacionCodigo.ListIndex = Me.cmbClasificacion.ListIndex
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmbUsuario_Click()
On Error GoTo ErrManager
    
        Me.cmbUsuarioCodigo.ListIndex = Me.cmbUsuario.ListIndex
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    
        ''Inicializar combo de Clasificaciones
        Set varClasificacion = New colClasificacion
        Set varClasificacion.proConexion = Me.proConexion
        
        If varClasificacion.FunGConsulta Then
            Call SubFLlenarComboClasificacion
        Else
            MsgBox "Error al consultar las clasificaciones.", vbCritical, App.Title
            Exit Sub
        End If
        
        'Inicializa combo de usuarios
        Set varUsuarios = New colUsuario
        Set varUsuarios.proConexion = Me.proConexion
        
        If varUsuarios.FunGConsulta Then
            Call SubFLlenarComboUsuarios
        Else
            MsgBox "Error al consultar los usuarios.", vbCritical, App.Title
            Exit Sub
        End If
End Sub
Private Sub SubFLlenarComboClasificacion()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cmbClasificacionCodigo.Clear
    Me.cmbClasificacion.Clear
    
    If varClasificacion.Count > 0 Then
        For varContador = 1 To varClasificacion.Count
            Me.cmbClasificacionCodigo.AddItem varClasificacion.Item(varContador).proClasificacionId
            Me.cmbClasificacion.AddItem varClasificacion.Item(varContador).proClasificacion
        Next
        
        Me.cmbClasificacionCodigo.ListIndex = 0
        Me.cmbClasificacion.ListIndex = 0
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Private Sub SubFLlenarComboUsuarios()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cmbUsuario.Clear
    Me.cmbUsuarioCodigo.Clear
    
    For varContador = 1 To varUsuarios.Count
        Me.cmbUsuarioCodigo.AddItem varUsuarios.Item(varContador).proUserId
        Me.cmbUsuario.AddItem varUsuarios.Item(varContador).proUserName
    Next
    
    Me.cmbUsuarioCodigo.ListIndex = 0
    Me.cmbUsuario.ListIndex = 0
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
Sub FunGuardaDatos()
On Error GoTo ErrorManager
    'Copia la descripción de la UserClasificacion
    Set varUsersClasificacion = New claUsersClasificacion
    Set varUsersClasificacion.proConexion = Me.proConexion
    
    Me.varUsersClasificacion.proUserId = Me.cmbUsuarioCodigo.Text
    Me.varUsersClasificacion.proClasificacionId = Me.cmbClasificacionCodigo.Text
    'Almacena en la base
    If Me.varUsersClasificacion.FunGInsertar = False Then
        MsgBox "No fue posible guardar el usuario por clasificación", vbInformation + vbOKOnly, App.Title
    End If
    If Me.varUsersClasificacion.proResultado <> "0" Then
        MsgBox Me.varUsersClasificacion.proMensaje, vbInformation, App.Title
    Else
        MsgBox "El usuario fue asignado exitosamente.", vbInformation, App.Title
    End If
    
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
