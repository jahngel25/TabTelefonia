VERSION 5.00
Begin VB.Form frmEditarSeguridad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Privilegios de usuario"
   ClientHeight    =   2250
   ClientLeft      =   5880
   ClientTop       =   7110
   ClientWidth     =   4125
   Icon            =   "frmEditarSeguridad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraUsuario 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6075
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   900
         TabIndex        =   5
         Text            =   "cboUsuario"
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox txtUsuario 
         Height          =   345
         Left            =   930
         TabIndex        =   6
         Top             =   120
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario :"
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
         Left            =   150
         TabIndex        =   7
         Top             =   150
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Privilegios del Usuario  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   4125
      Begin VB.CheckBox chkAdminNumeros 
         Caption         =   "Administración de Telefonía"
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
         Left            =   450
         TabIndex        =   8
         Top             =   630
         Width           =   3045
      End
      Begin VB.CheckBox chkAdministracion 
         Caption         =   "Administración del Sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   450
         TabIndex        =   3
         Top             =   210
         Width           =   3045
      End
      Begin VB.CheckBox chkNoValidar 
         Caption         =   "No validar proceso ONYX"
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
         Left            =   450
         TabIndex        =   2
         Top             =   960
         Width           =   3045
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
      Left            =   2700
      TabIndex        =   0
      Top             =   1950
      Width           =   1425
   End
End
Attribute VB_Name = "frmEditarSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proSeguridad As claSeguridad
Public proColseguridad As colSeguridad
Public proConexion As ADODB.Connection

Dim varColUsuarios As colUsuario
Dim varNuevo As Boolean

Sub SubFCopiarUsuarios()
Dim varCuenta As Integer
On Error GoTo ErrorManager

        Me.cboUsuario.Clear
        For varCuenta = 1 To varColUsuarios.Count
                Me.cboUsuario.AddItem varColUsuarios(varCuenta).proUserId
        Next varCuenta
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Function FunFCopiaDatosaClase() As Boolean
Dim varCadena As String
On Error GoTo ErrorManager

        Me.proSeguridad.proUserId = Me.cboUsuario.Text
        
        'Administración de permisos
        If Me.chkAdministracion.Value Then
            varCadena = "1"
        Else
            varCadena = "0"
        End If
        'Administración de Números
        If Me.chkAdminNumeros.Value Then
            varCadena = varCadena & "1"
        Else
            varCadena = varCadena & "0"
        End If
        'Validacion del proceso
        If Me.chkNoValidar.Value Then
            varCadena = varCadena & "1"
        Else
            varCadena = varCadena & "0"
        End If
        
        proSeguridad.proPrivilegios = varCadena
        
        FunFCopiaDatosaClase = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Function FunFCopiaClaseaForma() As Boolean
Dim varEncontro As Boolean
Dim varCuenta As Integer
On Error GoTo ErrorManager

        If Not varNuevo Then
                Me.fraUsuario.Enabled = False
        End If
        
        If Len(Trim(Me.proSeguridad.proUserId)) <> 0 Then
                varEncontro = False
                varCuenta = 0
                While varCuenta <= Me.cboUsuario.ListCount And varEncontro = False
                        If Trim(UCase(Me.cboUsuario.List(varCuenta))) = Trim(UCase(Me.proSeguridad.proUserId)) Then
                                varEncontro = True
                                Me.cboUsuario.ListIndex = varCuenta
                        Else
                                varCuenta = varCuenta + 1
                        End If
                Wend
        End If

        'Administración de permisos
        If Mid(Me.proSeguridad.proPrivilegios, 1, 1) = "1" Then
            Me.chkAdministracion.Value = 1
        End If
        
        'Administración de números
        If Mid(Me.proSeguridad.proPrivilegios, 2, 1) = "1" Then
            Me.chkAdminNumeros.Value = 1
        End If
        
        'Validación del proceso
        If Mid(Me.proSeguridad.proPrivilegios, 3, 1) = "1" Then
            Me.chkNoValidar.Value = 1
        End If
        
        FunFCopiaClaseaForma = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Function FunFCruzarUsuarios() As Boolean
Dim varCuenta As Integer
Dim varContador As Integer
Dim varEncontro As Boolean
On Error GoTo ErrorManager

        'ELimina los usuarios de la coleccion total que encuentre en la colección
        'de seguridad
        For varCuenta = 1 To Me.proColseguridad.Count
                varContador = 1
                varEncontro = False
                While varContador <= varColUsuarios.Count And varEncontro = False
                    If Me.proColseguridad.Item(varCuenta).proUserId = _
                            varColUsuarios(varContador).proUserId Then
                            varEncontro = True
                            varColUsuarios.Remove (varContador)
                    Else
                            varContador = varContador + 1
                    End If
                Wend
        Next varCuenta
        
        FunFCruzarUsuarios = True
        Exit Function
        
ErrorManager:
        SubGMuestraError
End Function

Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
Dim varCuenta As Integer
Dim varEncontro As Boolean
Dim varTamaño As Integer
Dim varCadena As String
On Error GoTo ErrorManager

                If KeyAscii = 13 Then
                        Me.chkAdministracion.SetFocus
                        Exit Sub
                End If
                
                If Me.cboUsuario.SelLength > 0 Then
                        If KeyAscii <> 8 Then
                                varCadena = Left(Me.cboUsuario.Text, Me.cboUsuario.SelStart) + Chr(KeyAscii)
                        Else
                                varCadena = Left(Me.cboUsuario.Text, Me.cboUsuario.SelStart - 1)
                        End If
                Else
                        If KeyAscii <> 8 Then
                                varCadena = Me.cboUsuario.Text & Chr(KeyAscii)
                        Else
                                varCadena = Left(Me.cboUsuario.Text, Len(Trim(Me.cboUsuario)) - 1)
                        End If
                End If
                varTamaño = Len(Trim(varCadena))
                
                'Busca cual es la posición que más se acomoda y la ubica en él
                varEncontro = False
                varCuenta = 0
                
                While varCuenta <= Me.cboUsuario.ListCount And varEncontro = False
                        If Trim(UCase(Left(Me.cboUsuario.List(varCuenta), varTamaño))) = Trim(UCase(varCadena)) Then
                                Me.cboUsuario.Text = Me.cboUsuario.List(varCuenta)
                                Me.cboUsuario.ListIndex = varCuenta
                                KeyAscii = 0
                                varEncontro = True
                        Else
                                varCuenta = varCuenta + 1
                        End If
                Wend
                
                If varEncontro Then
                        Me.cboUsuario.SelStart = varTamaño
                        Me.cboUsuario.SelLength = Len(Me.cboUsuario.Text) - varTamaño
                        Me.cboUsuario.Refresh
                End If
                Exit Sub

ErrorManager:
        SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
On Error GoTo ErrorManager
        
        'Valida que exista un usuarios
        If Me.cboUsuario.ListIndex = -1 Then
                MsgBox "Es indispensable indicar un usuario válido. Los usuarios que ya pertenecen a la aplicación no podrán ser reingresados.", vbInformation, App.Title
                Exit Sub
        End If

        FunFCopiaDatosaClase
        
        If varNuevo Then
            If Me.proSeguridad.FunGInsertar = False Then
                    MsgBox "No fue posible agregar al usuario", vbInformation, App.Title
            End If
        Else
            If Me.proSeguridad.FunGModificar = False Then
                    MsgBox "No fue posible modificar el usuario", vbInformation, App.Title
            End If
        End If
        
        MsgBox "Se almacenó exitosamente. Los cambios tendrán efecto la siguiente vez que el usuario ingrese a la aplicación", vbInformation, App.Title
        
        'Descarga la forma
        Unload Me
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager

        'Consulta la colección de usuarios
        Set varColUsuarios = New colUsuario
        Set varColUsuarios.proConexion = Me.proConexion

        'Consulta todos los usuarios
        varColUsuarios.FunGConsulta
        varNuevo = False
        'a a cruzar los usuarios para no mostrar los ya asignados para el caso de nuevo
        If Len(Trim(Me.proSeguridad.proUserId)) = 0 Then
                FunFCruzarUsuarios
                varNuevo = True
        End If
        
        'Copia los usuarios al combo de usuarios
        Call SubFCopiarUsuarios
        
        'Copia la seguridad a la forma
        FunFCopiaClaseaForma
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub



