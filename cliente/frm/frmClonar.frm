VERSION 5.00
Begin VB.Form frmClonar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clonar Información del Registro"
   ClientHeight    =   1125
   ClientLeft      =   5550
   ClientTop       =   6210
   ClientWidth     =   3840
   Icon            =   "frmClonar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   810
      Width           =   1305
   End
   Begin VB.CommandButton cmdClonar 
      Caption         =   "&Clonar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   810
      Width           =   1305
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   90
      MaxLength       =   2
      TabIndex        =   3
      Top             =   390
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Digite la cantidad de registros que desea replicar?"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmClonar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public proConexion As ADODB.Connection
Public proOnyx As EDCVoz.claONYX
Public proDatosProducto As claDatosProducto
Public proRegistro As Integer
Public proOrigen As String      'O   Originales
                                'M  Modificados


Private Sub cmdCancelar_Click()
    On Error GoTo ErrManager
    
    Unload Me
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdClonar_Click()
    Dim varContador As Integer
    Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
    On Error GoTo ErrManager
    
    If Trim(Me.txtCantidad.Text) = "0" Then
        MsgBox "La cantidad de registros a clonar debe ser superior a cero.", vbInformation, App.Title
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    'Crear copia del item de la coleccion
    Set varNovedadDetalleDatosProducto = New claNovedadDetalleDatosProducto
    Set varNovedadDetalleDatosProducto.proConexion = Me.proConexion
    varNovedadDetalleDatosProducto.proIncidentId = Me.proDatosProducto.proIncidentId
    varNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId = 0
    varNovedadDetalleDatosProducto.proTipoNovedadId = 1
    varNovedadDetalleDatosProducto.proDetalleDatosProductoId = 0
    If Me.proOrigen = "O" Then
        varNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proDatosProductoId
        varNovedadDetalleDatosProducto.proRecordStatus = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proRecordStatus
        varNovedadDetalleDatosProducto.proStatusId = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proStatusId
        varNovedadDetalleDatosProducto.proUser1 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser1
        varNovedadDetalleDatosProducto.proUser2 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser2
        varNovedadDetalleDatosProducto.proUser3 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser3
        varNovedadDetalleDatosProducto.proUser4 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser4
        varNovedadDetalleDatosProducto.proUser5 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser5
        varNovedadDetalleDatosProducto.proUser6 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser6
        varNovedadDetalleDatosProducto.proUser7 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser7
        varNovedadDetalleDatosProducto.proUser8 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser8
        varNovedadDetalleDatosProducto.proUser9 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser9
        varNovedadDetalleDatosProducto.proUser10 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser10
        varNovedadDetalleDatosProducto.proUser11 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser11
        varNovedadDetalleDatosProducto.proUser12 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser12
        varNovedadDetalleDatosProducto.proUser13 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser13
        varNovedadDetalleDatosProducto.proUser14 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser14
        varNovedadDetalleDatosProducto.proUser15 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser15
        varNovedadDetalleDatosProducto.proUser16 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser16
        varNovedadDetalleDatosProducto.proUser17 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser17
        varNovedadDetalleDatosProducto.proUser18 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser18
        varNovedadDetalleDatosProducto.proUser19 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser19
        varNovedadDetalleDatosProducto.proUser20 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser20
        varNovedadDetalleDatosProducto.proUser21 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser21
        varNovedadDetalleDatosProducto.proUser22 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser22
        varNovedadDetalleDatosProducto.proUser23 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser23
        varNovedadDetalleDatosProducto.proUser24 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser24
        varNovedadDetalleDatosProducto.proUser25 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser25
        varNovedadDetalleDatosProducto.proUser26 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser26
        varNovedadDetalleDatosProducto.proUser27 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser27
        varNovedadDetalleDatosProducto.proUser28 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser28
        varNovedadDetalleDatosProducto.proUser29 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser29
        varNovedadDetalleDatosProducto.proUser30 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser30
        varNovedadDetalleDatosProducto.proUser31 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser31
        varNovedadDetalleDatosProducto.proUser32 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser32
        varNovedadDetalleDatosProducto.proUser33 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser33
        varNovedadDetalleDatosProducto.proUser34 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser34
        varNovedadDetalleDatosProducto.proUser35 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser35
        varNovedadDetalleDatosProducto.proUser36 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser36
        varNovedadDetalleDatosProducto.proUser37 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser37
        varNovedadDetalleDatosProducto.proUser38 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser38
        varNovedadDetalleDatosProducto.proUser39 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser39
        varNovedadDetalleDatosProducto.proUser40 = Me.proDatosProducto.proDetalleDatosProducto.Item(Me.proRegistro).proUser40
            
    Else
        varNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proDatosProductoId
        varNovedadDetalleDatosProducto.proDetalleDatosProductoId = 0
        varNovedadDetalleDatosProducto.proRecordStatus = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proRecordStatus
        varNovedadDetalleDatosProducto.proStatusId = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proStatusId
        varNovedadDetalleDatosProducto.proUser1 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser1
        varNovedadDetalleDatosProducto.proUser2 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser2
        varNovedadDetalleDatosProducto.proUser3 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser3
        varNovedadDetalleDatosProducto.proUser4 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser4
        varNovedadDetalleDatosProducto.proUser5 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser5
        varNovedadDetalleDatosProducto.proUser6 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser6
        varNovedadDetalleDatosProducto.proUser7 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser7
        varNovedadDetalleDatosProducto.proUser8 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser8
        varNovedadDetalleDatosProducto.proUser9 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser9
        varNovedadDetalleDatosProducto.proUser10 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser10
        varNovedadDetalleDatosProducto.proUser11 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser11
        varNovedadDetalleDatosProducto.proUser12 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser12
        varNovedadDetalleDatosProducto.proUser13 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser13
        varNovedadDetalleDatosProducto.proUser14 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser14
        varNovedadDetalleDatosProducto.proUser15 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser15
        varNovedadDetalleDatosProducto.proUser16 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser16
        varNovedadDetalleDatosProducto.proUser17 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser17
        varNovedadDetalleDatosProducto.proUser18 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser18
        varNovedadDetalleDatosProducto.proUser19 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser19
        varNovedadDetalleDatosProducto.proUser20 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser20
        varNovedadDetalleDatosProducto.proUser21 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser21
        varNovedadDetalleDatosProducto.proUser22 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser22
        varNovedadDetalleDatosProducto.proUser23 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser23
        varNovedadDetalleDatosProducto.proUser24 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser24
        varNovedadDetalleDatosProducto.proUser25 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser25
        varNovedadDetalleDatosProducto.proUser26 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser26
        varNovedadDetalleDatosProducto.proUser27 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser27
        varNovedadDetalleDatosProducto.proUser28 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser28
        varNovedadDetalleDatosProducto.proUser29 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser29
        varNovedadDetalleDatosProducto.proUser30 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser30
        varNovedadDetalleDatosProducto.proUser31 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser31
        varNovedadDetalleDatosProducto.proUser32 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser32
        varNovedadDetalleDatosProducto.proUser33 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser33
        varNovedadDetalleDatosProducto.proUser34 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser34
        varNovedadDetalleDatosProducto.proUser35 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser35
        varNovedadDetalleDatosProducto.proUser36 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser36
        varNovedadDetalleDatosProducto.proUser37 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser37
        varNovedadDetalleDatosProducto.proUser38 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser38
        varNovedadDetalleDatosProducto.proUser39 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser39
        varNovedadDetalleDatosProducto.proUser40 = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(Me.proRegistro).proUser40
    End If

    For varContador = 1 To CInt(Trim(Me.txtCantidad.Text))
        varNovedadDetalleDatosProducto.proNovedadDetalleDatosProductoId = 0
        If varNovedadDetalleDatosProducto.MetGuardar Then
            If Not Me.proDatosProducto.MetAgregarNovedadDetalle(varNovedadDetalleDatosProducto) Then
                MsgBox "Error al agregar el detalle.", vbCritical, App.Title
                Screen.MousePointer = 0
                Exit Sub
            End If
        Else
            MsgBox "Error al guardar el detalle.", vbCritical, App.Title
            Screen.MousePointer = 0
            Exit Sub
        End If
    Next varContador
    Screen.MousePointer = 0
    Unload Me
    Exit Sub
ErrManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

Private Sub txtCantidad_Change()
    On Error GoTo ErrManager
    
    If Len(Trim(Me.txtCantidad.Text)) = 0 Then
        Me.cmdClonar.Enabled = False
    Else
        Me.cmdClonar.Enabled = True
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
    KeyAscii = FunGLeeNumerico(KeyAscii)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


