VERSION 5.00
Begin VB.Form frmEdicionClasificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edicion de clasificación"
   ClientHeight    =   1320
   ClientLeft      =   6270
   ClientTop       =   5355
   ClientWidth     =   6435
   Icon            =   "frmEdicionClasificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   6435
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
      Left            =   4890
      TabIndex        =   1
      Top             =   690
      Width           =   1485
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
      Height          =   255
      Left            =   1110
      MaxLength       =   255
      TabIndex        =   0
      Tag             =   "Es indispensable ingresar el nombre de la moneda"
      Top             =   360
      Width           =   5235
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmEdicionClasificacion.frx":0CCA
      Top             =   60
      Width           =   480
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
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      Top             =   30
      Width           =   1185
   End
   Begin VB.Label lblId 
      BackColor       =   &H00A7811D&
      BackStyle       =   0  'Transparent
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
      Left            =   900
      TabIndex        =   4
      Top             =   30
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00A7811D&
      BackStyle       =   0  'Transparent
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
      Left            =   180
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00A7811D&
      Height          =   1335
      Left            =   -30
      TabIndex        =   6
      Top             =   0
      Width           =   1545
   End
End
Attribute VB_Name = "frmEdicionClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' OBJETIVO: permitir la creacion y modificacion de las Clasificaciones
'****************************************************************
' parametros de entrada: Clasificacion a ser modificada de frmClasificacion
' parametros de salida: actualiza la colecccion de Clasificacion
' AUTOR: Hernan Botache
' FECHA: 02/09/2004
'****************************************************************

Option Explicit

'Propiedad que almacena el objeto clasificacion
Public proClasificacion As claClasificacion
Sub subFCopiarDatosAGUI()
On Error GoTo ErrorManager
    'Copia la descripción de la clasificaciob
    Me.txtId = Me.proClasificacion.proClasificacionId
    If Me.proClasificacion.proClasificacion <> "" Then
        Me.TxtDescripcion.Text = Trim(Me.proClasificacion.proClasificacion)
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Sub subFCopiarDatosAClase()
On Error GoTo ErrorManager
    'Copia la descripción de la clasificacion
    Me.proClasificacion.proClasificacion = Me.TxtDescripcion
    Me.proClasificacion.proClasificacionId = Me.txtId
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
On Error GoTo ErrorManager
    If FunGVerificaDatos(Me.TxtDescripcion) = False Then Exit Sub
    
    'Copia los datos de controles a las propiedades de la clase
    subFCopiarDatosAClase
    'Almacena en la base
    If Me.proClasificacion.FunGGuardar = False Then
        MsgBox "No fue posible guardar los cambios", vbInformation + vbOKOnly, App.Title
    End If
    
    Unload Me
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager
    'Copia los datos de la clase a la ventana
    subFCopiarDatosAGUI
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub








Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorManager

    KeyAscii = FunGLeeAlfaNumerico(KeyAscii)
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

