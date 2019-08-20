VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmClasificarNumerosSeleccionados 
   Caption         =   "Asignar clasificación a números seleccionados"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6735
   LinkTopic       =   "frmClasificarNumerosSeleccionados"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   3750
      TabIndex        =   6
      Top             =   1785
      Width           =   1920
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   405
      Left            =   1365
      TabIndex        =   5
      Top             =   1770
      Width           =   1920
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   1560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   2752
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
      BevelOuter      =   1
      Begin VB.OptionButton rbRemover 
         Caption         =   "&Remover clasificación"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cbCodigoClasificacion 
         Height          =   315
         Left            =   -105
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1215
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.OptionButton rbAdicionar 
         Caption         =   "A&dicionar clasificación"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton rbEstablecer 
         Caption         =   "&Establecer clasificación"
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cbClasificacion 
         Height          =   315
         ItemData        =   "frmClasificarNumerosSeleccionados.frx":0000
         Left            =   2625
         List            =   "frmClasificarNumerosSeleccionados.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   795
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C&lasificación:"
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmClasificarNumerosSeleccionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' OBJETIVO: permitir asignar clasificaciones manualmente
'****************************************************************
' parametros de entrada: Números a clasificar manualmente, seleccionados en el formulario frmConsultaNumeros
' parametros de salida: actualiza la colecccion de colNumeros
' AUTOR: Informática & Tecnología (PT)
' FECHA: 27/07/2006
'****************************************************************
Option Explicit

'Propiedad que almacena la colección de números
Public proModo As Integer
Public proNumeros As colNumero
Public proGuardado As Boolean
Public proConexion As ADODB.Connection
Private varClasificacion As colClasificacion

Private Sub cmdAceptar_Click()
On Error GoTo ErrorManager
    Dim iClasificacionId As Integer
    Dim iModo As Integer
    Dim iResultado As Integer
    Dim strError As String
    
    If cbClasificacion.Text = "" Then
        MsgBox "Debe elegir una clasificación de la lista desplegable.", vbInformation, App.Title
        Exit Sub
    End If
    
    cbCodigoClasificacion.ListIndex = cbClasificacion.ListIndex
    
    If rbAdicionar.Value Then
        iModo = 1
    ElseIf rbEstablecer.Value Then
        iModo = 2
    ElseIf rbRemover.Value Then
        iModo = 3
    End If
    Me.proModo = iModo
    iClasificacionId = CInt(cbCodigoClasificacion.Text)
    
    Me.proGuardado = False
    Screen.MousePointer = 11
    iResultado = Me.proNumeros.MetClasificacionManual(iClasificacionId, iModo, strError)
    If iResultado <= 0 Then
        MsgBox "No fue posible asignar la clasificación a los números seleccionados", vbInformation + vbOKOnly, App.Title
    ElseIf iResultado <> Me.proNumeros.proSeleccionados Then
        MsgBox "La clasificación fue asignada parcialmente.  " & vbCrLf & "No fue posible asignarles la clasificación a los números:" & vbCrLf & strError, vbInformation + vbOKOnly, App.Title
        Me.proGuardado = True
        Me.Hide
    Else
        If Me.proModo = 3 Then
            MsgBox "La clasificación fue removida exitosamente.", vbInformation, App.Title
        Else
            MsgBox "La clasificación fue asignada exitosamente.", vbInformation, App.Title
        End If
        Me.proGuardado = True
        Me.Hide
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancelar_Click()
    Me.proGuardado = False
    Me.Hide
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager
        
    ''Inicializar combo de Clasificaciones
    Set varClasificacion = New colClasificacion
    Set varClasificacion.proConexion = Me.proConexion
    Set Me.proNumeros.proConexion = Me.proConexion
    Me.proGuardado = False
    
    If varClasificacion.FunGConsulta Then
        Call SubFLlenarComboClasificacion
    Else
        MsgBox "Error al consultar las clasificaciones.", vbCritical, App.Title
        Exit Sub
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Sub SubFLlenarComboClasificacion()
On Error GoTo ErrorManager
    Dim i As Integer
    
    cbClasificacion.Clear
    cbCodigoClasificacion.Clear

    For i = 1 To cbClasificacion.ListCount
        cbClasificacion.RemoveItem i
        cbCodigoClasificacion.RemoveItem i
    Next
    
    For i = 1 To varClasificacion.Count
        cbClasificacion.AddItem varClasificacion.Item(i).proClasificacion
        cbCodigoClasificacion.AddItem CStr(varClasificacion.Item(i).proClasificacionId)
    Next
    
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


