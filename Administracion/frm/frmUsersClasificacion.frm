VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmUsersClasificacion 
   Caption         =   "Usuario por Clasificación"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   Icon            =   "frmUsersClasificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSeguridad 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   6255
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
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
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Crear una nueva clasificación"
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
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
         Left            =   5010
         TabIndex        =   2
         ToolTipText     =   "Eliminar una clasificación"
         Top             =   0
         Width           =   1065
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   585
      Left            =   5160
      TabIndex        =   0
      Top             =   30
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
         Left            =   60
         Picture         =   "frmUsersClasificacion.frx":0CCA
         Top             =   60
         Width           =   480
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdClasificacion 
      Height          =   6435
      Left            =   30
      TabIndex        =   4
      Top             =   270
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   11351
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblCaracteristicas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C09258&
      Caption         =   "Usuario por Clasificación                       "
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
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   6105
   End
End
Attribute VB_Name = "frmUsersClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' OBJETIVO: permitir la administración de usuarios por clasificación
'****************************************************************
' parametros de entrada: ninguna
' parametros de salida: ninguna
' AUTOR: Diana Milena Buenhombre
' FECHA: 30/01/2006
'****************************************************************
Option Explicit

'Propiedad de conexion
Public proConexion As ADODB.Connection

'Colección de usuarios por clasificación
Public proUsersClasificacion As colUsersClasificacion

'Bandera de Inicio de Ventana
Dim varBandera As Integer

Private Sub cmdEliminar_Click()
On Error GoTo ErrorManager

    If Me.grdClasificacion.Row = 0 Or Me.grdClasificacion.RowHeight(Me.grdClasificacion.RowSel) = 0 Then
       MsgBox "Debe seleccionar una clasificacion a eliminar", vbInformation, App.Title
        Exit Sub
    End If
    
    If Me.proUsersClasificacion.FunGEliminarUserClasificacion(Me.grdClasificacion.Row) = False Then
        MsgBox "No fue posible eliminar el usuario de la clasificación", vbInformation + vbOKOnly, App.Title
    End If
    If FunFConsultaUsuarioClasificacion Then
        'Muestra los usuarios por clasificación
        Call FunFPintaUsuersClasificacion
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdNuevo_Click()
Dim varUsersClasificacion As claUsersClasificacion
On Error GoTo ErrorManager

    Set varUsersClasificacion = New claUsersClasificacion
    Set varUsersClasificacion.proConexion = Me.proConexion
    
    Set frmEdicionUsersClasificacion.proConexion = Me.proConexion
     frmEdicionUsersClasificacion.Show vbModal
    
    If FunFConsultaUsuarioClasificacion Then
        'Muestra los usuarios por clasificación
        Call FunFPintaUsuersClasificacion
    End If
    
    Set varUsersClasificacion = Nothing
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorManager

    If varBandera = 0 Then
        Unload Me
    End If
    varBandera = 2
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager

    
    varBandera = 1
    
    'Instancia de la colección
    Set Me.proUsersClasificacion = New colUsersClasificacion
    Set Me.proUsersClasificacion.proConexion = Me.proConexion
    
    'Fubcion que consulta
    
    If FunFConsultaUsuarioClasificacion Then
        'Muestra los usuarios por clasificación
        FunFPintaUsuersClasificacion
    End If
    
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Function FunFPintaUsuersClasificacion() As Boolean
Dim varContador As Integer
On Error GoTo ErrorManager

    'Adecua la grilla
    grdClasificacion.Rows = 2
    grdClasificacion.Cols = 4
    grdClasificacion.TextMatrix(0, 0) = "ID Clasificación"
    grdClasificacion.TextMatrix(0, 1) = "Descripción"
    grdClasificacion.TextMatrix(0, 2) = "ID Usuario"
    grdClasificacion.TextMatrix(0, 3) = "Usuario"
    grdClasificacion.FixedRows = 1
    grdClasificacion.ColWidth(0) = 500
    grdClasificacion.ColWidth(1) = 1900
    grdClasificacion.ColWidth(2) = 1000
    grdClasificacion.ColWidth(3) = 2500
    grdClasificacion.Rows = 1
    
    
    
    For varContador = 1 To Me.proUsersClasificacion.Count
        Me.grdClasificacion.AddItem Me.proUsersClasificacion.Item(varContador).proClasificacionId & vbTab & _
                              Me.proUsersClasificacion.Item(varContador).proClasificacionDescripcion & vbTab & _
                              Me.proUsersClasificacion.Item(varContador).proUserId & vbTab & _
                              Me.proUsersClasificacion.Item(varContador).proUserName
    
    Next varContador
    grdClasificacion.Row = 0
    grdClasificacion.Col = 0
    FunFPintaUsuersClasificacion = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Function FunFConsultaUsuarioClasificacion() As Boolean
Dim varContador As Integer
On Error GoTo ErrorManager

    Me.proUsersClasificacion.proClasificacionId = 0
    Me.proUsersClasificacion.proTodas = "1"
    If Me.proUsersClasificacion.MetConsultarUserClasificacion = False Then
            MsgBox "No fue posible realizar la consulta de usuarios por clasificación", vbInformation + vbOKOnly, App.Title
            varBandera = 0
    End If
    
    FunFConsultaUsuarioClasificacion = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function


