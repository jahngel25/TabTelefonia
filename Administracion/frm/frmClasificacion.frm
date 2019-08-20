VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmClasificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clasificacion"
   ClientHeight    =   7275
   ClientLeft      =   4110
   ClientTop       =   1545
   ClientWidth     =   6105
   Icon            =   "frmClasificacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   6105
   Begin VB.CheckBox chkVerActivos 
      Caption         =   "&Ver activos solamente"
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
      Left            =   30
      TabIndex        =   6
      Top             =   6720
      Value           =   1  'Checked
      Width           =   2235
   End
   Begin VB.Frame fraSeguridad 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   6990
      Width           =   6255
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
         TabIndex        =   3
         ToolTipText     =   "Eliminar una clasificación"
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
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
         Left            =   1020
         TabIndex        =   2
         ToolTipText     =   "Modificar una clasificación"
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdReglas 
         Caption         =   "&Reglas"
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
         Left            =   2460
         TabIndex        =   8
         ToolTipText     =   "Asignar reglas ya creadas a una clasificación"
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdActivar 
         Caption         =   "&Activar"
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
         Left            =   3960
         TabIndex        =   7
         ToolTipText     =   "Activar una clasificación ya eliminada"
         Top             =   0
         Width           =   1065
      End
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
         TabIndex        =   4
         ToolTipText     =   "Crear una nueva clasificación"
         Top             =   0
         Width           =   1035
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdClasificacion 
      Height          =   6435
      Left            =   30
      TabIndex        =   0
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmClasificacion.frx":0CCA
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblCaracteristicas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C09258&
      Caption         =   "Clasificación de la Numeración                          "
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
Attribute VB_Name = "frmClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
' OBJETIVO: permitir la consulta de las clasificaciones
'****************************************************************
' parametros de entrada: ninguna
' parametros de salida: ninguna
' AUTOR: Hernan Botache
' FECHA: 02/09/2004
'****************************************************************
Option Explicit

'Propiedad de conexion
Public proConexion As ADODB.Connection

'Colección de monedas
Public proClasificacion As colClasificacion
'Bandera de Inicio de Ventana
Dim varBandera As Integer

Function FunFPintaClasificacion() As Boolean
Dim varContador As Integer
On Error GoTo ErrorManager

    'Adecua la grilla
    grdClasificacion.Rows = 2
    grdClasificacion.Cols = 3
    grdClasificacion.TextMatrix(0, 0) = "ID Clasificación"
    grdClasificacion.TextMatrix(0, 1) = "Descripción"
    grdClasificacion.TextMatrix(0, 2) = "tiRecordStatus"
    grdClasificacion.FixedRows = 1
    grdClasificacion.ColWidth(0) = 1300
    grdClasificacion.ColWidth(1) = 2700
    grdClasificacion.ColWidth(2) = 1300
    grdClasificacion.Rows = 1
    
    For varContador = 1 To Me.proClasificacion.Count
        Me.grdClasificacion.AddItem Me.proClasificacion.Item(varContador).proClasificacionId & vbTab & _
                              Me.proClasificacion.Item(varContador).proClasificacion & vbTab & _
                              Me.proClasificacion.Item(varContador).proRecordStatus
    
    If Me.proClasificacion.Item(varContador).proRecordStatus = "0" Then
            SUbFPintaEliminados True, Me.chkVerActivos.Value
        End If
    Next varContador
    grdClasificacion.Row = 0
    grdClasificacion.Col = 0
    FunFPintaClasificacion = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function


Private Sub chkVerActivos_Click()
On Error GoTo ErrorManager

        FunFPintaClasificacion
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub cmdActivar_Click()
 On Error GoTo ErrManager
    
    If Me.grdClasificacion.Row = 0 Or Me.grdClasificacion.RowHeight(Me.grdClasificacion.RowSel) = 0 Then
       MsgBox "Debe seleccionar la clasificación a Activar.", vbInformation, App.Title
        Exit Sub
    End If
    
    If Me.proClasificacion.Item(Me.grdClasificacion.Row).proRecordStatus = 1 Then
        MsgBox "Este registro ya se encuentra Activado.", vbInformation, App.Title
        Exit Sub
    End If
    
    If MsgBox("Desea Activar la clasificacion: " & Me.proClasificacion.Item(Me.grdClasificacion.Row).proClasificacion, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Me.proClasificacion.Item(Me.grdClasificacion.Row).proRecordStatus = 1
       If Me.proClasificacion.Item(Me.grdClasificacion.Row).FunGModificar = True Then
            MsgBox "La clasificación fue activada exitosamente.", vbInformation, App.Title
        Else
            MsgBox "Error al activar la clasificación.", vbCritical, App.Title
        End If
    End If
    
    Call FunFPintaClasificacion
    
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Private Sub cmdEliminar_Click()
On Error GoTo ErrorManager

    If Me.grdClasificacion.Row = 0 Or Me.grdClasificacion.RowHeight(Me.grdClasificacion.RowSel) = 0 Then
       MsgBox "Debe seleccionar una clasificacion a eliminar", vbInformation, App.Title
        Exit Sub
    End If
    
    If Me.proClasificacion.FunGEliminarClasificacion(Me.grdClasificacion.Row) = False Then
        MsgBox "No fue posible eliminar la clasificacion", vbInformation + vbOKOnly, App.Title
    End If
    Call FunFPintaClasificacion
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdModificar_Click()
On Error GoTo ErrorManager

    If Me.grdClasificacion.Row = 0 Or Me.grdClasificacion.RowHeight(Me.grdClasificacion.RowSel) = 0 Then
   
        MsgBox "Debe seleccionar una clasificación a modificar", vbInformation, App.Title
        Exit Sub
    End If
    
    Set frmEdicionClasificacion.proClasificacion = Me.proClasificacion.Item(Me.grdClasificacion.Row)
        
    frmEdicionClasificacion.Show vbModal
    Call FunFPintaClasificacion
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdNuevo_Click()
Dim varClasificacion As claClasificacion
On Error GoTo ErrorManager

    Set varClasificacion = New claClasificacion
    Set varClasificacion.proConexion = Me.proConexion
    
    Set frmEdicionClasificacion.proClasificacion = varClasificacion
     frmEdicionClasificacion.Show vbModal
    
    If Len(Trim(varClasificacion.proClasificacionId)) <> 0 Then
        Me.proClasificacion.Add Me.proConexion, varClasificacion.proRecordStatus, varClasificacion.proClasificacion, varClasificacion.proClasificacionId
    End If
    
    Call FunFPintaClasificacion
    
    Set varClasificacion = Nothing
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdReglas_Click()
   On Error GoTo ErrorManager

    If Me.grdClasificacion.Row > 0 Then
        Set frmReglasClasificacion.proConexion = Me.proConexion
        frmReglasClasificacion.proClasificacionId = Me.proClasificacion.Item(Me.grdClasificacion.Row).proClasificacionId
        frmReglasClasificacion.Show 1
    Else
        MsgBox "Debe seleccionar una clasificación "
        Exit Sub
    End If

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
    Set Me.proClasificacion = New colClasificacion
    Set Me.proClasificacion.proConexion = Me.proConexion
    
    If Me.proClasificacion.FunGConsulta = False Then
            MsgBox "No fue posible realizar la consulta de clasificacion", vbInformation + vbOKOnly, App.Title
            varBandera = 0
    End If
    
    'Muestra las clasificaciones
    FunFPintaClasificacion
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorManager

        Set proClasificacion = Nothing
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Sub SUbFPintaEliminados(parEliminado As Boolean, Optional parOcultar As Variant)
Dim varColumna As Integer
On Error GoTo ErrorManager

    Me.grdClasificacion.Row = Me.grdClasificacion.Rows - 1
    For varColumna = 0 To Me.grdClasificacion.Cols - 1
            Me.grdClasificacion.Col = varColumna
            Me.grdClasificacion.CellForeColor = &HC0C0C0
    Next varColumna
    
    If parEliminado Then
        If IsMissing(parOcultar) = False Then
            If parOcultar Then
            ' esconde la columna
                Me.grdClasificacion.RowHeight(Me.grdClasificacion.Rows - 1) = 0
            End If
        End If
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Private Sub grdClasificacion_DblClick()

On Error GoTo ErrManager
    Call cmdModificar_Click
       Exit Sub
ErrManager:
    SubGMuestraError
End Sub


