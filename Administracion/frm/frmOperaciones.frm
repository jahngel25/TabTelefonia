VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOperaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones por novedad"
   ClientHeight    =   6045
   ClientLeft      =   1320
   ClientTop       =   1950
   ClientWidth     =   7905
   Icon            =   "frmOperaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRangos 
      BackColor       =   &H00C09258&
      Caption         =   "Rangos (Minutos x cantidad de lineas)"
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
      Height          =   6000
      Left            =   0
      TabIndex        =   3
      Top             =   -45
      Width           =   7845
      Begin VB.Frame Frame2 
         BackColor       =   &H00C09258&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   5580
         Width           =   7080
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
            Left            =   90
            TabIndex        =   1
            ToolTipText     =   "Nuevo Tramo"
            Top             =   45
            Width           =   1185
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
            Left            =   1305
            TabIndex        =   2
            ToolTipText     =   "EliminarTramo"
            Top             =   45
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdOperaciones 
         Height          =   5265
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   9287
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Propiedad de conexion
Public proConexion As ADODB.Connection

'Colección de OperacionOnyx
Public proOperacionOnyx As colOperacionOnyx
Public proclaOperacionOnyx As claOperacionOnyx
'Bandera de Inicio de Ventana
Dim varBandera As Integer
Dim varFUsuariosONYX As colUsuario
Function FunFPintaOperaciones() As Boolean
Dim varContador As Integer
On Error GoTo ErrorManager
    grdOperaciones.Clear
    grdOperaciones.Rows = 1
    grdOperaciones.Cols = 8
    grdOperaciones.AllowUserResizing = flexResizeColumns
    grdOperaciones.TextMatrix(0, 0) = "CategoriaID"
    grdOperaciones.TextMatrix(0, 1) = "Categoria"
    grdOperaciones.TextMatrix(0, 2) = "TipoID"
    grdOperaciones.TextMatrix(0, 3) = "Tipo"
    grdOperaciones.TextMatrix(0, 4) = "OperacionID"
    grdOperaciones.TextMatrix(0, 5) = "Operacion"
    grdOperaciones.TextMatrix(0, 6) = "SeccionID"
    grdOperaciones.TextMatrix(0, 7) = "Seccion"
    grdOperaciones.ColWidth(0) = 0
    grdOperaciones.ColWidth(1) = 1900
    grdOperaciones.ColWidth(2) = 0
    grdOperaciones.ColWidth(3) = 2600
    grdOperaciones.ColWidth(4) = 0
    grdOperaciones.ColWidth(5) = 1900
    grdOperaciones.ColWidth(6) = 0
    grdOperaciones.ColWidth(7) = 1900
    If grdOperaciones.Rows > 1 Then grdOperaciones.FixedRows = 1
    For varContador = 1 To Me.proOperacionOnyx.Count
            grdOperaciones.AddItem Me.proOperacionOnyx.Item(varContador).proIncidentCategory & vbTab & _
            Me.proOperacionOnyx.Item(varContador).proNombreIncidente & vbTab & _
            Me.proOperacionOnyx.Item(varContador).proIncidentTypeId & vbTab & _
            Me.proOperacionOnyx.Item(varContador).proNombreTipoIncidente & vbTab & _
            Me.proOperacionOnyx.Item(varContador).proTipoNovedadId & vbTab & _
            Me.proOperacionOnyx.Item(varContador).proNombreTipoNovedad & vbTab & _
            Me.proOperacionOnyx.Item(varContador).proTipoSeccionId & vbTab & _
            Me.proOperacionOnyx.Item(varContador).proNombreSeccion
    Next varContador
    grdOperaciones.CellAlignment = flexAlignGeneral
    FunFPintaOperaciones = True
    Exit Function
    
ErrorManager:
    SubGMuestraError

End Function

Private Sub cmdEliminar_Click()

On Error GoTo ErrorManager

    If Me.grdOperaciones.RowSel <= 0 Then
        MsgBox "Debe seleccionar una registro a eliminar", vbInformation, App.Title
        Exit Sub
    End If
    
        'ELimina el elemento en la base de datos
        proclaOperacionOnyx.proIncidentTypeId = Me.proOperacionOnyx.Item(grdOperaciones.Row).proIncidentTypeId
        proclaOperacionOnyx.proTipoNovedadId = Me.proOperacionOnyx.Item(grdOperaciones.Row).proTipoNovedadId
        proclaOperacionOnyx.proIncidentCategory = Me.proOperacionOnyx.Item(grdOperaciones.Row).proIncidentCategory
        If MsgBox("Esta seguro de eliminar la operación con ID " & proOperacionOnyx.Item(Me.grdOperaciones.Row).proNombreTipoIncidente & "?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            Exit Sub
        End If
        If proOperacionOnyx.Item(Me.grdOperaciones.Row).MetEliminar = False Then
                MsgBox "No fue posible eliminar el registro de la base de datos", vbInformation, App.Title
                Exit Sub
        End If
    'ELimina el elemento de la colección
    If proOperacionOnyx.FunGEliminar(grdOperaciones.Row) = False Then
        MsgBox "No fue posible eliminar el registro de la colección", vbInformation + vbOKOnly, App.Title
    End If
    
    Call FunFPintaOperaciones
    Exit Sub
    
ErrorManager:
    SubGMuestraError



End Sub

Private Sub cmdNuevo_Click()

On Error GoTo ErrorManager
Dim varproclaOperacionOnyx As claOperacionOnyx
Set varproclaOperacionOnyx = New claOperacionOnyx

Set varproclaOperacionOnyx.proConexion = Me.proConexion
Set frmEditarOperaciones.proOperacionOnyx = varproclaOperacionOnyx
Set frmEditarOperaciones.proConexion = Me.proConexion

    'Despliegue de ventana de edición
    frmEditarOperaciones.Show vbModal
    
    If Len(Trim(varproclaOperacionOnyx.proIncidentCategory)) <> 0 Then
        'Agrega la recién creada a la colección
        Me.proOperacionOnyx.Add Me.proConexion, varproclaOperacionOnyx.proIncidentCategory, varproclaOperacionOnyx.proTipoNovedadId, varproclaOperacionOnyx.proIncidentTypeId, varproclaOperacionOnyx.proTipoSeccionId
    End If
    
    'Desapliega la colección en la grilla
    Call FunFPintaOperaciones
   
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    
    Set proclaOperacionOnyx = New claOperacionOnyx
    Set proclaOperacionOnyx.proConexion = Me.proConexion
    Set proOperacionOnyx = New colOperacionOnyx
    Set proOperacionOnyx.proConexion = Me.proConexion
    proOperacionOnyx.MetConsultarTodos
    Call FunFPintaOperaciones
End Sub

