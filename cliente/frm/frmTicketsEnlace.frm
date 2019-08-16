VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTicketsEnlace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Peticiones vigentes sobre el enlace"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   45
      TabIndex        =   2
      Top             =   3915
      Width           =   5160
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Peticiones con vinculación"
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
         Left            =   2790
         TabIndex        =   7
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label lblSeleccionModificacion 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   165
         Left            =   2595
         TabIndex        =   6
         Top             =   240
         Width           =   165
      End
      Begin VB.Label lblInsertar 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   165
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Peticiones sin vinculaciones"
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
         Left            =   420
         TabIndex        =   4
         Top             =   180
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdSeleccionarTodosModificacion 
      Caption         =   "Listado"
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
      Left            =   5265
      TabIndex        =   1
      Top             =   4095
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdDeseleccionarTodosModificacion 
      Caption         =   "Vinculación"
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
      Left            =   7020
      TabIndex        =   3
      Top             =   4095
      Width           =   1785
   End
   Begin MSFlexGridLib.MSFlexGrid grdTickets 
      Height          =   3840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   6773
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmTicketsEnlace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Propiedad de conexion
Public proConexion As ADODB.Connection
Public proOnyx As EDCVoz.claONYX
Public proDatosProductoId As String
Public proDatosProducto As claDatosProducto
'Colección de TicketEnlace
Public proIncidentId As Long
Public provchSerialNumber As String
Public proTicketEnlace As colticketEnlace
Public proclaTicketEnlace As claTicketEnlace
Dim varFUsuariosONYX As EDCAdminVoz.colUsuario
Function FunFPintaOperaciones() As Boolean
Dim varContador As Integer
Dim i As Integer
Dim iEstadoAtencionCerrado As Long
iEstadoAtencionCerrado = 104
On Error GoTo ErrorManager
    grdTickets.Clear
    grdTickets.Rows = 1
    grdTickets.Cols = 6
    grdTickets.FixedCols = 0
    grdTickets.AllowUserResizing = flexResizeColumns
    grdTickets.TextMatrix(0, 0) = "Id"
    grdTickets.TextMatrix(0, 1) = "Descripción"
    grdTickets.TextMatrix(0, 2) = "Asignado a"
    grdTickets.TextMatrix(0, 3) = "Tipo"
    grdTickets.TextMatrix(0, 4) = "Fecha"
    grdTickets.TextMatrix(0, 5) = "Estado"
    grdTickets.ColWidth(0) = 700
    grdTickets.ColWidth(1) = 1900
    grdTickets.ColWidth(2) = 1000
    grdTickets.ColWidth(3) = 1900
    grdTickets.ColWidth(4) = 1900
    grdTickets.ColWidth(5) = 1900

    For varContador = 1 To Me.proTicketEnlace.Count
        grdTickets.AddItem Me.proTicketEnlace.Item(varContador).proiIncidentId & vbTab & _
        Me.proTicketEnlace.Item(varContador).provchDesc1 & vbTab & _
        Me.proTicketEnlace.Item(varContador).prochAssignedTo & vbTab & _
        Me.proTicketEnlace.Item(varContador).provchParameterDesc & vbTab & _
        Me.proTicketEnlace.Item(varContador).prodtInsertDate & vbTab & _
        Me.proTicketEnlace.Item(varContador).prosEstado
    Next varContador
    
    For i = 1 To grdTickets.Rows - 1
        If Me.proclaTicketEnlace.proTieneAsociaciones(Me.proTicketEnlace.Item(i).proiIncidentId) Then
            Call SubFPintarFila(Me.grdTickets, i, Me.lblSeleccionModificacion.BackColor)
        Else
            Call SubFPintarFila(Me.grdTickets, i, Me.lblInsertar.BackColor)
        End If
    Next
    grdTickets.CellAlignment = flexAlignLeftTop
    FunFPintaOperaciones = True
    Exit Function
    
ErrorManager:
    SubGMuestraError

End Function

Private Sub cmdDeseleccionarTodosModificacion_Click()
Dim iEstadoAtencionCerrado As String
    iEstadoAtencionCerrado = 104
    If grdTickets.RowSel <= 0 Then
        MsgBox "Debe seleccionar un registro"
        Exit Sub
    End If
    Set frmAsociacionTicketServicios.proConexion = Me.proConexion
    Set frmAsociacionTicketServicios.proOnyx = Me.proOnyx
    Set frmAsociacionTicketServicios.proDatosProducto = Me.proDatosProducto
    frmAsociacionTicketServicios.provchSerialNumber = Me.provchSerialNumber
    frmAsociacionTicketServicios.proIncidenteId = Me.grdTickets.TextMatrix(grdTickets.Row, 0)
    If Me.proTicketEnlace.Item(grdTickets.RowSel).proiStatusId = iEstadoAtencionCerrado Then
        MsgBox "El incidente está cerrado, No podrá asociar ni guardar sus modificaciones"
        frmAsociacionTicketServicios.btnGuardar.Enabled = False
    End If
    frmAsociacionTicketServicios.Show vbModal
    Call FunFPintaOperaciones
End Sub

Private Sub cmdSeleccionarTodosModificacion_Click()
   Load frmReporte
   
   With frmReporte
      .InitForm Me.grdTickets, Me.Caption
      .Show vbModal
   End With
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorManager
    Set proclaTicketEnlace = New claTicketEnlace
    proclaTicketEnlace.proiIncidentId = Me.proIncidentId
    Set proclaTicketEnlace.proConexion = Me.proConexion
    Set proTicketEnlace = New colticketEnlace
    proTicketEnlace.proiIncidentId = Me.proIncidentId
    proTicketEnlace.provchSerialNumber = Me.provchSerialNumber
    
    Set proTicketEnlace.proConexion = Me.proConexion
    If Not proTicketEnlace.FunGConsulta Then
        MsgBox "No fue posible consultar los tickets "
    End If
    Call FunFPintaOperaciones
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub
Private Sub grdTickets_SelChange()
On Error GoTo ErrorManager
If grdTickets.Row > 0 Then
    grdTickets.RowSel = grdTickets.Row
End If
    Exit Sub
ErrorManager:
        SubGMuestraError
End Sub
