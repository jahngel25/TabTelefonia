VERSION 5.00
Begin VB.Form frmVozIncident 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incidente a relacionar"
   ClientHeight    =   2385
   ClientLeft      =   3780
   ClientTop       =   5010
   ClientWidth     =   7650
   Icon            =   "frmVozIncident.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7650
   Begin VB.Frame fraDatosIncidenteNuevo 
      BackColor       =   &H00C09258&
      Caption         =   "  Información del Incidente  "
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
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7785
   End
   Begin VB.Frame fraIncidenteNuevo 
      Height          =   1605
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   7665
      Begin VB.TextBox txtCodigoProducto 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1590
         TabIndex        =   15
         Top             =   1200
         Width           =   1635
      End
      Begin VB.TextBox txtNombreProducto 
         Enabled         =   0   'False
         Height          =   345
         Left            =   3270
         TabIndex        =   14
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txtCodigoEnlace 
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
         Height          =   345
         Left            =   1590
         TabIndex        =   5
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox txtProductId 
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
         Height          =   345
         Left            =   1590
         MaxLength       =   9
         TabIndex        =   4
         Top             =   480
         Width           =   1635
      End
      Begin VB.Frame fraComentariosNuevo 
         Caption         =   "Comentarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   3270
         TabIndex        =   2
         Top             =   0
         Width           =   4395
         Begin VB.TextBox txtComentariosNuevo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   90
            TabIndex        =   3
            Top             =   180
            Width           =   4245
         End
      End
      Begin VB.TextBox txtIdAsuntoNuevo 
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
         Left            =   1590
         MaxLength       =   9
         TabIndex        =   1
         Top             =   150
         Width           =   1635
      End
      Begin VB.Label lblProductMaster 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   780
         TabIndex        =   13
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label lblCodigoEnlace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Enlace:"
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
         Left            =   210
         TabIndex        =   8
         Top             =   870
         Width           =   1290
      End
      Begin VB.Label lblProductId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id del Producto:"
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
         Left            =   390
         TabIndex        =   7
         Top             =   510
         Width           =   1110
      End
      Begin VB.Label lblIdAsuntoNuevo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID del Asunto:"
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
         Left            =   510
         TabIndex        =   6
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame fraBotones 
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   1710
      Width           =   7665
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar Información"
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
         Height          =   315
         Left            =   5130
         TabIndex        =   12
         Top             =   210
         Width           =   1935
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   315
         Left            =   690
         TabIndex        =   11
         ToolTipText     =   "Guardar Datos de la Facturacion"
         Top             =   210
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmVozIncident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection
Public proDatosProducto As claDatosProducto
Public proOnyx As claONYX
Public proProceso As claProceso

Public proClienteTelefonia As claClienteTelefonia

Public proInsUpd As String

Private Sub cmdEditar_Click()
    Dim varProductMaster As EDCAdminVoz.colProductMaster
    Dim varDatosProductoIncident As EDCAdminVoz.claDatosProductoIncident
    Dim varcolRestriccionTabTel As colRestriccionTabTel
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    Set Me.proProceso = New claProceso
    Set Me.proProceso.proConexion = Me.proConexion
    
    If Me.txtIdAsuntoNuevo.Enabled Then
        Me.proProceso.proCompanyId = Me.proOnyx.ContactID
        Me.proProceso.proIncidentId = Trim(Me.txtIdAsuntoNuevo.Text)
        Me.proProceso.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
        Me.proProceso.proProductId = Me.proDatosProducto.proProductId
        
        If Not Me.proProceso.MetValidarAsunto() Then
            Me.txtIdAsuntoNuevo.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        'Validar incidentes no permitidos en este tab
        Set varcolRestriccionTabTel = New colRestriccionTabTel
        Set varcolRestriccionTabTel.proConexion = Me.proConexion
        If Not varcolRestriccionTabTel.FunGValidarRestriccionesTab(CLng(Me.txtIdAsuntoNuevo.Text), Me.proConexion) Then
            Me.txtIdAsuntoNuevo.SetFocus
            Exit Sub
        End If
        
    End If
    'Validar la informacion del producto digitado
    If Trim(Me.txtProductId) = "" Or Trim(Me.txtProductId) = "0" Then
        'Si es un producto nuevo debo buscar el producto en la venta
        Set varProductMaster = New EDCAdminVoz.colProductMaster
        Set varProductMaster.proConexion = Me.proConexion
        varProductMaster.proIncidentId = Trim(Me.txtIdAsuntoNuevo.Text)
        
        If varProductMaster.MetConsultarxIncidente Then
            Me.proDatosProducto.proProductNumber = varProductMaster.Item(1).proProductNumber
            Me.proDatosProducto.proProductName = varProductMaster.Item(1).proDescription
            Me.txtCodigoProducto.Text = varProductMaster.Item(1).proProductNumber
            Me.txtNombreProducto.Text = varProductMaster.Item(1).proDescription
            If Me.proDatosProducto.proParametrosProducto Is Nothing Then
                Set Me.proDatosProducto.proParametrosProducto = New EDCAdminVoz.colParametroProducto
                Set Me.proDatosProducto.proParametrosProducto.proConexion = Me.proConexion
            
                Me.proDatosProducto.proParametrosProducto.proProductNumber = Me.proDatosProducto.proProductNumber
                If Me.proDatosProducto.proParametrosProducto.metConsultarxProducto Then
                    If Me.proDatosProducto.proParametrosProducto.Count = 0 Then
                        MsgBox "El producto relacionado en el incidente, no tiene parámetros relacionados.", vbInformation, App.Title
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                Else
                    MsgBox "Error al consultar los parametros por producto.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
            
        Else
            MsgBox "Error al consultar el producto del Incidente.", vbCritical, App.Title
            Screen.MousePointer = 0
            Exit Sub
        End If
    Else
        Me.proProceso.proCompanyId = Me.proOnyx.ContactID
        Me.proProceso.proIncidentId = Me.proDatosProducto.proIncidentId
        Me.proProceso.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
        Me.proProceso.proProductId = Me.proDatosProducto.proProductId
        
        Me.txtCodigoEnlace.Text = Me.proProceso.MetValidarProducto()
        If Me.txtCodigoEnlace.Text = "" Then
            Me.txtProductId.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
     Set Me.proDatosProducto.proConexion = Me.proConexion
    Me.proDatosProducto.proComentarios = Me.txtComentariosNuevo
    Me.proDatosProducto.proRecordStatus = 1
    Set varDatosProductoIncident = Nothing
    Set varDatosProductoIncident = New EDCAdminVoz.claDatosProductoIncident
    Set varDatosProductoIncident.proConexion = Me.proConexion
    varDatosProductoIncident.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
    varDatosProductoIncident.proIncidentId = Trim(Me.txtIdAsuntoNuevo.Text)
    varDatosProductoIncident.proFechaModificacion = Format(Now, "mm/dd/yyyy HH:mm:ss")
    If Not Me.proDatosProducto.proDatosProductoIncident Is Nothing Then
        Me.proDatosProducto.proIncidentId = varDatosProductoIncident.proIncidentId
        If Not Me.proDatosProducto.MetVerificarExistenciaIncidente Then
        
            If Not Me.proDatosProducto.MetAgregarIncidente(varDatosProductoIncident) Then
                MsgBox "Error al agregar este incidente al TAB de Datos por Servicio.", vbCritical, App.Title
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    Else
        If Not Me.proDatosProducto.MetAgregarIncidente(varDatosProductoIncident) Then
            MsgBox "Error al agregar este incidente al TAB de Datos por Servicio.", vbCritical, App.Title
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    Me.proDatosProducto.proIncidentId = varDatosProductoIncident.proIncidentId
    If Me.proDatosProducto.proiVentaid = "" Then
        Me.proDatosProducto.proiVentaid = varDatosProductoIncident.proIncidentId
    End If
    Set frmDetalleDatosProducto.proDatosProducto = Me.proDatosProducto
    Set frmDetalleDatosProducto.proConexion = Me.proConexion
    Set frmDetalleDatosProducto.proOnyx = Me.proOnyx
    Set frmDetalleDatosProducto.proClienteTelefonia = Me.proClienteTelefonia
    Screen.MousePointer = 0
    
    frmDetalleDatosProducto.Show (1)
    
    
    Unload Me
    
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo ErrManager
    
    Me.proDatosProducto.proComentarios = Me.txtComentariosNuevo.Text
    
    If Me.proDatosProducto.MetGuardar Then
        MsgBox "La información se actualizó exitosamente.", vbInformation, App.Title
    Else
        MsgBox "Error al actualizar la información.", vbCritical, App.Title
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    Me.txtCodigoEnlace.Text = Me.proDatosProducto.proCodigoEnlace
    Me.txtProductId.Text = Me.proDatosProducto.proProductId
    
    If Val(Trim(Me.txtIdAsuntoNuevo.Text)) = 0 Then
        Me.cmdEditar.Enabled = False
    End If
    
    If Me.proInsUpd = "U" Then
        Me.txtIdAsuntoNuevo.Text = Me.proDatosProducto.proIncidentId
        Me.txtIdAsuntoNuevo.Enabled = False
        Me.txtCodigoProducto.Text = Me.proDatosProducto.proProductNumber
        Me.txtNombreProducto.Text = Me.proDatosProducto.proProductName
        Me.cmdGuardar.Visible = True
        Me.cmdGuardar.Enabled = True
        Me.txtComentariosNuevo.Text = Me.proDatosProducto.proComentarios
    End If
    
    If Val(Trim(Me.txtIdAsuntoNuevo)) = 0 Then
        If Mid$(Me.proOnyx.AlternateID, 1, 3) = "CRM" Then
            Me.txtIdAsuntoNuevo = Mid$(Me.proOnyx.AlternateID, 4, Len(Me.proOnyx.AlternateID) - 1)
        Else
            Me.txtIdAsuntoNuevo = ""
        End If
        Me.cmdEditar.Enabled = True
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrManager
    
    If Me.txtIdAsuntoNuevo.Enabled Then
        Me.txtIdAsuntoNuevo.SetFocus
        Me.cmdGuardar.Visible = False
    Else
        Me.txtComentariosNuevo.SetFocus
        Me.cmdGuardar.Visible = True
    End If
    Exit Sub
ErrManager:
    SubGMuestraError
 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorManager

    If KeyAscii = 13 Then
            If Me.cmdEditar.Enabled Then
                Me.cmdEditar.SetFocus
            End If
    End If
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub


Private Sub txtComentariosNuevo_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtComentariosNuevo.SelStart = 0
    Me.txtComentariosNuevo.SelLength = Len(Me.txtComentariosNuevo.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtIdAsuntoNuevo_Change()
    On Error GoTo ErrManager
    
        If Trim(Me.txtIdAsuntoNuevo) = "" Then
            Me.cmdEditar.Enabled = False
        Else
            Me.cmdEditar.Enabled = True
        End If
        
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtIdAsuntoNuevo_GotFocus()
    On Error GoTo ErrManager
    
    Me.txtIdAsuntoNuevo.SelStart = 0
    Me.txtIdAsuntoNuevo.SelLength = Len(Me.txtIdAsuntoNuevo.Text)
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub txtIdAsuntoNuevo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrManager
    
        KeyAscii = FunGLeeNumerico(KeyAscii)
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrManager
    frmVoz.Tag = "1"
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
