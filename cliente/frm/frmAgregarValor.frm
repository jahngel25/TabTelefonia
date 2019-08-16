VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAgregarValor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar Valores"
   ClientHeight    =   3825
   ClientLeft      =   5505
   ClientTop       =   5835
   ClientWidth     =   4215
   Icon            =   "frmAgregarValor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   2295
      Left            =   0
      TabIndex        =   8
      Top             =   1530
      Width           =   4245
      _Version        =   65536
      _ExtentX        =   7488
      _ExtentY        =   4048
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin MSFlexGridLib.MSFlexGrid grdValores 
         Height          =   2175
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   2730
      TabIndex        =   7
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   1200
      Width           =   1485
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   885
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   4245
      _Version        =   65536
      _ExtentX        =   7488
      _ExtentY        =   1561
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
      BorderWidth     =   2
      BevelInner      =   1
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   450
         Width           =   2355
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   510
         Width           =   885
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   630
         TabIndex        =   2
         Top             =   180
         Width           =   540
      End
   End
   Begin Threed.SSPanel pnlTituloValores 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   503
      _StockProps     =   15
      Caption         =   "Asignación de los valores"
      ForeColor       =   16777215
      BackColor       =   12620376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAgregarValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection
Public proOnyx As EDCVoz.claONYX
Public proNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
Public proParametroProducto As EDCAdminVoz.claParametroProducto

Private Sub cmdAgregar_Click()
    Dim varValor As EDCAdminVoz.claValor
    Dim varValorCampoProducto As EDCAdminVoz.claValoresCampoProducto
    On Error GoTo ErrManager
    
    If Trim(Me.txtDescripcion.Text) = "" Then
        MsgBox "Debe digitar la descripción del valor.", vbInformation, App.Title
        Exit Sub
    End If
    
    Set varValorCampoProducto = New EDCAdminVoz.claValoresCampoProducto
    Set varValorCampoProducto.proConexion = Me.proConexion
    varValorCampoProducto.proCampo = Me.proParametroProducto.proCampo
    varValorCampoProducto.proProductNumber = Me.proParametroProducto.proProductNumber
    
    Set varValor = New EDCAdminVoz.claValor
    Set varValor.proConexion = Me.proConexion
        
    varValor.proValorDesc = Trim(Me.txtDescripcion.Text)
    varValor.proRecordStatus = 1
    
    'Verificar que el valor no exista
    If varValor.MetConsultarxDescripcion Then
        'Si no existe lo inserta como valor y lo relaciona con el campo
       If varValor.proValorId = 0 Then
            If varValor.MetInsertar Then
                varValorCampoProducto.proValorId = varValor.proValorId
                varValorCampoProducto.proValorDesc = varValor.proValorDesc
                Me.txtCodigo.Text = varValor.proValorId
                
                If varValorCampoProducto.MetInsertar Then
                    If Me.proParametroProducto.MetAgregarValor(varValorCampoProducto) Then
                        Call SubFPintarGrid
                        MsgBox "El valor se agregó exitosamente.", vbInformation, App.Title
                        Me.txtCodigo.Text = ""
                        Me.txtDescripcion.Text = ""
                    Else
                        MsgBox "Error al agregar el valor.", vbCritical, App.Title
                    End If
                Else
                    MsgBox "Error al ligar el valor al campo.", vbCritical, App.Title
                End If
            Else
                MsgBox "Error al insertar el valor.", vbCritical, App.Title
            End If
        Else
        'Si existe debe verificar que aun no este en la relacion
            varValorCampoProducto.proValorId = varValor.proValorId
            varValorCampoProducto.proValorDesc = varValor.proValorDesc
            
            Me.txtCodigo.Text = varValor.proValorId
            If Not varValorCampoProducto.MetValidarExistencia Then
                If varValorCampoProducto.MetInsertar Then
                    If Me.proParametroProducto.MetAgregarValor(varValorCampoProducto) Then
                        Call SubFPintarGrid
                        MsgBox "El valor se agregó exitosamente.", vbInformation, App.Title
                        Me.txtCodigo.Text = ""
                        Me.txtDescripcion.Text = ""
                    Else
                        MsgBox "Error al agregar el valor.", vbCritical, App.Title
                    End If
                Else
                    MsgBox "Error al ligar el valor al campo.", vbCritical, App.Title
                End If
            Else
                MsgBox "El valor ya se encuentra relacionado en este campo.", vbInformation, App.Title
            End If
        End If
    Else
        MsgBox "Error al verificar la existencia del valor", vbCritical, App.Title
    End If
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub CmdCancelar_Click()
    On Error GoTo ErrManager
    
    Unload Me
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    Call SubFInicializarGrid
    
    Call SubFPintarGrid
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Sub SubFInicializarGrid()
    On Error GoTo ErrManager
    
    With Me.grdValores
        .Rows = 1
        .Cols = 4
        .Row = 0
        
        .Col = 0
        .ColWidth(0) = 0
        .CellAlignment = 4
        .TextMatrix(0, 0) = "Codigo Producto"
        
        .Col = 1
        .ColWidth(1) = 0
        .CellAlignment = 4
        .TextMatrix(0, 1) = "Campo"
        
        .Col = 2
        .ColWidth(2) = 1500
        .CellAlignment = 4
        .TextMatrix(0, 2) = "Codigo"
        
        .Col = 3
        .ColWidth(3) = 2000
        .CellAlignment = 4
        .TextMatrix(0, 3) = "Descripcion"
        
        .Row = 0
        .Col = 0
    End With
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Public Sub SubFPintarGrid()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.grdValores.Rows = 1
    
    For varContador = 1 To Me.proParametroProducto.proValores.Count
        Me.grdValores.AddItem Me.proParametroProducto.proValores.Item(varContador).proProductNumber & vbTab & _
                              Me.proParametroProducto.proValores.Item(varContador).proCampo & vbTab & _
                              Me.proParametroProducto.proValores.Item(varContador).proValorId & vbTab & _
                              Me.proParametroProducto.proValores.Item(varContador).proValorDesc
    Next varContador
    
    Me.grdValores.Row = 0
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub
