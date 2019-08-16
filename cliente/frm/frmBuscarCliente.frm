VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBuscarCliente 
   Caption         =   "Búsqueda de Cliente"
   ClientHeight    =   6315
   ClientLeft      =   540
   ClientTop       =   3555
   ClientWidth     =   14340
   Icon            =   "frmBuscarCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   14340
   Begin VB.Frame fraFinanciacion 
      Caption         =   "  Datos de Financiación  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14325
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
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
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   10
         Tag             =   "Es indispensable indicar el nombre del descuento"
         Top             =   330
         Width           =   1965
      End
      Begin VB.TextBox txtNombreCliente 
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
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   9
         Tag             =   "Es indispensable indicar el nombre del descuento"
         Top             =   660
         Width           =   6945
      End
      Begin VB.Frame fraTitulo 
         BackColor       =   &H00C09258&
         Caption         =   "  Datos de Búsqueda"
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
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   14355
      End
      Begin VB.CommandButton cmdBuscarCliente 
         Caption         =   "&Buscar Cliente"
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
         Left            =   7080
         TabIndex        =   7
         ToolTipText     =   "Refleja las cuotas del Tiempo de Financiación"
         Top             =   1380
         Width           =   1665
      End
      Begin VB.TextBox txtNIT 
         Alignment       =   1  'Right Justify
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
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   6
         Tag             =   "Es indispensable indicar el nombre del descuento"
         Top             =   990
         Width           =   1965
      End
      Begin MSMask.MaskEdBox txtEnlace 
         Height          =   345
         Left            =   1890
         TabIndex        =   5
         Top             =   1320
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "???AA##"
         PromptChar      =   " "
      End
      Begin VB.Label lblIgual 
         Caption         =   "(= IGUAL A)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D6980A&
         Height          =   195
         Index           =   3
         Left            =   3930
         TabIndex        =   18
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblIgual 
         Caption         =   "(LIKE)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D6980A&
         Height          =   315
         Index           =   2
         Left            =   3930
         TabIndex        =   17
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblIgual 
         Caption         =   "(LIKE)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D6980A&
         Height          =   315
         Index           =   1
         Left            =   8910
         TabIndex        =   16
         Top             =   690
         Width           =   435
      End
      Begin VB.Label lblIgual 
         Caption         =   "(= IGUAL A)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D6980A&
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   15
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lblId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID del Cliente"
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
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   330
         Width           =   1785
      End
      Begin VB.Label lblNombreCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre del Cliente"
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
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   660
         Width           =   1785
      End
      Begin VB.Label lblNIT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NIT"
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
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   990
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C09258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enlace"
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
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   1320
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Datos de Financiación  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   1740
      Width           =   14355
      Begin VB.CommandButton cmdSeleccionarCliente 
         Caption         =   "&Seleccionar Cliente"
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
         Left            =   12540
         TabIndex        =   3
         ToolTipText     =   "Refleja las cuotas del Tiempo de Financiación"
         Top             =   4230
         Width           =   1665
      End
      Begin VB.Frame fraTitulo 
         BackColor       =   &H00C09258&
         Caption         =   " Resultado de la Búsqueda"
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
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   30
         Width           =   14385
      End
      Begin MSFlexGridLib.MSFlexGrid grdClientes 
         Height          =   3915
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   6906
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
End
Attribute VB_Name = "frmBuscarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public proCliente As claCliente

Dim varColCliente As colCliente

Function FunFPintaClientes() As Boolean
Dim varContador As Integer
Dim varCliente As claCliente

On Error GoTo ErrorManager

    'Adecua la grilla
    Me.grdClientes.Rows = 2
    Me.grdClientes.Cols = 7
    Me.grdClientes.TextMatrix(0, 0) = "ID"
    Me.grdClientes.TextMatrix(0, 1) = "Cliente"
    Me.grdClientes.TextMatrix(0, 2) = "NIT"
    Me.grdClientes.TextMatrix(0, 3) = "Estado"
    Me.grdClientes.TextMatrix(0, 4) = "Ciudad"
    Me.grdClientes.TextMatrix(0, 5) = "Direccion"
    Me.grdClientes.TextMatrix(0, 6) = "Sede"
    Me.grdClientes.FixedRows = 1
    Me.grdClientes.ColWidth(0) = 1080
    Me.grdClientes.ColWidth(1) = 3525
    Me.grdClientes.ColWidth(2) = 1635
    Me.grdClientes.ColWidth(3) = 1635
    Me.grdClientes.ColWidth(4) = 1635
    Me.grdClientes.ColWidth(5) = 1635
    Me.grdClientes.ColWidth(6) = 2000
    Me.grdClientes.Rows = 1
    
    For varContador = 1 To varColCliente.Count
        Set varCliente = Nothing
        Set varCliente = New claCliente
        
        varCliente.proClienteId = varColCliente.Item(varContador).proClienteId
        varCliente.proNombreCliente = varColCliente.Item(varContador).proNombreCliente
        varCliente.proNIT = varColCliente.Item(varContador).proNIT
        varCliente.proEstadoCliente = varColCliente.Item(varContador).proEstadoCliente
        varCliente.proCiudad = varColCliente.Item(varContador).proCiudad
        varCliente.proDireccion = varColCliente.Item(varContador).proDireccion
        varCliente.proSede = varColCliente.Item(varContador).proSede
        
        Me.grdClientes.AddItem varCliente.proClienteId & vbTab & _
                                            varCliente.proNombreCliente & vbTab & _
                                            varCliente.proNIT & vbTab & _
                                            varCliente.proEstadoCliente & vbTab & _
                                            varCliente.proCiudad & vbTab & _
                                            varCliente.proDireccion & vbTab & _
                                            varCliente.proSede
    Next varContador
    
    Me.grdClientes.Row = 0
    
    Set varCliente = Nothing
    FunFPintaClientes = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Private Sub cmdBuscarCliente_Click()
On Error GoTo ErrorManager

    Set varColCliente = New colCliente
    Set varColCliente.proConexion = Me.proCliente.proConexion
    
    'Valida que exista algúna criterio de búsqueda
    If Len(Trim(Me.txtID)) = 0 And Len(Trim(Me.txtNombreCliente)) = 0 And Len(Trim(Me.txtNIT)) = 0 And Len(Trim(Me.txtEnlace)) = 0 Then
        MsgBox "Debe indicar algún criterio de búsqueda", vbInformation, App.Title
        Exit Sub
    End If
    
    If Len(Trim(Me.txtID)) Then
                varColCliente.proClienteId = Trim(Me.txtID)
                If varColCliente.funGConsultaClientexID = False Then
                        MsgBox "No fue posible realizar la consulta por ID del Cliente", vbInformation, App.Title
                        Exit Sub
                 End If
    End If
    If Len(Trim(Me.txtNombreCliente)) Then
                varColCliente.proNombreCliente = Trim(Me.txtNombreCliente)
                If varColCliente.FunGConsultaClientexNombre = False Then
                        MsgBox "No fue posible realizar la consulta por Nombre del Cliente", vbInformation, App.Title
                        Exit Sub
                 End If
    End If
    If Len(Trim(Me.txtNIT)) Then
                varColCliente.proNIT = Trim(Me.txtNIT)
                If varColCliente.funGConsultaClientexNIT = False Then
                        MsgBox "No fue posible realizar la consulta por NIT del Cliente", vbInformation, App.Title
                        Exit Sub
                 End If
    End If
    If Len(Trim(Me.txtEnlace)) Then
                varColCliente.proEnlace = Trim(Me.txtEnlace)
                If varColCliente.funGConsultaClientexEnlace = False Then
                        MsgBox "No fue posible realizar la consulta por Enlace del Cliente", vbInformation, App.Title
                        Exit Sub
                 End If
    End If
    
    If varColCliente.Count = 0 Then
            MsgBox "No se han encontrado resultados por el criterio de búsqueda", vbInformation, App.Title
    End If
    
    'Despliegue de los clientes encontrados
    Call FunFPintaClientes
    
    Me.cmdSeleccionarCliente.SetFocus
    Exit Sub
    
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdSeleccionarCliente_Click()
On Error GoTo ErrorManager

    If Me.grdClientes.Rows = 2 And Me.grdClientes.Row = 0 Then
            Me.grdClientes.Row = 1
    End If
    
    If Me.grdClientes.Row = 0 Then
            MsgBox "Debe seleccionar un cliente en la grilla de resultados", vbInformation, App.Title
            Exit Sub
    End If
    
    Me.proCliente.proClienteId = varColCliente(Me.grdClientes.Row).proClienteId
    Me.proCliente.proNombreCliente = varColCliente(Me.grdClientes.Row).proNombreCliente
    Me.proCliente.proNIT = varColCliente(Me.grdClientes.Row).proNIT
    Me.proCliente.proAntVenc = varColCliente(Me.grdClientes.Row).proAntVenc
        
    Unload Me
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub



Private Sub Form_Load()
On Error GoTo ErrorManager

        Set varColCliente = New colCliente
        Set varColCliente.proConexion = Me.proCliente.proConexion
        
        'Prepara la grilla
        Call FunFPintaClientes
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub


Private Sub grdClientes_DblClick()
On Error GoTo ErrorManager

        Call cmdSeleccionarCliente_Click
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub



Private Sub txtENlace_GotFocus()
On Error GoTo ErrorManager

        Me.txtEnlace.SelStart = 0
        Me.txtEnlace.SelLength = Len(Trim(Me.txtEnlace))
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtENlace_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorManager

        If KeyAscii = 13 Then
                Me.cmdBuscarCliente.SetFocus
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii Then
                Me.txtID = ""
                Me.txtNIT = ""
                Me.txtNombreCliente = ""
        End If
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtID_GotFocus()
On Error GoTo ErrorManager

        Me.txtID.SelStart = 0
        Me.txtID.SelLength = Len(Trim(Me.txtID))
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
Dim varMascara As String
On Error GoTo ErrorManager

        If KeyAscii = 13 Then
                Me.cmdBuscarCliente.SetFocus
        End If
        KeyAscii = FunGLeeNumerico(KeyAscii)
        If KeyAscii Then
                Me.txtNIT = ""
                Me.txtNombreCliente = ""
                varMascara = Me.txtEnlace.Mask
                Me.txtEnlace.Mask = ""
                Me.txtEnlace = ""
                Me.txtEnlace.Mask = varMascara
        End If
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtNIT_GotFocus()
On Error GoTo ErrorManager

        Me.txtNIT.SelStart = 0
        Me.txtNIT.SelLength = Len(Trim(Me.txtNIT))
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtNIT_KeyPress(KeyAscii As Integer)
Dim varMascara As String
On Error GoTo ErrorManager

        If KeyAscii = 13 Then
                Me.cmdBuscarCliente.SetFocus
        End If
        KeyAscii = FunGLeeNumerico(KeyAscii)
        If KeyAscii Then
                Me.txtID = ""
                Me.txtNombreCliente = ""
                varMascara = Me.txtEnlace.Mask
                Me.txtEnlace.Mask = ""
                Me.txtEnlace = ""
                Me.txtEnlace.Mask = varMascara
        End If
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtNombreCliente_GotFocus()

On Error GoTo ErrorManager

        Me.txtNombreCliente.SelStart = 0
        Me.txtNombreCliente.SelLength = Len(Trim(Me.txtNombreCliente))
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub txtNombreCliente_KeyPress(KeyAscii As Integer)
Dim varMascara As String
On Error GoTo ErrorManager

        If KeyAscii = 13 Then
                Me.cmdBuscarCliente.SetFocus
        End If
        KeyAscii = FunGLeeAlfaNumerico(KeyAscii)
        If KeyAscii Then
                Me.txtNIT = ""
                Me.txtID = ""
                varMascara = Me.txtEnlace.Mask
                Me.txtEnlace.Mask = ""
                Me.txtEnlace = ""
                Me.txtEnlace.Mask = varMascara
        End If
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub


