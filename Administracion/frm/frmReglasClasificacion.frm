VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmReglasClasificacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reglas para una clasificación"
   ClientHeight    =   2940
   ClientLeft      =   2100
   ClientTop       =   3135
   ClientWidth     =   7410
   Icon            =   "frmReglasClasificacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFondo 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7395
      Begin VB.Frame fraTitulo 
         BackColor       =   &H00C09258&
         Caption         =   "Reglas elegidas"
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
         Index           =   3
         Left            =   3660
         TabIndex        =   6
         Top             =   0
         Width           =   3735
      End
      Begin VB.Frame fraTitulo 
         BackColor       =   &H00C09258&
         Caption         =   "Reglas"
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
         TabIndex        =   5
         Top             =   0
         Width           =   3675
      End
      Begin VB.CommandButton cmdIncluir 
         BackColor       =   &H00FF8080&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3540
         TabIndex        =   1
         Top             =   1260
         Width           =   345
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H00FF8080&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3540
         TabIndex        =   3
         Top             =   1560
         Width           =   345
      End
      Begin VB.Frame fraDescuento 
         Caption         =   "  Usuarios del Descuento  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Index           =   0
         Left            =   3690
         TabIndex        =   8
         Top             =   30
         Width           =   3705
         Begin MSFlexGridLib.MSFlexGrid grdReglasSeleccionadas 
            Height          =   2475
            Left            =   180
            TabIndex        =   2
            Top             =   300
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   4366
            _Version        =   393216
            FixedCols       =   0
            GridLines       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
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
      Begin VB.Frame fraDescuento 
         Caption         =   " Usuarios de ONYX "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   3675
         Begin MSFlexGridLib.MSFlexGrid grdReglasTotales 
            Height          =   2505
            Left            =   60
            TabIndex        =   0
            Top             =   270
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   4419
            _Version        =   393216
            FixedCols       =   0
            GridLines       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
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
End
Attribute VB_Name = "frmReglasClasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmReglasClasificacionFormulario
' Fecha  : 28/09/2004 10:13
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Administración de la relación m a n entre Clasificación y Reglas
'---------------------------------------------------------------------------------------


Option Explicit
'Conexion
Public proConexion As ADODB.Connection

'ID del Clasificacion
Public proClasificacionId As String

'Clasificacion sobre el cual se van a agregar las Reglas
Public procolClasificacion As colReglasClasificacion
Public proclaClasificacion As claReglasClasificacion

'Coleccion de Reglas
Dim procolReglasTodas As colReglasClasificacion

Dim varFBandera As Integer

Sub SubFExcluyeReglas(procolReglasTodas As colReglasClasificacion, procolClasificacion As colReglasClasificacion)

Dim varCuenta As Integer
Dim varCuentaSel As Integer
Dim varEncontro As Boolean
On Error GoTo ErrorManager
    For varCuenta = 1 To procolClasificacion.Count
        For varCuentaSel = 1 To procolReglasTodas.Count
                If Trim(procolClasificacion.Item(varCuenta).proiReglaId) = Trim(procolReglasTodas.Item(varCuentaSel).proiReglaId) Then
                    grdReglasTotales.RowHeight(varCuentaSel) = 0
                End If
        Next
    Next
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo ErrorManager
    Screen.MousePointer = 11
    If Me.grdReglasSeleccionadas.Row = 0 Then
        MsgBox "Debe seleccionar una Regla a excluir", vbInformation + vbOKOnly, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    Me.proclaClasificacion.proiReglaId = Me.procolClasificacion.Item(Me.grdReglasSeleccionadas.Row).proiReglaId
    Me.proclaClasificacion.proiClasificacionId = Me.proClasificacionId
    
    If Me.proclaClasificacion.FunGEliminar() = False Then
            MsgBox "No fue posible eliminar la Regla ", vbInformation + vbOKOnly, App.Title
            Exit Sub
    End If
   
    Call RefrescarListas
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub
Private Sub RefrescarListas()
        
   On Error GoTo ErrorManager

    If procolClasificacion.FunGConsultaSeleccionados = False Then
        MsgBox "No fue posible realizar la consulta de Reglas seleccionadas", vbInformation + vbOKOnly, App.Title
        varFBandera = 1
        Exit Sub
    End If
    
    If procolReglasTodas.FunGConsultaTodas = False Then
        MsgBox "No fue posible realizar la consulta de Reglas totales", vbInformation + vbOKOnly, App.Title
        varFBandera = 1
        Exit Sub
    End If
            
    'Despliega los usuarios en la grilla
    subFPintaReglas Me.grdReglasTotales, procolReglasTodas
    subFPintaReglas Me.grdReglasSeleccionadas, procolClasificacion
    Call SubFExcluyeReglas(procolReglasTodas, procolClasificacion)

      Exit Sub
ErrorManager:
    SubGMuestraError
    
End Sub
Private Sub cmdIncluir_Click()
On Error GoTo ErrorManager
    Screen.MousePointer = 11
    If Me.grdReglasTotales.Row = 0 Then
        MsgBox "Debe seleccionar una Regla a ser incluida dentro del Clasificacion", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    proclaClasificacion.proiClasificacionId = Me.proClasificacionId
    proclaClasificacion.proiReglaId = procolReglasTodas.Item(Me.grdReglasTotales.Row).proiReglaId
    'Agrega la Regla a la colección de Reglas
    If Me.proclaClasificacion.FunGInsertar() = False Then
            MsgBox "No fue posible agregar la Regla a la Clasificacion", vbInformation + vbOKOnly, App.Title
            Exit Sub
    End If
    Call RefrescarListas
    Screen.MousePointer = 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorManager

    If varFBandera Then
        varFBandera = 0
        Unload Me
    End If
    Call RefrescarListas
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Sub subFPintaReglas(parGrid As MSFlexGrid, parColReglas As colReglasClasificacion, Optional parCadena As Variant)

Dim varCuenta As Integer
Dim varTamaño As Integer
On Error GoTo ErrorManager
    
    parGrid.Redraw = False
    
    'Adecuación de la grilla
    parGrid.Rows = 2
    parGrid.Cols = 2
    parGrid.TextMatrix(0, 0) = "ID"
    parGrid.TextMatrix(0, 1) = "Regla"
    parGrid.FixedRows = 1
    parGrid.ColWidth(0) = 420
    parGrid.ColWidth(1) = 2595
    parGrid.ColAlignment(1) = 1
    parGrid.Rows = 1
    
    'Busca el tamaño de la cadena
    If IsMissing(parCadena) = False Then
        varTamaño = Len(Trim(parCadena))
    End If

    'Recorre la coleccion para llenar la grilla
    For varCuenta = 1 To parColReglas.Count
    
            parGrid.AddItem parColReglas.Item(varCuenta).proiReglaId & vbTab & _
                            parColReglas.Item(varCuenta).ProDescripcionRegla
                            
    Next varCuenta

    parGrid.Redraw = True
    parGrid.Refresh
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager

    Set proclaClasificacion = New claReglasClasificacion
    Set proclaClasificacion.proConexion = Me.proConexion
    
    Set procolClasificacion = New colReglasClasificacion
    Set procolClasificacion.proConexion = Me.proConexion
    
    Set procolReglasTodas = New colReglasClasificacion
    Set procolReglasTodas.proConexion = Me.proConexion
    
    procolClasificacion.proiClasificacionId = Me.proClasificacionId
    procolClasificacion.FunGConsultaSeleccionados
    procolReglasTodas.proiClasificacionId = Me.proClasificacionId
    

    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorManager

    varFBandera = 0
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub grdReglasSeleccionadas_DblClick()
On Error GoTo ErrorManager

    cmdExcluir_Click
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

Private Sub grdReglasTotales_DblClick()
On Error GoTo ErrorManager

    cmdIncluir_Click
    Exit Sub
    
ErrorManager:
    SubGMuestraError
End Sub

