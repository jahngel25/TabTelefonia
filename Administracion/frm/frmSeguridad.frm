VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSeguridad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administradores de Telefonía"
   ClientHeight    =   7770
   ClientLeft      =   2505
   ClientTop       =   3150
   ClientWidth     =   8625
   Icon            =   "frmSeguridad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   7800
      TabIndex        =   8
      Top             =   120
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   1191
      _StockProps     =   15
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
         Left            =   90
         Picture         =   "frmSeguridad.frx":0CCA
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.Frame fraTitulo 
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   8625
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A7811D&
         Caption         =   "Usuarios y Privilegios                                         "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   8535
      End
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar usuario..."
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
      Left            =   0
      TabIndex        =   5
      Top             =   7500
      Width           =   1575
   End
   Begin VB.CommandButton cmdELiminar 
      Caption         =   "&Eliminar usuario"
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
      Left            =   6690
      TabIndex        =   4
      Top             =   7500
      Width           =   1485
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar permisos..."
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
      Left            =   1590
      TabIndex        =   3
      Top             =   7500
      Width           =   1875
   End
   Begin VB.Frame fraFondo 
      Height          =   7155
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   8625
      Begin MSFlexGridLib.MSFlexGrid grdUsuarios 
         Height          =   6615
         Left            =   90
         TabIndex        =   6
         Top             =   420
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   11668
         _Version        =   393216
         GridLinesFixed  =   1
         BorderStyle     =   0
      End
      Begin MSFlexGridLib.MSFlexGrid grdTituloUsuarios 
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   450
         _Version        =   393216
         GridLines       =   0
         GridLinesFixed  =   1
         BorderStyle     =   0
      End
      Begin VB.Shape Shape1 
         Height          =   6885
         Left            =   60
         Top             =   180
         Width           =   8085
      End
   End
End
Attribute VB_Name = "frmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public proConexion As ADODB.Connection
Dim varColSeguridad As colSeguridad

Dim varBandera As Integer

Function FunFPintaUsuarios() As Boolean
Dim varContador As Integer
Dim varSeguridad As claSeguridad
Dim varCuenta As Integer

'Variables para pintar los permisos
Dim varPermisoAdministracion As String
Dim varPermisoAdministracionTelefonia As String
Dim varPermisoNoValidarProceso As String
Dim varPermisoNoValidarPorcentaje As String
On Error GoTo ErrorManager

    'Adecua la grilla
    Me.grdTituloUsuarios.Rows = 3
    Me.grdTituloUsuarios.Cols = 4
    Me.grdTituloUsuarios.TextMatrix(0, 0) = "ID"
    Me.grdTituloUsuarios.TextMatrix(0, 1) = "Administrar"
    Me.grdTituloUsuarios.TextMatrix(0, 2) = "Administrar"
    Me.grdTituloUsuarios.TextMatrix(0, 3) = "No Validar"
    Me.grdTituloUsuarios.FixedRows = 1
    Me.grdTituloUsuarios.FixedCols = 1
    Me.grdTituloUsuarios.ColWidth(0) = 1300
    Me.grdTituloUsuarios.ColWidth(1) = 1200
    Me.grdTituloUsuarios.ColWidth(2) = 1200
    Me.grdTituloUsuarios.ColWidth(3) = 1200
    Me.grdTituloUsuarios.Col = 0
    Me.grdTituloUsuarios.Row = 0
    Me.grdTituloUsuarios.CellAlignment = 4
    Me.grdTituloUsuarios.Col = 1
    Me.grdTituloUsuarios.CellAlignment = 4
    Me.grdTituloUsuarios.Col = 2
    Me.grdTituloUsuarios.CellAlignment = 4
    Me.grdTituloUsuarios.Col = 3
    Me.grdTituloUsuarios.CellAlignment = 4
    Me.grdTituloUsuarios.Rows = 1
    
    Me.grdUsuarios.Redraw = False
    Me.grdUsuarios.Rows = 2
    Me.grdUsuarios.Cols = 4
    Me.grdUsuarios.TextMatrix(0, 0) = "ONYX"
    Me.grdUsuarios.TextMatrix(0, 1) = "Sistema"
    Me.grdUsuarios.TextMatrix(0, 2) = "Telefonía"
    Me.grdUsuarios.TextMatrix(0, 3) = "Proceso ONYX"
    Me.grdUsuarios.FixedRows = 1
    Me.grdUsuarios.FixedCols = 1
    Me.grdUsuarios.ColWidth(0) = 1300
    Me.grdUsuarios.ColWidth(1) = 1200
    Me.grdUsuarios.ColWidth(2) = 1200
    Me.grdUsuarios.ColWidth(3) = 1200
    Me.grdUsuarios.Col = 0
    Me.grdUsuarios.Row = 0
    Me.grdUsuarios.CellAlignment = 4
    Me.grdUsuarios.ColAlignment(1) = 4
    Me.grdUsuarios.ColAlignment(2) = 4
    Me.grdUsuarios.ColAlignment(3) = 4
    Me.grdUsuarios.Rows = 1
    
    For varContador = 1 To varColSeguridad.Count
        Set varSeguridad = Nothing
        Set varSeguridad = New claSeguridad
        
        'La nueva variable asume la posición de la colección
        varSeguridad.proPrivilegios = varColSeguridad.Item(varContador).proPrivilegios
        varSeguridad.proUserId = varColSeguridad.Item(varContador).proUserId
        
        'Descompone los privilegios para pintar en la grilla
        varPermisoAdministracion = "N"
        If Mid(varSeguridad.proPrivilegios, 1, 1) = "1" Then
                varPermisoAdministracion = "S"
        End If
        
        varPermisoAdministracionTelefonia = "N"
        If Mid(varSeguridad.proPrivilegios, 2, 1) = "1" Then
                varPermisoAdministracionTelefonia = "S"
        End If
        
        varPermisoNoValidarProceso = "N"
        If Mid(varSeguridad.proPrivilegios, 3, 1) = "1" Then
            varPermisoNoValidarProceso = "S"
        End If
        
        Me.grdUsuarios.AddItem varSeguridad.proUserId & vbTab & _
                                            varPermisoAdministracion & vbTab & _
                                            varPermisoAdministracionTelefonia & vbTab & _
                                            varPermisoNoValidarProceso
    Next varContador
    
    Me.grdUsuarios.Row = 0
    
    Set varSeguridad = Nothing
    Me.grdUsuarios.Redraw = True
    FunFPintaUsuarios = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Private Sub cmdAgregar_Click()
Dim varSeguridad As claSeguridad
On Error GoTo ErrorManager

        Set varSeguridad = New claSeguridad
        Set varSeguridad.proConexion = Me.proConexion
        varSeguridad.proAplicacionId = AplicacionID
        
        Set frmEditarSeguridad.proSeguridad = varSeguridad
        Set frmEditarSeguridad.proColseguridad = varColSeguridad
        Set frmEditarSeguridad.proConexion = Me.proConexion
        frmEditarSeguridad.Show vbModal
        
        If Len(Trim(varSeguridad.proUserId)) <> 0 Then
                varColSeguridad.Add Me.proConexion, varSeguridad.proPrivilegios, varSeguridad.proAplicacionId, varSeguridad.proUserId
        End If
        
        'ReDespliega los usuarios y permisos
        FunFPintaUsuarios
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub


Private Sub cmdEliminar_Click()
On Error GoTo ErrorManager

        If Me.grdUsuarios.Row = 0 Then
            MsgBox "Debe seleccionar un usuario a eliminar", vbInformation, App.Title
            Exit Sub
        End If

        'ASegura que sea este usuario quien se va a eliminar
        If MsgBox("Está seguro de eliminar al usuario " & varColSeguridad.Item(Me.grdUsuarios.Row).proUserId & " ?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
        
        'ELimina el elemento en la base de datos
        If varColSeguridad.Item(Me.grdUsuarios.Row).FunGEliminar = False Then
                MsgBox "No fue posible eliminar al usuario", vbInformation, App.Title
                Exit Sub
        End If
        
        'ELimina el elemento de la colección
        varColSeguridad.Remove (Me.grdUsuarios.Row)
        
        'Despliega la colección
        FunFPintaUsuarios
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub cmdModificar_Click()
On Error GoTo ErrorManager

        If Me.grdUsuarios.Row = 0 Then
            MsgBox "Debe seleccionar un usuario a modificar", vbInformation, App.Title
            Exit Sub
        End If

        Set frmEditarSeguridad.proSeguridad = varColSeguridad.Item(Me.grdUsuarios.Row)
        Set frmEditarSeguridad.proColseguridad = varColSeguridad
        Set frmEditarSeguridad.proConexion = Me.proConexion
        frmEditarSeguridad.Show vbModal
        
        'ReDespliega los usuarios y permisos
        FunFPintaUsuarios
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorManager

        'No muestra la pantalla sino pudo cargar los datos de seguridad
        If varBandera = 1 Then
                Unload Me
        End If
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub Form_Load()
On Error GoTo ErrorManager

        varBandera = 0
        
        'Consulta los usuarios para la aplicación
        Set varColSeguridad = New colSeguridad
        Set varColSeguridad.proConexion = Me.proConexion
        
        varColSeguridad.proAplicacionId = AplicacionID
        
        If varColSeguridad.FunGConsulta = False Then
                MsgBox "No fue posible realizar la consulta de usuarios y permisos.", vbInformation, App.Title
                Exit Sub
        End If
        
        'Despliega los usuarios y sus permisos
        FunFPintaUsuarios
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub grdUsuarios_DblClick()
On Error GoTo ErrorManager


        If Me.grdUsuarios.Row <> 0 Then
                Call cmdModificar_Click
        End If
        Exit Sub
        
ErrorManager:
        SubGMuestraError
End Sub

Private Sub grdUsuarios_Scroll()
    On Error GoTo ErrorManager
                
        Me.grdTituloUsuarios.LeftCol = Me.grdUsuarios.LeftCol
        Exit Sub
                
ErrorManager:
        SubGMuestraError
End Sub


